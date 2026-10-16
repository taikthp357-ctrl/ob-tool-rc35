# ĐẶC TẢ KIẾN TRÚC & KỸ THUẬT: DỰ ÁN MA CONTROL — WEBAPP ĐIỀU HÀNH FG (AIM-1167)

- **Mã dự án / Issue cha**: `AIM-1167`
- **Kiến trúc sư trưởng / Điều phối**: Astra (`e5dfdb49-dd75-4a07-9522-4c4ab786e84b`)
- **Ngày ban hành**: 13/10/2026
- **Thư mục làm việc chuẩn**: `G:\AI-Manager\projects\ma-control\`
- **Tài liệu gốc tham chiếu**:
  - Brief nghiệp vụ: `G:\AI-Manager\work\manager\briefs\BRIEF_MA_CONTROL_FG_WEBAPP.md`
  - Kế hoạch nền tảng: `G:\AI-Manager\projects\ma-control\docs\MA_CONTROL_FG_PYTHON_LOCAL_PLAN.md`
  - Mã nguồn VBA RC52: `G:\AI-Manager\projects\ma-control\baseline\rc52\` (`ModuleMain.bas`, `ModuleSAP.bas`, `ModuleHelper.bas`, `ModuleLog.bas`)
  - Gói Baseline FG: `G:\AI-Manager\projects\ma-control\baseline\fg\extracted\` (`261008_BBBG_FG_V2_2_14.xlsm`, `261008_Stock_FG_MA_BODY_V3_6.xlsm`, `261002_MIGO_ATT_V3_4_3.xlsm`)

---

## 1. Mục tiêu Hệ thống & Nguyên tắc Bất biến (System Invariants)

Hợp nhất 3 công cụ Excel rời rạc (`Stock FG V3.6`, `OB Tool V1 RC52`, `BBBG FG V2.2.14`) thành một ứng dụng Webapp Local chạy trên Windows (Python FastAPI + SQLite + Excel-like Grid UI) theo luồng 5 bước chuẩn hóa:

```text
[1. Tải tồn] ➔ [2. Điền OB] ➔ [3. Ra đơn pick] ➔ [4. Người dùng CF] ➔ [5. BBBG]
```

### Các Invariants Bắt Buộc Bảo Toàn:
1. **Cấm tự động CF ngầm (Zero Silent Auto-CF)**:
   - Sau Bước 3 (`Ra đơn pick`), mọi đơn pick bắt buộc mang trạng thái `Chờ người dùng CF` (dòng OB mang trạng thái `Chờ pick / CF`).
   - Hệ thống **tuyệt đối không** tự động xác nhận CF ngầm. Người dùng phải kiểm tra thực tế và bấm nút **`Kiểm tra CF`** / xác nhận CF thủ công trên giao diện để chuyển trạng thái đơn pick sang `Đã CF`.
   - Bước 5 (`BBBG`) chặn cứng (reject HTTP 400) mọi yêu cầu tạo BBBG từ đơn pick chưa có trạng thái `Đã CF`.
2. **Kế thừa trọn vẹn cải tiến OB Tool RC52**:
   - **Nhánh `Body • 3001`**: Ưu tiên OB treo/draft, kiểm tra batch và số lượng thiếu; hỗ trợ quy trình xóa draft 3001 qua VL02N (`SAPX_DeleteDraftOB`: nhấn `btn[14]`, xác nhận `btnSPOP-OPTION1`, hậu kiểm live trả về `DELETE_VERIFIED`, `WAIT_OB_LOCK`, `WAIT_STO_LOCK`, `DELETE_FAILED`).
   - **Nhánh `EDC • 4007`**:
     - Gom nhóm STO chủ đạo (`Dominant STO Clustering` - `mEDCDominantSTO`) hoặc STO có capacity lớn nhất mà không hiện popup `InputBox`/`MsgBox` và không dính lỗi mảng/object.
     - Tách biệt lỗi khóa STO (`WAIT_STO_LOCK`: `"Chung tu bi khoa boi nguoi khac; vui long yeu cau nha khoa"`) khỏi lỗi thiếu tồn ATP; tuyệt đối không gọi xóa draft 3001 khi STO bị khóa.
     - Khôi phục lập kế hoạch (`OB_DELETED_REPLAN`) khi OB tạo thiếu đã bị người dùng xóa khỏi SAP (`capacity = 0` & `OB_DELETED_CONFIRMED`), gỡ bỏ bẫy `OB_ALREADY_CREATED_WAIT`.
     - Tối ưu tốc độ staging VL10D theo trang (`VerticalScrollbar.Position`), không cuộn lại từ đầu cho từng mã.
3. **Bảo toàn chuẩn Tồn FG V3.6 & BBBG FG V2.2.14**:
   - **Tồn kho**: Tách bạch theo `Plant + SLoc + Material + Batch + Stock Type + Unit`; không cộng trùng giữa các vùng `3093`, `3000`, `3001`. Giữ nguyên số `0` ở đầu mã hàng, batch, chứng từ.
   - **BBBG**: Không chèn dữ liệu vào giữa phiếu cũ (chỉ nối thêm cuối sổ tổng theo khóa chống trùng `doc_type + year + plant_shop + slip_no`). Hỗ trợ hàng ngoài hệ thống (bỏ qua check SAP). Mapping chuẩn model LIMO/VF7 (`BIW71001001`, `BIW70002022`). Đính kèm GOS (`BBBG` + `VAT`, bỏ qua VAT trùng tên).
4. **Kiến trúc Worker & Phục hồi**:
   - Cô lập SAP GUI Scripting COM/pywin32 trên 1 worker queue duy nhất (không chạy COM trực tiếp trên HTTP request thread).
   - Chống bấm lặp `Run All` bằng khóa idempotent (`business_key`). Hỗ trợ resume từ đúng dòng/bước bị gián đoạn; trạng thái `uncertain` khi crash sau thao tác ngoại vi (in/mail/GOS) phải đối chiếu trước khi chạy lại.

---

## 2. Cấu trúc Thư mục Dự án (`G:\AI-Manager\projects\ma-control\`)

```text
G:\AI-Manager\projects\ma-control\
├── app\
│   ├── __init__.py
│   ├── main.py                  # FastAPI entrypoint, mount static & templates, routers
│   ├── api\
│   │   ├── __init__.py
│   │   ├── routes_workflow.py   # API 5 bước: stock, ob, pick-orders, cf, bbbg, jobs
│   │   └── routes_config.py     # API cấu hình, model mapping, lịch sử, healthcheck
│   ├── core\
│   │   ├── __init__.py
│   │   ├── config.py            # Cấu hình ứng dụng, đường dẫn DB, chế độ SAP (mock/live)
│   │   └── database.py          # SQLite connection, WAL mode, khởi tạo schema & seed data
│   ├── models\
│   │   ├── __init__.py
│   │   └── schemas.py           # Pydantic models cho request/response 5 bước
│   ├── services\
│   │   ├── __init__.py
│   │   ├── stock_service.py     # Bước 1: Tải tồn, phân vùng 3093/3000/3001, cảnh báo & ưu tiên
│   │   ├── ob_service.py        # Bước 2: Kiểm tra & Điền OB (Body 3001 & EDC 4007 RC52 logic)
│   │   ├── pick_cf_service.py   # Bước 3 & 4: Sinh đơn pick & Xác nhận CF thủ công bắt buộc
│   │   ├── bbbg_service.py      # Bước 5: Tạo BBBG, in, mail, xuất sổ tổng/shop, GOS BBBG/VAT
│   │   └── sap_worker.py        # Single-worker Job Queue, idempotent lock, resume & audit log
│   ├── static\
│   │   ├── css\
│   │   │   └── styles.css       # Giao diện chuẩn mockup Excel-like Grid
│   │   └── js\
│   │       └── app.js           # Logic Grid (Ctrl+V, Tab/Enter, inline edit, 5-step flow)
│   └── templates\
│       └── index.html           # Giao diện chính MA CONTROL — Điều hành FG
├── data\
│   └── ma_control_fg.db         # SQLite local database (tự động khởi tạo)
├── docs\
│   ├── MA_CONTROL_FG_PYTHON_LOCAL_PLAN.md
│   └── ARCHITECTURE_SPEC_AIM1167.md
├── tests\
│   ├── test_schema_and_stock.py
│   ├── test_ob_rc52_logic.py
│   ├── test_pick_and_manual_cf.py
│   ├── test_bbbg_and_recovery.py
│   └── test_ui_and_api_e2e.py
└── requirements.txt
```

---

## 3. Thiết kế Cơ sở Dữ liệu SQLite (`data/ma_control_fg.db`)

1. **`materials`**:
   - `material_code TEXT PRIMARY KEY`, `description TEXT`, `model TEXT`, `shop TEXT`, `standard_pack REAL`, `default_sloc TEXT`, `is_external INTEGER DEFAULT 0`, `updated_at TEXT`
2. **`model_mappings`**:
   - `material_code TEXT PRIMARY KEY`, `model_code TEXT NOT NULL`, `shop_code TEXT`, `note TEXT` (Seed sẵn `BIW71001001`, `BIW70002022` cho LIMO/VF7).
3. **`stock_snapshots`**:
   - `id INTEGER PRIMARY KEY AUTOINCREMENT`, `source_type TEXT`, `snapshot_time TEXT`, `plant TEXT`, `sloc TEXT`, `total_rows INTEGER`, `notes TEXT`
4. **`stock_items`**:
   - `id INTEGER PRIMARY KEY AUTOINCREMENT`, `snapshot_id INTEGER`, `plant TEXT`, `sloc TEXT`, `material_code TEXT`, `batch TEXT`, `stock_type TEXT`, `unrestricted_qty REAL`, `unit TEXT`, `priority_rank INTEGER`, `pull_line_warning TEXT`
5. **`ob_items`**:
   - `id INTEGER PRIMARY KEY AUTOINCREMENT`, `branch_tab TEXT NOT NULL` (`'BODY_3001'` hoặc `'EDC_4007'`), `row_order INTEGER`, `selected INTEGER DEFAULT 1`, `material_code TEXT`, `model TEXT`, `req_qty REAL`, `avail_qty REAL`, `missing_qty REAL`, `batch TEXT`, `shop TEXT`, `sto_number TEXT`, `ob_number TEXT`, `status TEXT DEFAULT 'Chưa kiểm tra'`, `rc52_code TEXT`, `note TEXT`, `pick_order_id INTEGER`, `updated_at TEXT`
6. **`pick_orders`**:
   - `id INTEGER PRIMARY KEY AUTOINCREMENT`, `pick_order_code TEXT UNIQUE`, `branch_tab TEXT`, `shop TEXT`, `model TEXT`, `item_count INTEGER`, `total_qty REAL`, `status TEXT DEFAULT 'Chờ người dùng CF'`, `cf_verified_by TEXT`, `cf_verified_at TEXT`, `created_at TEXT`, `updated_at TEXT`
7. **`pick_order_items`**:
   - `id INTEGER PRIMARY KEY AUTOINCREMENT`, `pick_order_id INTEGER`, `ob_item_id INTEGER`, `material_code TEXT`, `model TEXT`, `batch TEXT`, `shop TEXT`, `sto_number TEXT`, `ob_number TEXT`, `qty REAL`, `cf_status TEXT DEFAULT 'PENDING_CF'`, `cf_actual_qty REAL`
8. **`bbbg_records`**:
   - `id INTEGER PRIMARY KEY AUTOINCREMENT`, `bbbg_uuid TEXT UNIQUE`, `dedup_key TEXT UNIQUE`, `pick_order_id INTEGER`, `slip_no TEXT`, `doc_no TEXT`, `ob_number TEXT`, `shop TEXT`, `model TEXT`, `handover_person TEXT`, `receiver_person TEXT`, `storage_location_note TEXT`, `is_external INTEGER DEFAULT 0`, `template_mode TEXT DEFAULT 'Auto'`, `font_size INTEGER DEFAULT 14`, `email_status TEXT`, `print_status TEXT`, `unc_export_status TEXT`, `gos_bbbg_attached INTEGER DEFAULT 0`, `gos_vat_attached INTEGER DEFAULT 0`, `status TEXT`, `created_at TEXT`
9. **`job_events`**:
   - `id INTEGER PRIMARY KEY AUTOINCREMENT`, `job_id TEXT`, `business_key TEXT`, `job_type TEXT`, `step_name TEXT`, `status TEXT`, `duration_ms INTEGER`, `sap_status_msg TEXT`, `evidence_json TEXT`, `created_at TEXT`
10. **`settings`**:
    - `key TEXT PRIMARY KEY`, `value_json TEXT`, `version INTEGER DEFAULT 1`, `updated_at TEXT`

---

## 4. Đặc tả Giao diện Người dùng (Excel-like Grid UI)

Giao diện tại `http://127.0.0.1:8000` phải khớp 100% bố cục và thành phần theo thiết kế mockup:
- **Sidebar trái**: `Điều hành FG`, `Tồn & ưu tiên`, `Điền OB / Đơn pick` (active), `Chờ CF`, `BBBG`, `Lịch sử`, `Cấu hình`.
- **Header**: Tiêu đề `MA CONTROL — Điều hành FG`, badge `SAP đã kết nối` (màu xanh lá), badge phiên bản `RC52 • FG`, thông tin ca/người dùng.
- **Thanh tiến trình 5 bước (Stepper)**:
  - `1. Tải tồn` (tick xanh hoàn tất) ➔ `2. Điền OB` (active xanh dương) ➔ `3. Ra đơn pick` ➔ `4. Người dùng CF` ➔ `5. BBBG`.
- **Toolbar & Phân nhánh**:
  - 2 Tabs chuyển nhánh: `Body • 3001` và `EDC • 4007`.
  - Nhóm nút nạp dữ liệu: `Lấy từ tồn`, `Dán từ Excel`, `+ Thêm dòng`.
  - Bộ lọc: Ô tìm kiếm `Tìm mã / OB`, Dropdown `Tất cả model`, Dropdown `Tất cả trạng thái`.
  - Nhóm nút hành động nghiệp vụ: `Kiểm tra` (outline), `Điền OB` (primary), `Ra đơn pick` (primary).
- **Lưới dữ liệu trung tâm (Excel-like Data Grid)**:
  - Dòng tiêu đề chữ cái cột Excel: `A, B, C, D, E, F, G, H, I` và cột số thứ tự dòng `1, 2, 3...`.
  - Checkbox chọn từng dòng và chọn tất cả.
  - Các cột: `Mã hàng`, `Model`, `SL yêu cầu`, `Batch`, `Shop`, `STO`, `OB`, `Trạng thái`.
  - **Tương tác chuẩn Excel**:
    - Dán `Ctrl+V` đa ô / đa dòng từ clipboard Excel (tách `\r\n` và `\t`, tự động điền vào các ô hoặc thêm dòng mới mà không vỡ định dạng, giữ số `0` đầu).
    - Di chuyển ô bằng phím `Tab`, `Enter`, phím mũi tên; sửa trực tiếp trên ô (inline editing).
    - Ô đang chọn có viền xanh lá mỏng kiểu Excel (`#107c41` / `#16a34a`).
  - **Chip trạng thái**:
    - `Sẵn sàng điền` (xanh lá)
    - `Chưa kiểm tra` (xám)
    - `Thiếu SL OB` (vàng cam amber kèm icon cảnh báo `⚠`)
    - `Chờ pick / CF` (xanh dương)
  - **Footer bảng**: `Đã chọn X mã • Tổng SL Y PCS` bên trái; hướng dẫn `Ctrl+V để dán nhiều ô • Tab / Enter để di chuyển` bên phải.
- **Bảng "Đơn pick đã tạo"** (nằm ngay dưới lưới chính):
  - Các cột: `STT`, `Số đơn pick` (VD: `PICK-20261013-001`), `Ngày tạo`, `Shop`, `Model`, `Số mã hàng`, `Tổng SL (PCS)`, `Trạng thái` (`Chờ người dùng CF` / `Đã CF`), `Thao tác` (nút `Kiểm tra CF`, menu `...`).
  - Ghi chú cảnh báo rõ: `Sau khi CF mới chuyển sang BBBG` và nút điều hướng `Mở danh sách chờ CF`.
- **Các màn hình / panel chức năng liên thông**:
  - Cho phép thao tác đầy đủ cả 5 bước: nạp tồn & xem cảnh báo (Bước 1), kiểm tra & điền OB (Bước 2), sinh đơn pick (Bước 3), bấm `Kiểm tra CF` xác nhận thủ công (Bước 4), và tạo/xuất BBBG + đính kèm GOS (Bước 5).
