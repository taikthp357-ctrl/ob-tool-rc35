# BÁO CÁO ĐIỀU PHỐI & PHÂN RÃ NHIỆM VỤ: DỰ ÁN MA CONTROL — ĐIỀU HÀNH FG (AIM-1167)

- **Người lập**: Astra — Giám đốc điều hành & Kiến trúc sư trưởng (`e5dfdb49-dd75-4a07-9522-4c4ab786e84b`)
- **Mã công việc cha**: `AIM-1167`
- **Thư mục làm việc**: `G:\AI-Manager\projects\ma-control\`
- **Tài liệu đặc tả kiến trúc**: `G:\AI-Manager\projects\ma-control\docs\ARCHITECTURE_SPEC_AIM1167.md`
- **Trạng thái nhiệm vụ cha**: `in_progress` (Đang giám sát các tác nhân chuyên trách thực thi)

---

### 1. Phân Tích Kỹ Thuật & Kiến Trúc Hệ Thống Hợp Nhất

Astra đã hoàn tất rà soát hồ sơ brief, mã nguồn baseline VBA RC52 và gói Excel V2.2.14. Dự án xây dựng một Webapp Local chuẩn chạy trên Windows hợp nhất bộ 3 công cụ:
1. **`Stock FG V3.6`**: Quản lý tồn kho theo snapshot, phân tách rõ Plant/SLoc 3093/3000/3001 (không cộng trùng), xác định mã ưu tiên xuất và cảnh báo kéo line.
2. **`OB Tool V1 RC52`**: Kế thừa toàn bộ logic SAP cốt lõi vừa phát hành:
   - Nhánh `Body • 3001`: Ưu tiên OB treo, xóa draft qua `VL02N` (`deleteob.vbs`).
   - Nhánh `EDC • 4007`: Gom nhóm STO chủ đạo (`Dominant STO Clustering`) không bị lỗi array object 424, không văng popup modal; tách biệt lỗi `WAIT_STO_LOCK` (chứng từ bị khóa bởi người khác, tuyệt đối không hủy draft 3001); khôi phục lập kế hoạch `OB_DELETED_REPLAN` khi OB thiếu bị xóa trên SAP; tối ưu staging VL10D theo trang.
3. **`BBBG FG V2.2.14`**: Tạo số phiếu, đối soát, ghi sổ tổng append-only không chèn giữa, đính kèm GOS BBBG/VAT (bỏ qua VAT trùng), hỗ trợ hàng ngoài hệ thống.

#### Luồng Nghiệp Vụ Chuẩn Hóa 5 Bước (Bảo Toàn Invariants):
```text
[1. Tải tồn] ➔ [2. Điền OB] ➔ [3. Ra đơn pick] ➔ [4. Người dùng CF] ➔ [5. BBBG]
```
> ⛔ **INVARIANT CỐT LÕI BƯỚC 4**: Sau Bước 3 (Ra đơn pick), đơn pick sinh mã (ví dụ `PICK-YYYYMMDD-XXX`) và mang trạng thái `Chờ người dùng CF`. **TUYỆT ĐỐI KHÔNG TỰ ĐỘNG CF NGẦM**. Người dùng phải kiểm tra thực tế và bấm nút **`Kiểm tra CF`** trên UI mới chuyển sang `Đã CF`. Bước 5 (BBBG) chặn cứng mọi yêu cầu khi đơn pick chưa được CF hợp lệ.

#### Giao Diện Chuẩn Mockup (Excel-like Grid UI):
- Lưới trung tâm có dòng tiêu đề chữ cái cột Excel (`A, B, C, D...`), số thứ tự dòng (`1, 2, 3...`), phím tắt `Ctrl+V` paste đa ô/đa dòng từ Excel không mất số `0` đầu, phím `Tab`/`Enter` điều hướng, inline edit, viền chọn xanh lá mỏng, chip trạng thái trực quan.
- Bảng "Đơn pick đã tạo" nằm dưới lưới: hiển thị danh sách đơn pick, trạng thái `Chờ người dùng CF`, nút `Kiểm tra CF`, và liên kết điều hướng.

---

### 2. Ma Trận Ủy Quyền & Chuỗi Nhiệm Vụ Phân Rã (Child Issues)

Astra đã khởi tạo thành công 3 child issues với chuỗi phụ thuộc (dependency blocking chain) chặt chẽ:

| Mã Task | Tác nhân phụ trách | Vai trò | Trạng thái | Điều kiện kích hoạt (Blocker) |
|---|---|---|---|---|
| **AIM-1168** | **Kai** (`7878f735-db2f-4842-a65b-e4c9a2fb2bc4`) | Lập trình phần mềm (Senior Dev) | `in_progress` | Kích hoạt ngay |
| **AIM-1169** | **Luna** (`e229aed5-811d-4327-acae-9ac496f7ca87`) | Kiểm toán chất lượng (Code Review) | `todo` (Blocked) | Chờ Kai hoàn thành AIM-1168 |
| **AIM-1170** | **Nova** (`3d3b8059-60e8-44dd-a664-bbaf294b0bc0`) | Kiểm thử độc lập (QA Engineer) | `todo` (Blocked) | Chờ Luna hoàn thành AIM-1169 |

#### Chi tiết Giao việc:
1. **Kai (AIM-1168 - Dev)**:
   - Khởi tạo SQLite DB `data/ma_control_fg.db` với 10 bảng chuẩn, WAL mode, bảo toàn số 0 đầu.
   - Xây dựng FastAPI backend phục vụ 5 bước nghiệp vụ và single-worker SAP queue.
   - Hiện thực hóa Excel-like Grid UI (`index.html`, `styles.css`, `app.js`) khớp 100% mockup thiết kế, hỗ trợ `Ctrl+V`, `Tab`, `Enter`, chip trạng thái, và bảng đơn pick chờ CF.
   - Đảm bảo trọn vẹn logic RC52 và cấm tuyệt đối auto-CF ngầm.
   - Viết unit & integration tests trong `tests/`.
2. **Luna (AIM-1169 - Review)**:
   - Rà soát toàn diện mã nguồn theo `ARCHITECTURE_SPEC_AIM1167.md`.
   - Đối soát nghiêm ngặt invariant không auto-CF ngầm tại Bước 4.
   - Đối soát logic VBA RC52 (xóa draft 3001, tách `WAIT_STO_LOCK`, dominant STO, `OB_DELETED_REPLAN`).
   - Kiểm tra tính toàn vẹn dữ liệu SQLite và quy tắc append-only của sổ tổng BBBG.
3. **Nova (AIM-1170 - QA)**:
   - Kiểm thử E2E luồng 5 bước: xác nhận Bước 4 chặn tạo BBBG khi chưa CF, và thông luồng khi bấm "Kiểm tra CF".
   - Kiểm thử tương tác Excel Grid: paste `Ctrl+V` đa dòng, phím Tab/Enter, inline edit.
   - Chạy toàn bộ test suite và lập bảng kết quả kiểm thử.

---

### 3. Kế Hoạch Giám Sát & Nghiệm Thu

- **Astra** tiếp tục theo dõi tiến độ thực thi của Kai, Luna, và Nova.
- Sau khi Nova hoàn tất kiểm thử và xuất báo cáo QA Acceptance PASS 100%, Astra sẽ thực hiện tổng kết, cập nhật biên bản nghiệm thu và đóng task cha AIM-1167.
