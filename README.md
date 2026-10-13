# OB TOOL V1 RC56

Canonical XLSM: `261010_OB_TOOL_V1_RC56.xlsm`  
Canonical SHA256: `8F36ECFBC4E6527822E4A69765EECB6CBB6B1AB3676F0DEC6CCCB78AAFB828F8`  
Handoff Archive: `RC56_HANDOFF.zip`  
Archive SHA256: `3914E6C6285B769D13704D4922132610C23A926C8B6E001F99BC1ABB4F9D7ABA`

## Key Improvements in RC56
- **P0 — Vá lỗi Schema Deserializer `VL10D_PENDING_KEY` 6 trường (`ModuleMain.bas`)**:
  - Chuẩn hóa schema pending key 6 trường: `Request | Material | Qty | STO | Item | Timestamp`.
  - Nâng cấp `EDCResolvePendingCreate` và `EDCResolvePendingGroup` (`UBound(parts) >= 4`) chấp nhận cả định dạng 5 trường cũ và 6 trường mới.
  - Bảo vệ pending key trong khối `Waiting:` không bị xóa nhầm do lỗi giao diện SAP tạm thời.
- **P0 — Bảo toàn Pending Key trong quy trình `AUTO_RETRY_VL10D` (`ModuleMain.bas`)**:
  - `EDCCommitGroup` hỗ trợ `keepPendingKey = True` khi có đợt retry sắp chạy, ngăn chặn việc xóa pending sớm.
  - Tái tạo trạng thái `VL10D_GROUP_PENDING` và `VL10D_PENDING_KEY` cho các mã thiếu trước khi gọi SAP Create trong retry.
  - Cơ chế fallback nhận diện STO/Request từ danh sách `plans` trong `EDCVerifyGroup` chống lỗi `GROUP_RECEIPT_MISSING`.
- **P0 — Nhánh Xử Lý Tạo 1 Nửa STO & Tự Động Xóa OB Thiếu (`ModuleMain.bas`)**:
  - Khi VL10D chỉ tạo được một phần số lượng và nhánh 3001 không giữ tồn hoặc retry không thể tạo thêm:
  - Tự động kiểm tra nếu OB chỉ chứa mã bị thiếu (hoặc không chứa mã khác đã hoàn thành đủ), gọi `ModuleSAP.SAPX_DeleteDraftOB` per `deleteob.vbs` để xóa OB dở dang và hoàn trả Due về STO.
  - Cập nhật ghi chú trên INPUT: `DA XOA OB THIEU (TAO .../... PCS): HET TON & 3001 KO GIU TON; TRA LAI STO DUE`.
- **P1 — Gỡ bỏ bẫy `PRESAVE_GATE_FAIL: EXTRA_ITEM` trong VL02N (`ModuleSAP.bas`)**:
  - Trong `VerifyDeliveryPlan` và `SavedSnapshotMatchesPlan`: Cho phép lưu OB khi trên chứng từ có item khác chưa có kế hoạch trong lượt chạy hiện tại; ghi log `PRESAVE_EXTRA_ITEM_INFO` thay vì chặn lưu.
- **P1 — Khắc phục lỗi đọc `TableField` trong VL02N (`ModuleSAP.bas`)**:
  - `TableField` trả về `Nothing` an toàn khi đọc tràn dòng thay vì ném Runtime Error 4205.
  - `OverviewSnapshot` và `VisibleItemRow` kiểm tra `fld Is Nothing` an toàn, ngăn ngừa hoàn toàn Runtime Error 91 và 4205.
- **P1 — Đối chiếu an toàn Draft 3001 & Mở rộng phạm vi ngày Planned GI (`ModuleSAP.bas`, `ModuleMain.bas`)**:
  - `SAPX_FindAndReleaseDraft3001OB` đọc cột STO và Creator; tính điểm ưu tiên xóa đúng chứng từ do người dùng hiện tại tạo (+1000) và đúng STO mục tiêu (+500).
  - Mở rộng `IT_WADAT-HIGH` trong VL06O thành `DateAdd("d", 60, Date)` (+60 ngày).
  - `EDCPlanOne` ưu tiên chọn STO có Due đáp ứng trọn vẹn nguyên batch (`EDCWholeBatchFitQty > 0`); tiếp tục lập kế hoạch cho phần thiếu khi STO có thêm Due.
  - `PrepareStock` đồng bộ kiểm tra loại kho qua `POST_ALLOWED_STOCK_TYPES`.
- **Biên dịch & Bảo toàn Hệ thống**:
  - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
  - Gắn macro nút bấm chuẩn xác (`INPUT!OBV1__0` -> `OB_Check`, `INPUT!OBV1__1` -> `OB_Execute`).
  - Khử bỏ triệt để `forceFullCalc="1"` trong `xl/workbook.xml`.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC56 (14/10/2026)

### 📦 Các tệp đính kèm RC56

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC56_HANDOFF.zip`** | 731,775 bytes | `3914E6C6285B769D13704D4922132610C23A926C8B6E001F99BC1ABB4F9D7ABA` | **Gói bàn giao đầy đủ RC56 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC56.xlsm`** | 795,147 bytes | `8F36ECFBC4E6527822E4A69765EECB6CBB6B1AB3676F0DEC6CCCB78AAFB828F8` | Tệp Excel Macro phát hành RC56 (Đã tối ưu & Compile) |
| **`VBA_RC55_TO_RC56.diff`** | 28,628 bytes | `B814343F783336357C07E618CEFED573DC64BCFBF9DC8B580A6FA41AB6117C10` | Chi tiết bản vá mã nguồn VBA từ RC55 lên RC56 |
| **`CHANGELOG_RC55_TO_RC56.md`** | 7,550 bytes | `CFCD2632BD7F1F2391AD5A65029A9349D4B99348B4520A4FCAD848791244D3FB` | Nhật ký thay đổi và giải pháp kỹ thuật RC56 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC55 (14/10/2026)
* **`261010_OB_TOOL_V1_RC55.xlsm`**: SHA-256 `547A36809C42263FCA5050C6CA162E089A0C3B77D0A4729FF3A5EEE2300C99AF`

### OB Tool V1 RC52 (13/10/2026)
* **`RC52_HANDOFF.zip`**: SHA-256 `E6FA4A895F2CE2C281F2AECA11FA4137F2F3B6CA27C3BA7E7DDE6D17EB73C27C`
* **`261010_OB_TOOL_V1_RC52.xlsm`**: SHA-256 `C506832470E005CFF848A2F9319FA9266DC3136B0F5D2683053DFD1025B99979`
