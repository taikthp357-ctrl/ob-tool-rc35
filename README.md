# OB TOOL V1 RC60

Canonical XLSM: `261010_OB_TOOL_V1_RC60.xlsm`  
Canonical SHA256: `FF86BF33B8CB638755E479B7071C680E36639A25EB840688ABB6E72CD92A83FF`  
Handoff Archive: `RC60_HANDOFF.zip`  
Archive SHA256: `8EA43D6F7A4E835D5DD6EE8A122321762EC9777CF3D35DA52E0E587B9DBA2A15`

## Key Improvements in RC60
- **P0 — Khử bỏ kiểm tra SLoc dư thừa trong VL02N (`ModuleSAP.bas`)**:
  * Dựa trên thực tế vận hành SAP: Khi kiểm tra tại `VL06O` thấy `rowSloc = ""` thì toàn bộ các dòng hàng trong OB đó đều không có SLoc (không tồn tại tình trạng dòng có dòng không).
  * Trong `SAPX_DeleteDraftOB`, khi giải phóng draft OB 3001 (`isDraft3001Release = True`), hệ thống nhảy thẳng sang lệnh xóa toàn bộ chứng từ qua `btn[14]`, khử bỏ hoàn toàn việc cuộn duyệt bảng `OverviewSnapshot` tốn thời gian.
- **P0 — Báo Thành Công Ngay Sau Khi Xóa — Khử Bỏ Tra Cứu Lặp Lại (Zero Re-query)**:
  * Sau khi bấm `btn[14]` và xác nhận pop-up xóa `btnSPOP-OPTION1`, hệ thống kiểm tra trực tiếp trạng thái thanh công cụ SAP (`sType <> "E"`). Báo thành công (`DELETE_VERIFIED`) ngay lập tức mà không cần thoát ra gõ lại `/nVL02N` để tra cứu lần 2.
- **P0 — Gom Các Bước Xử Lý Hàng Loạt Cho Plant 3001 (Batch Consolidation)**:
  * Trong `SAPX_FindAndReleaseDraft3001OB_Batch`: Quét toàn bộ ALV `VL06O` 1 lần duy nhất cho tất cả các mã thiếu, gom toàn bộ draft OB không có SLoc và chưa xử lý (`rowBatch = ""`, `rowSloc = ""`) khớp với danh sách thiếu hoặc `targetSTO`.
  * Mở `VL02N` 1 lần và xóa tuần tự siêu tốc toàn bộ danh sách chứng từ gom được. Đợi 2s đồng bộ tồn ATP 1 lần duy nhất sau khi hoàn tất cả batch.
- **Biên dịch & Bảo toàn Hệ thống**:
  * Native VBA Debug Compile PASS 100% (`before=True`, `after=False`).
  * Gắn macro nút bấm chuẩn xác (`INPUT!OBV1__0` -> `OB_Check`, `INPUT!OBV1__1` -> `OB_Execute`).
  * Khử bỏ triệt để `forceFullCalc="1"` trong `xl/workbook.xml`.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC60 (14/10/2026)

### 📦 Các tệp đính kèm RC60

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC60_HANDOFF.zip`** | 892,447 bytes | `8EA43D6F7A4E835D5DD6EE8A122321762EC9777CF3D35DA52E0E587B9DBA2A15` | **Gói bàn giao đầy đủ RC60 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC60.xlsm`** | 911,151 bytes | `FF86BF33B8CB638755E479B7071C680E36639A25EB840688ABB6E72CD92A83FF` | Tệp Excel Macro phát hành RC60 (Đã tối ưu & Compile) |
| **`VBA_RC59_TO_RC60.diff`** | 9,833 bytes | `71BA447E4AE8DEC8EC9147D896C2ADB2EC1576BE3164DC7DD13082DC8D38A79F` | Chi tiết bản vá mã nguồn VBA từ RC59 lên RC60 |
| **`CHANGELOG_RC59_TO_RC60.md`** | 4,333 bytes | `0C1D5BE4002E0CB3E8F3CD2EC375BBCB4F0E4EDA845C8B816A347DEF2EC7946E` | Nhật ký thay đổi và giải pháp kỹ thuật RC60 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC59 (14/10/2026)
* **`RC59_HANDOFF.zip`**: SHA-256 `27E645BAD57A722A4E49B94ED358E9431764B42DB9BA6F83D35FEE8015E5AC22`
* **`261010_OB_TOOL_V1_RC59.xlsm`**: SHA-256 `1B80D0AD1223B8736A66D4A487DB0FD55C0D3A66CADE08BA09B8587E9D8542A7`

### OB Tool V1 RC58 (14/10/2026)
* **`RC58_HANDOFF.zip`**: SHA-256 `A706786D3E846F711C457C50B35914C892158239F0CC9D3AA1C60C4EA1DE315E`
* **`261010_OB_TOOL_V1_RC58.xlsm`**: SHA-256 `298C458259268CC602EBEA890840C5877F4F6F397B5157C68E9D675DF04C4973`
