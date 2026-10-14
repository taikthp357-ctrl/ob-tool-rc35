# OB TOOL V1 RC61

Canonical XLSM: `261010_OB_TOOL_V1_RC61.xlsm`  
Canonical SHA256: `E446C065F22FD3D25CD14318AA54CB66C7599C7E6637650DE2C2A974C1DFDF5D`  
Handoff Archive: `RC61_HANDOFF.zip`  
Archive SHA256: `0DB0D4437D316843DC3DE7CFA517DD2ADE8432631B0E98CD78504A13519180B0`

## Key Improvements in RC61
- **P0 — Vá triệt để lỗi `GROUP_RECEIPT_MATERIAL_MISMATCH` khi Auto-Retry (`ModuleMain.bas`)**:
  * Khi kích hoạt `AUTO_RETRY_VL10D` tạo bù các mã thiếu (deficit), hệ thống bảo toàn danh sách chứng từ Lượt 1 vào `priorVerifiedReceipts` và xóa trắng `VL10D_GROUP_OB_RECEIPTS`.
  * Khử bỏ tình trạng cache chứng từ Lượt 1 gây lỗi runtime `-2147217235` khi xác minh chứng từ ở Lượt Retry.
- **P0 — Phòng thủ đa tầng trong xác minh biên nhận (Defensive Receipt Filtering)**:
  * Trong `EDCVerifyGroupReceipts`, nếu gặp dòng chứng từ thuộc mã không nằm trong phạm vi kế hoạch retry (`Not plans.exists`), hệ thống tự động bỏ qua (skip) thay vì ném ngoại lệ làm sập luồng sang `Failed:`.
- **P0 — Cách ly Candidate Discovery khỏi OB cũ Lượt 1**:
  * Trích xuất các số OB đã sinh ở Lượt 1 nạp vào danh sách loại trừ baseline (`rPlanned(8)`), đảm bảo `EDCDiscoverReceiptsFromVL06O` chỉ nhận diện và bắt chính xác OB mới sinh của Lượt Retry.
- **P0 — Hợp nhất chứng từ & Cộng dồn số lượng giữa các lượt**:
  * Sau khi retry hoàn tất, gộp danh sách chứng từ cả 2 lượt vào `VL10D_GROUP_OB_RECEIPTS`.
  * Cộng dồn số lượng thực tế giữa Lượt 1 và Lượt Retry (`vPlan(3) = vPlan(3) + rPlan(3)`). Khi đủ 100% nhu cầu, tự động xóa ghi chú thiếu trên Cột D sheet `INPUT`.
- **Biên dịch & Bảo toàn Hệ thống**:
  * Native VBA Debug Compile PASS 100% (`before=True`, `after=False`).
  * Gắn macro nút bấm chuẩn xác (`INPUT!OBV1__0` -> `OB_Check`, `INPUT!OBV1__1` -> `OB_Execute`).
  * Khử bỏ triệt để `forceFullCalc="1"` trong `xl/workbook.xml`.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC61 (15/10/2026)

### 📦 Các tệp đính kèm RC61

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC61_HANDOFF.zip`** | 897,131 bytes | `0DB0D4437D316843DC3DE7CFA517DD2ADE8432631B0E98CD78504A13519180B0` | **Gói bàn giao đầy đủ RC61 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC61.xlsm`** | 916,221 bytes | `E446C065F22FD3D25CD14318AA54CB66C7599C7E6637650DE2C2A974C1DFDF5D` | Tệp Excel Macro phát hành RC61 (Đã tối ưu & Compile) |
| **`VBA_RC60_TO_RC61.diff`** | 5,595 bytes | `536DEB189867DB52F7F48D96C1BD1BFCBBEEDEF2328F3A722E9B905FF3B556ED` | Chi tiết bản vá mã nguồn VBA từ RC60 lên RC61 |
| **`CHANGELOG_RC60_TO_RC61.md`** | 4,643 bytes | `E592533AF92E977C0BE21E12E8F8CF3C34A2F521816D40A7D2A6C1CC607DDBCC` | Nhật ký thay đổi và giải pháp kỹ thuật RC61 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC60 (14/10/2026)
* **`RC60_HANDOFF.zip`**: SHA-256 `8EA43D6F7A4E835D5DD6EE8A122321762EC9777CF3D35DA52E0E587B9DBA2A15`
* **`261010_OB_TOOL_V1_RC60.xlsm`**: SHA-256 `FF86BF33B8CB638755E479B7071C680E36639A25EB840688ABB6E72CD92A83FF`

### OB Tool V1 RC59 (14/10/2026)
* **`RC59_HANDOFF.zip`**: SHA-256 `27E645BAD57A722A4E49B94ED358E9431764B42DB9BA6F83D35FEE8015E5AC22`
* **`261010_OB_TOOL_V1_RC59.xlsm`**: SHA-256 `1B80D0AD1223B8736A66D4A487DB0FD55C0D3A66CADE08BA09B8587E9D8542A7`
