# OB TOOL V1 RC59

Canonical XLSM: `261010_OB_TOOL_V1_RC59.xlsm`  
Canonical SHA256: `1B80D0AD1223B8736A66D4A487DB0FD55C0D3A66CADE08BA09B8587E9D8542A7`  
Handoff Archive: `RC59_HANDOFF.zip`  
Archive SHA256: `27E645BAD57A722A4E49B94ED358E9431764B42DB9BA6F83D35FEE8015E5AC22`

## Key Improvements in RC59
- **P0 — Phân định rõ Thông điệp Ghi chú Cột D Sheet INPUT (`ModuleMain.bas`)**:
  * Khi xóa cả chứng từ đơn mã: `"DA XOA OB THIEU (TAO .../... PCS) DE TRA STO DUE"`.
  * Khi chỉ xóa dòng trên chứng từ đa mã: `"DA XOA DONG ITEM THIEU (... PCS) TREN OB ...; GIU NGUYEN CAC DONG DA DU"`.
- **P0 — Nhánh Xóa Toàn Bộ Draft OB 3001 Chưa Có SLoc & Chưa Xử Lý (`ModuleSAP.bas`)**:
  * Khi giải phóng tồn từ draft OB nhánh 3001, nếu OB chứa nhiều mã nhưng toàn bộ các dòng đều chưa có SLoc (`sloc = ""`) và chưa ai xử lý (unpicked / `batch = ""`), tự động xóa TOÀN BỘ chứng từ qua `btn[14]` thay vì xóa lẻ từng dòng.
  * Giúp khôi phục ngay lập tức 100% dung lượng tồn và Open Quantity trên STO cho tất cả các mã hàng.
- **Biên dịch & Bảo toàn Hệ thống**:
  * Native VBA Debug Compile PASS 100% (`before=True`, `after=False`).
  * Gắn macro nút bấm chuẩn xác (`INPUT!OBV1__0` -> `OB_Check`, `INPUT!OBV1__1` -> `OB_Execute`).
  * Khử bỏ triệt để `forceFullCalc="1"` trong `xl/workbook.xml`.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC59 (14/10/2026)

### 📦 Các tệp đính kèm RC59

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC59_HANDOFF.zip`** | 889,600 bytes | `27E645BAD57A722A4E49B94ED358E9431764B42DB9BA6F83D35FEE8015E5AC22` | **Gói bàn giao đầy đủ RC59 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC59.xlsm`** | 910,021 bytes | `1B80D0AD1223B8736A66D4A487DB0FD55C0D3A66CADE08BA09B8587E9D8542A7` | Tệp Excel Macro phát hành RC59 (Đã tối ưu & Compile) |
| **`VBA_RC58_TO_RC59.diff`** | 5,405 bytes | `7D27E5C16C07A65987E334ED082E252A681D2482B32251759FF06BA01704065F` | Chi tiết bản vá mã nguồn VBA từ RC58 lên RC59 |
| **`CHANGELOG_RC58_TO_RC59.md`** | 3,013 bytes | `8CE2423D203BE3BF06C8F7FFAA81961E6498A35380367AB5FF3B3842148C9582` | Nhật ký thay đổi và giải pháp kỹ thuật RC59 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC58 (14/10/2026)
* **`RC58_HANDOFF.zip`**: SHA-256 `A706786D3E846F711C457C50B35914C892158239F0CC9D3AA1C60C4EA1DE315E`
* **`261010_OB_TOOL_V1_RC58.xlsm`**: SHA-256 `298C458259268CC602EBEA890840C5877F4F6F397B5157C68E9D675DF04C4973`

### OB Tool V1 RC57 (14/10/2026)
* **`RC57_HANDOFF.zip`**: SHA-256 `E1D019013F7EC27D861D02C1F9A7ED739F57908B17135C5BDAC01A4EBA2EBB3B`
* **`261010_OB_TOOL_V1_RC57.xlsm`**: SHA-256 `16FE33682976C325B6AF97E59489BD5198C4BA862F8A4FF957D71C1E699B5C98`
