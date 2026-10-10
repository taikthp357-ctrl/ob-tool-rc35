# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC36 (10/10/2026)

### 📦 Các tệp đính kèm RC36

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC36_HANDOFF.zip`** | 980,267 bytes | `35D034EC709FAB0724FDD50C3E0C525D8B5D362F7C3486D2F6929EE02D14B6B7` | **Gói bàn giao đầy đủ RC36 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC36.xlsm`** | 1,070,522 bytes | `5BA00C8A1D6744A9CB2291A2DB924E26DA890B16EF1F43143796761EB1D609D7` | Tệp Excel Macro phát hành RC36 |
| **`VBA_RC35_TO_RC36.diff`** | 207 dòng | — | Chi tiết bản vá mã nguồn VBA từ RC35 lên RC36 |
| **`CHANGELOG_RC36.md`** | 1,845 bytes | — | Nhật ký thay đổi và giải pháp kỹ thuật RC36 |

### 🚀 Điểm cải tiến cốt lõi trên RC36:
* **Tự động bắt số OB từ VL06O sau khi tạo ở VL10D:**  
  Không cần người dùng phải tra Creation Log thủ công hay copy-paste biên nhận vào sheet CONFIG. Tool tự động làm mới VL06O và đối soát theo bộ 4 tiêu chí:
  1. **Mã STO** (`VGBEL`) trùng với STO đang xử lý.
  2. **Mã Vật tư** (`MATNR`) trùng với mã INPUT.
  3. **Người tạo** (`ERNAM`) trùng khớp với SAP User ID đang chạy tool (`session.Info.User`).
  4. **Trạng thái** (`CANDIDATE`, GM = 'A', chưa gán SLoc/Batch).
* **Biên dịch & Kiểm thử:** VBE Debug Compile PASS 100%, bảo toàn 100% 4 Invariants nhánh 3001 Shop Body.

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC35 R5
* **`RC35_R5_HANDOFF.zip`**: SHA-256 `BBAAD4ABAF3C91C3B428C5CF177CEEC0D8FE94585F070ACBFFC4839CDD122A74`
* **`261010_OB_TOOL_V1_RC35_R5.xlsm`**: SHA-256 `98EE42AD70097067436EFD674AE069F1B9C2A72D66BEC3D4856394E2A833FF53`
* Khắc phục lỗi Active Error Handler (`Resume RetryWait`) và chốt chặn `EDCEnsureGroupInSnapshot`.
