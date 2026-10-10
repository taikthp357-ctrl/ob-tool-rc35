# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC38 (10/10/2026)

### 📦 Các tệp đính kèm RC38

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC38_HANDOFF.zip`** | 735,873 bytes | `0864D60669766357A17001A4C52BEDD8D2F42453E4F11315D58CE1A532D411D6` | **Gói bàn giao đầy đủ RC38 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC38.xlsm`** | 782,925 bytes | `6100917C4BAD6F68A3DED74A2E74CFEBA7A9006790458055DDD500D61CA51DC8` | Tệp Excel Macro phát hành RC38 (Đã dọn dẹp & Compile) |
| **`VBA_RC37_TO_RC38.diff`** | 14,834 bytes | `755BE34D28E3D12337C42CB6C52879F8DC39C9AAA3FF06C984FD145D246D0B4C` | Chi tiết bản vá mã nguồn VBA từ RC37 lên RC38 |
| **`CHANGELOG_RC38.md`** | 6,199 bytes | `359CEB1D3BAACF59F66995D63497872A649AA751C6B0B01DFE336A65A162D0E3` | Nhật ký thay đổi và giải pháp kỹ thuật RC38 |

### 🚀 Điểm cải tiến cốt lõi trên RC38:
1. **[P0 - Nhánh 3001] Khôi phục Fast Save từ RC34 (Triệt tiêu lỗi văng VL03N):**  
   - Chấp nhận biên nhận lưu hợp lệ từ SAP (`statusType = "S"` và `IsSaveReceipt` như `VL 337`) để ghi nhận `DONE` ngay lập tức, không bắt buộc mở lại VL03N để hậu kiểm trừ khi có cờ `STRICT_AUDIT` hoặc biên nhận không rõ ràng.
   - Triệt tiêu hoàn toàn lỗi màn hình VL03N thiếu cột layout `LIPS-POSNR` hạ trạng thái lưu hợp lệ thành `SAVED_UNVERIFIED`.
2. **[P0 - Nhánh 4007] Cơ chế đọc Due Qty VL10D thông suốt khi thiếu OB:**  
   - Luôn cho phép kiểm tra số lượng STO còn Due trong VL10D ở chế độ Read-Only mà không bị chặn bởi cờ `VL10D_PENDING_KEY`.
   - Đối soát pending cũ tự động nếu STO còn Due, cho phép tạo tiếp phần còn thiếu. Nếu không đủ Due, cảnh báo an toàn `WAIT_STO_DUE_4007`.
3. **[P1] Tái sử dụng phiên VL02N (VL02N Session Reuse):**  
   - Hỗ trợ `ToolConfigBool("VL02N_SESSION_REUSE", True)`. Kết hợp với Fast Save giữ nguyên màn hình ban đầu của VL02N, tool nhập trực tiếp số OB tiếp theo mà không cần điều hướng lại từ `/nVL02N`.
4. **[P1] Tách bạch chứng cứ VL10D và VL06O:**  
   - Phân định rõ ràng chứng cứ `VL06O_AUTO_DISCOVERED` với `VL10D_CREATION_LOG`, hỗ trợ cả 2 nguồn trong khâu xác minh `EDCVerifyGroup`.
5. **Nghiệm thu 6/6 Acceptance Criteria & VBE Debug Compile PASS 100%:**  
   - Biên dịch sạch không lỗi trên Excel Native COM.
   - Gắn macro nút bấm chuẩn xác (`OB_Check` và `OB_Execute`).
   - Bảo toàn nguyên vẹn 4 Invariants nhánh 3001 Shop Body, WholeBatchGate và nguyên tắc FIFO.

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC37 (10/10/2026)
* **`RC37_HANDOFF.zip`**: SHA-256 `C762540D94D7366F13AE4CE650C2BB1DF5314947BE3A22B598A42EFC560E9EE5`
* **`261010_OB_TOOL_V1_RC37.xlsm`**: SHA-256 `E66F327F3CC98A830E6C9037C735D5A7FE2FAC2FCBC1950DB26DC0840D84FB33`
* Vá lỗi bẫy rỗng VL06O cho mọi Plant, dọn dẹp lịch sử LOG/LEDGER, tự động bắt lại OB từ VL06O, cảnh báo tồn kho khóa QI/QU/QT.

### OB Tool V1 RC36 (10/10/2026)
* **`RC36_HANDOFF.zip`**: SHA-256 `35D034EC709FAB0724FDD50C3E0C525D8B5D362F7C3486D2F6929EE02D14B6B7`
* **`261010_OB_TOOL_V1_RC36.xlsm`**: SHA-256 `5BA00C8A1D6744A9CB2291A2DB924E26DA890B16EF1F43143796761EB1D609D7`
* Tự động bắt số OB từ VL06O theo STO + Material + Người tạo (ERNAM).

### OB Tool V1 RC35 R5 (10/10/2026)
* **`RC35_R5_HANDOFF.zip`**: SHA-256 `BBAAD4ABAF3C91C3B428C5CF177CEEC0D8FE94585F070ACBFFC4839CDD122A74`
* **`261010_OB_TOOL_V1_RC35_R5.xlsm`**: SHA-256 `98EE42AD70097067436EFD674AE069F1B9C2A72D66BEC3D4856394E2A833FF53`
* Khắc phục lỗi Active Error Handler (`Resume RetryWait`) và chốt chặn `EDCEnsureGroupInSnapshot`.
