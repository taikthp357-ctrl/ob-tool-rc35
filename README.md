# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC37 (10/10/2026)

### 📦 Các tệp đính kèm RC37

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC37_HANDOFF.zip`** | 738,195 bytes | `C762540D94D7366F13AE4CE650C2BB1DF5314947BE3A22B598A42EFC560E9EE5` | **Gói bàn giao đầy đủ RC37 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC37.xlsm`** | 784,358 bytes | `E66F327F3CC98A830E6C9037C735D5A7FE2FAC2FCBC1950DB26DC0840D84FB33` | Tệp Excel Macro phát hành RC37 (Đã dọn dẹp) |
| **`VBA_RC36_TO_RC37.diff`** | 31,668 bytes | — | Chi tiết bản vá mã nguồn VBA từ RC36 lên RC37 |
| **`CHANGELOG_RC37.md`** | 3,876 bytes | — | Nhật ký thay đổi và giải pháp kỹ thuật RC37 |

### 🚀 Điểm cải tiến cốt lõi trên RC37:
1. **Vá lỗi bẫy rỗng VL06O cho mọi Plant (Cả 3001 và 4007):**  
   Xóa bỏ điều kiện hạn chế `ship = "4007"`, áp dụng bẫy trạng thái rỗng ("no deliveries", "no data", "keine lieferungen") cho mọi Plant, xóa sạch 16 cột (`A6:P`) và thoát êm, triệt tiêu 100% nguy cơ treo 20s và lỗi Runtime Error 4202.
2. **Dọn dẹp siêu tốc lịch sử LOG và LEDGER:**  
   Dọn dẹp sạch 4,298 dòng LOG cũ và 250 dòng LEDGER cũ, giữ nguyên toàn bộ cấu trúc Header. Giảm dung lượng workbook từ 1,070,522 bytes xuống 784,358 bytes giúp mở, chạy và lưu siêu tốc.
3. **Cơ chế tự động bắt lại OB từ VL06O khi OB cũ bị lỗi/xóa:**  
   Khi OB gặp lỗi không tồn tại hoặc bị xóa trên SAP, tool tự động đánh dấu loại trừ (`mExcludedOB`), tải lại VL06O bảo toàn 16 cột (A:P), tự động chuyển sang OB mới nếu cùng STO sinh ra. Nếu không còn OB nào trong VL06O, hiển thị cảnh báo rõ ràng trên sheet INPUT: `"THIEU OB (VL06O) - Nhan nut 1 de tao lai qua VL10D hoac mo VL10D kiem tra"` kèm popup hướng dẫn người dùng.
4. **Cơ chế 1 lượt Retry tinh gọn cho OB bị Lock:**  
   OB bị lock/conflict đẩy về cuối hàng đợi và chỉ thử lại đúng 1 lần duy nhất, tránh vòng lặp.
5. **Cảnh báo tồn kho khóa / kiểm định (QI / QU / QT):**  
   Trong `PrepareStock` và `AllocateRequest`, nếu thiếu tồn khả dụng và kho có tồn thuộc QT, QI, QU, tự động bổ sung cảnh báo tại cột D sheet INPUT: `; (CO TON BI KHOA/QI/QU - XEM POSTING J)`.
6. **Biên dịch & Bảo toàn Invariants:**  
   VBE Debug Compile PASS 100% (`control.Enabled = False`), bảo toàn trọn vẹn 4 Invariants nhánh 3001 Shop Body, WholeBatchGate và nguyên tắc FIFO.

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC36 (10/10/2026)
* **`RC36_HANDOFF.zip`**: SHA-256 `35D034EC709FAB0724FDD50C3E0C525D8B5D362F7C3486D2F6929EE02D14B6B7`
* **`261010_OB_TOOL_V1_RC36.xlsm`**: SHA-256 `5BA00C8A1D6744A9CB2291A2DB924E26DA890B16EF1F43143796761EB1D609D7`
* Tự động bắt số OB từ VL06O theo STO + Material + Người tạo (ERNAM).

### OB Tool V1 RC35 R5 (10/10/2026)
* **`RC35_R5_HANDOFF.zip`**: SHA-256 `BBAAD4ABAF3C91C3B428C5CF177CEEC0D8FE94585F070ACBFFC4839CDD122A74`
* **`261010_OB_TOOL_V1_RC35_R5.xlsm`**: SHA-256 `98EE42AD70097067436EFD674AE069F1B9C2A72D66BEC3D4856394E2A833FF53`
* Khắc phục lỗi Active Error Handler (`Resume RetryWait`) và chốt chặn `EDCEnsureGroupInSnapshot`.
