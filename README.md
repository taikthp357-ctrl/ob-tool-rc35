# OB TOOL V1 RC40

Canonical XLSM: `261010_OB_TOOL_V1_RC40.xlsm`  
Canonical SHA256: `191D37DAD2E884A5B26026B9332F2BF01CB0C64C8427D9985B5F267DEFAAF4C7`  
Handoff Archive: `RC40_HANDOFF.zip`  
Archive SHA256: `6E3B0C9B59815808EFEF63B5AFC1843C70708623F8947069C4BE9708A5E56574`

## Key Improvements in RC40
- **Smart Lock Backoff (Nhánh 3001)**: Tự động chờ 6s-5s (2 nhịp retry) khi gặp SAP `DO_AUTO` khóa PO, hiển thị đếm ngược trên Excel Status Bar và ghi log `WAIT_LOCK_BACKOFF`.
- **Mở khóa VL10D Initial Selection (Nhánh 4007)**: Khôi phục chuẩn RC34, bỏ chặn cứng `ST_RESWK-LOW` để gửi phím F8 (`sendVKey 8`) vào danh sách Due của VL10D.
- **Tự động dọn dẹp Pending cũ**: Xóa sạch cờ `VL10D_PENDING_KEY` khi bắt đầu `NEW_REQUEST` hoặc khi mã hàng đổi khác INPUT.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC39 (10/10/2026)

### 📦 Các tệp đính kèm RC39

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC39_HANDOFF.zip`** | 738,913 bytes | `36BAF1F6154731D8B363D2FEB462746042E32175DA6D18C3ED2870B27A9BAF7E` | **Gói bàn giao đầy đủ RC39 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC39.xlsm`** | 786,760 bytes | `1C374619A87D3A888DC8CA82E307FE8784747DE63AED9AE3FB6ABB30E3A27E32` | Tệp Excel Macro phát hành RC39 (Đã dọn dẹp & Compile) |
| **`VBA_RC38_TO_RC39.diff`** | 9,936 bytes | `911A9A9BA91E2C7A4F594CA547CD3D8E8C1AB526FD95FA0DB73E3D5ECDBBE877` | Chi tiết bản vá mã nguồn VBA từ RC38 lên RC39 |
| **`CHANGELOG_RC39.md`** | 5,690 bytes | `5823FA66D79BC1F7A5EF02E3D9E1F69131E4416AE8C465131FD8B6A9804D63AD` | Nhật ký thay đổi và giải pháp kỹ thuật RC39 |

### 🚀 Điểm cải tiến cốt lõi trên RC39 (Tối ưu hóa siêu tốc VL02N):
1. **[Dừng quét ngay khi hết item - Early Stop trong `OverviewSnapshot`]:**  
   - Bảng Delivery Overview của SAP (`tblSAPMV50ATC_LIPS_OVER`) luôn sắp xếp item liên tục từ dòng 0 xuống dưới.
   - Khi quét qua các dòng: Nếu gặp dòng trống (`itemNo = "" And material = ""`), tool kích hoạt dừng quét ngay lập tức (`GoTo finished`), triệt tiêu hoàn toàn việc cuộn qua vùng buffer 20-50 dòng trống của SAP.
   - Với đơn 1-3 item: Dừng ngay tại trang đầu tiên trong 0.05s, xóa bỏ 100% hiện tượng cuộn xuống đáy bảng!
2. **[Xóa bỏ lệnh gọi kép trùng lặp trong `OB_Apply`]:**  
   - Thay thế `Set mUnplannedOriginalSnapshot = OverviewSnapshot(ses)` bằng `Set mUnplannedOriginalSnapshot = mOverviewIndex`, loại bỏ hoàn toàn 1 lần quét toàn bảng trùng lặp.
3. **[Tối ưu hóa sau khi xóa item trong `RemoveUnplannedItems`]:**  
   - Sau khi bấm nút xóa và confirm: các item còn lại tự dồn lên đầu bảng (từ dòng 0).
   - Tool chụp nhanh trang hiện tại (dòng 0) để cập nhật `mOverviewIndex` cho các item còn lại chỉ trong ~0.02s mà không cần quét lại toàn bộ bảng.
4. **[Tối ưu hóa `LocateItem` - Tìm kiếm trực tiếp không hủy cache]:**  
   - Quét nhanh các dòng nhìn thấy trên trang hiện tại (`0 To VisibleRowCount - 1`) trước khi tính đến chuyện cuộn.
   - Không tự ý gán `Set mOverviewIndex = Nothing` để tránh kích hoạt `OverviewSnapshot` quét lại toàn bảng ở Pass 2.
   - Tìm kiếm với cơ chế Early Exit, dừng ngay khi tìm thấy item.
5. **Nghiệm thu & VBE Debug Compile PASS 100%:**  
   - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
   - Gắn macro nút bấm chuẩn xác (`OB_Check` và `OB_Execute`).
   - Bảo toàn trọn vẹn Rule 1, Rule 2, Rule 3, Rule 4, 4 Invariants nhánh 3001 Shop Body, WholeBatchGate và nguyên tắc FIFO.

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC38 (10/10/2026)
* **`RC38_HANDOFF.zip`**: SHA-256 `0864D60669766357A17001A4C52BEDD8D2F42453E4F11315D58CE1A532D411D6`
* **`261010_OB_TOOL_V1_RC38.xlsm`**: SHA-256 `6100917C4BAD6F68A3DED74A2E74CFEBA7A9006790458055DDD500D61CA51DC8`
* Khôi phục Fast Save từ RC34 (bỏ bắt buộc VL03N), khai thông đọc VL10D live Due Qty cho nhánh 4007 STO, hỗ trợ VL02N session reuse.

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
