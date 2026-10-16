# Stock FG MA BODY V3.7 — Báo Cáo & Quản Trị Tồn Kho Body - MA (Bản Chuẩn Hóa: Tiêu Đề Động & Hiển Thị Mã Trước / Sau Sơn)

## 1. Thông Tin Bản Phát Hành & Mã Băm SHA-256
- **Tên Workbook**: `261008_Stock_FG_MA_BODY_V3_7.xlsm`
  - **Kích thước**: 860,105 bytes
  - **Mã băm SHA-256**: `A40F61D7BD54A53A14F050213BAB38828735C3901E79DDF5B343CACA5708A851`
- **Gói ZIP đính kèm**: `261008_Stock_FG_MA_BODY_V3_7.zip`
  - **Kích thước**: 769,913 bytes
  - **Mã băm SHA-256**: `EE9F2520B4567D90FB1FA2401D4F3456FAA95B77896AE194507332E63B9123EF`
  - **Nội dung gói ZIP**: Đóng gói tinh gọn, **CHỈ chứa** tệp `261008_Stock_FG_MA_BODY_V3_7.xlsm` và `261008_Stock_FG_MA_BODY_V3_7.xlsm.sha256`.

---

## 2. Các Nâng Cấp Kỹ Thuật Đã Xử Lý Hoàn Toàn

### A. Tiêu Đề Động Theo Từng Shop (Không Còn Bị Cố Định Là Body)
1. **Tiêu đề Sheet `CANH BAO` (Ô A1)**:
   - Khi xem **Shop BODY**: `CẢNH BÁO TỒN KHO BODY – MA V3.7`
   - Khi bấm chuyển **Shop XMĐ**: Tự động đổi thành `CẢNH BÁO TỒN KHO XMĐ (XE MÁY ĐIỆN) – MA V3.7`
2. **Ghi chú tính toán (Ô A10)**:
   - Khi xem **Shop BODY**: `Chỉ cảnh báo DOH < ngưỡng model. Can build = INT(Tổng Body / US).`
   - Khi xem **Shop XMĐ**: `Chỉ cảnh báo DOH < ngưỡng model. Can build = INT(Tổng ESC / US).`
3. **Nội dung cảnh báo hành động**:
   - Khi xem **Shop BODY**: `Ưu tiên xuất Body` / `Chờ nhận hàng sang Body` / `MA hết tồn`
   - Khi xem **Shop XMĐ**: `Ưu tiên xuất XMĐ` / `Chờ nhận hàng sang XMĐ` / `MA hết tồn` (Tuyệt đối không còn chữ "Body").

---

### B. Biểu Hiện Đầy Đủ & Tách Biệt Các Mã Trước Sơn & Sau Sơn (Dành Riêng Cho XMĐ)
Khi chuyển sang **Shop XMĐ**, bảng tự động kích hoạt chế độ quản trị 2 mã vật tư:
- **Cột 1**: `FG trước sơn` (Mã linh kiện trước sơn, ví dụ: `BIW70011341`)
- **Cột 2**: `FG sau sơn` (Mã linh kiện sau sơn, ví dụ: `BIW70011351`)
- **Cột 3**: `Usage`
- **Cột 4**: `ESC trước sơn` (Tồn của mã trước sơn tại ESC: SLoc 3002, 1001, 600P, 1D01 Plant 2001)
- **Cột 5**: `ESC sau sơn` (Tồn của mã sau sơn tại ESC: SLoc 4001, 5001, 5002, 600P, 1D01, 3001 Plant 2001)
- **Cột 6**: `ESC Carset` (Khả năng lắp ráp xe tổng từ cả trước và sau sơn: `Int((ESC trước sơn + ESC sau sơn) / Usage)`)
- **Cột 7**: `Intransit`
- **Cột 8**: `Tồn Kho MA` (Plant 3093 của cả 2 mã)
- **Cột 9**: `Tồn Line MA` (Plant 3093 của cả 2 mã)
- **Cột 10**: `Model` (21 Model XMĐ)
- **Cột 11**: `Trạm` (Trích xuất từ File 2 KHSX XMĐ)
- **Cột 12** *(trên CANH BAO)*: `Cảnh báo`

---

### C. Shop BODY Giữ Nguyên 100% Theo Chuẩn Cũ V3.6
Khi chuyển lại về **Shop BODY**:
- Bảng tự động trở về **10 cột (TON TONG) / 11 cột (CANH BAO)**:
  `Material` | `Usage` | `Kho Body` | `Line Body` | `Can build Body` | `Intransit` | `Tồn Kho MA` | `Tồn Line MA` | `Model` | `Trạm` [| `Cảnh báo`]
- Không áp dụng khái niệm trước/sau sơn.
- Không lấy tồn Plant 2001.
- Bảo toàn đúng 24 cảnh báo chuẩn xác của Model VF8.

---

### D. Bố Cục Giao Diện Mới N & O
- Bảng chọn Checkbox Model và Nút Chuyển Shop `Btn_SwitchShop` được chuyển sang cột `N` và `O` (với cột `M` làm dải phân cách thẩm mỹ).
- Bảng dữ liệu tự động co giãn độ rộng cột tương ứng, không bị đè nút bấm hay che khuất thông tin.

---

## 3. Kiểm Thử Hệ Thống (Verification Suite)
- **Native VBE Debug Compile**: `PASS 100%` (`control.Enabled = False`).
- **XML Structure**: Khử hoàn toàn thuộc tính `forceFullCalc="1"`.
- **Test Shop BODY**: 24/24 cảnh báo VF8, 1,633 vật tư, tiêu đề chuẩn BODY.
- **Test Shop XMĐ**: 134 vật tư, hiển thị đầy đủ `FG trước sơn` và `FG sau sơn`, tiêu đề chuẩn XMĐ/ESC, cảnh báo ghi rõ XMĐ.
