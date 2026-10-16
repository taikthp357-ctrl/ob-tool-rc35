# Stock FG MA BODY V3.7 — Báo Cáo & Quản Trị Tồn Kho Body - MA (Bản Chuẩn Hóa Toàn Diện: Căn Chỉnh UI & Phân Biệt Mã Trước / Sau Sơn)

## 1. Thông Tin Bản Phát Hành & Mã Băm SHA-256
- **Tên Workbook**: `261008_Stock_FG_MA_BODY_V3_7.xlsm`
  - **Kích thước**: 863,451 bytes
  - **Mã băm SHA-256**: `55E014E6D66D3324C571BDC7B7DF35E76CB16F2301B5FDC68ABB61C30D67137F`
- **Gói ZIP đính kèm**: `261008_Stock_FG_MA_BODY_V3_7.zip`
  - **Kích thước**: 773,300 bytes
  - **Mã băm SHA-256**: `986636358B0ED79DE45CDCF21AA6E99560A6B7B38F4B41E485B631AFCC6DBDDD`
  - **Nội dung gói ZIP**: Đóng gói tinh gọn, **CHỈ chứa duy nhất** tệp `261008_Stock_FG_MA_BODY_V3_7.xlsm` và `261008_Stock_FG_MA_BODY_V3_7.xlsm.sha256`.

---

## 2. Chi Tiết Tối Ưu Giao Diện (UI) & Tính Năng

### A. Căn Chỉnh Giao Diện Thẩm Mỹ & Thẳng Hàng Tuyệt Đối
1. **Khối Header "CHỌN MODEL" bên phải**:
   - Được dời sang gộp ô **`N1:O2`** (thay vì `M1:N2` lệch sang trái như trước).
   - Màu nền Navy Blue `RGB(23, 58, 89)`, chữ trắng in đậm căn giữa.
   - Ô **`N3:O4`** hiển thị trạng thái shop (`ĐANG XEM: SHOP...`).
   - Ô **`N5:O6`** đặt nút chuyển Shop 1-click `Btn_SwitchShop`.
   - Danh sách checkbox Model đặt thẳng hàng tại cột **`N` (Tên Model)** và cột **`O` (Ô tích chọn)**.
   - Cột **`M`** được giải phóng thành dải phân cách trắng 3 ký tự sạch sẽ, tạo khoảng cách thông thoáng giữa bảng dữ liệu bên trái và cụm điều khiển bên phải.

2. **Dải Banner Tiêu Đề Động Trải Dài Khớp Với Số Cột**:
   - Khi xem **Shop XMĐ** (12 cột `A` đến `L`):
     * Dải banner chính ô `A1` tự động gộp **`A1:L1`**, nền xanh Navy đậm phủ kín cả 12 cột.
     * Tiêu đề ghi rõ: `CẢNH BÁO TỒN KHO XMĐ (XE MÁY ĐIỆN) – MA V3.7`.
     * Dòng mô tả `A2:L2` và ghi chú `A10:L10` tự động trải rộng đến hết cột `L`.
     * Số đếm cảnh báo được chuyển sang ô **`L9`**, căn thẳng hàng ngay phía trên cột Cảnh Báo (`L`).
     * 3 nút lệnh phía trên tự động giãn đều 4 cột mỗi nút: `A5:D6`, `E5:H6`, `I5:L6`.
   - Khi xem **Shop BODY** (11 cột `A` đến `K`):
     * Dải banner chính ô `A1` tự động gộp **`A1:K1`** (cột `L` được xóa viền và màu nền).
     * Tiêu đề ghi rõ: `CẢNH BÁO TỒN KHO BODY – MA V3.7`.
     * Số đếm cảnh báo đặt tại ô **`K9`**, thẳng hàng ngay phía trên cột Cảnh Báo (`K`).
     * 3 nút lệnh co giãn chuẩn: `A5:D6`, `E5:G6`, `H5:K6`.

---

### B. Biểu Hiện Rõ Ràng Mã Trước Sơn & Sau Sơn (Chỉ Dành Cho XMĐ)
- Bảng hiển thị song song cả 2 mã:
  * **Cột 1**: `FG trước sơn` (Mã linh kiện trước sơn, ví dụ: `BIW70011341`)
  * **Cột 2**: `FG sau sơn` (Mã linh kiện sau sơn, ví dụ: `BIW70011351`)
  * **Cột 4**: `ESC trước sơn` (Tồn của mã trước sơn tại ESC)
  * **Cột 5**: `ESC sau sơn` (Tồn của mã sau sơn tại ESC)
  * **Cột 6**: `ESC Carset` (Khả năng lắp ráp xe tổng: `Int((ESC trước sơn + ESC sau sơn) / Usage)`)
- Cảnh báo hành động ghi rõ: `Ưu tiên xuất XMĐ` / `Chờ nhận hàng sang XMĐ` / `MA hết tồn` (Hoàn toàn không còn chữ "Body").

---

### C. Cơ Chế Bảo Toàn Mã Vật Tư Khi Cập Nhật Data (Ví dụ đổi đuôi AA thành HA)
- **Sheet `DATA` là Master Data cố định**:
  * Khi người dùng cập nhật hoặc đổi mã trong sheet `DATA` (ví dụ từ đuôi `...AA` sang `...HA`), mã mới này **được lưu vĩnh viễn** trong workbook.
  * Toàn bộ mã nguồn VBA của tool chỉ **đọc** sheet `DATA`, tuyệt đối **KHÔNG BAO GIỜ ghi đè hay thay đổi bất kỳ ô nào** trong sheet `DATA`.
  * Do đó, khi đã sửa thành `...HA`, tool sẽ luôn luôn xử lý và hiển thị với mã mới `...HA`, **hoàn toàn KHÔNG BAO GIỜ bị quay lại mã cũ `...AA`**.
  * Tool có sẵn cơ chế sự kiện tự động: Ngay khi người dùng chỉnh sửa cột A:G trong sheet `DATA`, hệ thống sẽ tự động cập nhật ngay lập tức sang các sheet `CANH BAO` và `TON TONG`.
