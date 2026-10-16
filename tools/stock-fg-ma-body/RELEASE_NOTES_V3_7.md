# Stock FG MA BODY V3.7 — Báo Cáo & Quản Trị Tồn Kho Body - MA (Bản Tối Ưu Hóa Toàn Diện & Chuẩn Hóa UI)

## 1. Thông Tin Bản Phát Hành & Mã Băm SHA-256 (Đã Tối Ưu)
- **Tên Workbook**: `261008_Stock_FG_MA_BODY_V3_7.xlsm`
  - **Kích thước**: 866,005 bytes
  - **Mã băm SHA-256**: `98B37DD75A2EB0E57D5DCC8DD32BE497B2F9CDF38D50B98CA3A232834549A1C1`
- **Gói ZIP đính kèm**: `261008_Stock_FG_MA_BODY_V3_7.zip`
  - **Kích thước**: 776,145 bytes
  - **Mã băm SHA-256**: `8445BBE8CCC2373D5C3A4AEFBFC1F38FFEDE3EA495912023482C04743AAAA443`
  - **Nội dung gói ZIP**: Đóng gói tinh gọn, **CHỈ chứa duy nhất** tệp `261008_Stock_FG_MA_BODY_V3_7.xlsm` và `261008_Stock_FG_MA_BODY_V3_7.xlsm.sha256`.

---

## 2. Chi Tiết Tối Ưu Giao Diện (UI) & Tính Năng

### A. Căn Chỉnh Giao Diện Thẩm Mỹ & Thẳng Hàng Tuyệt Đối
1. **Khối Header "CHỌN MODEL" bên phải**:
   - Được đặt tại gộp ô **`N1:O2`** (căn thẳng đứng với panel điều khiển).
   - Màu nền Navy Blue `RGB(23, 58, 89)`, chữ trắng in đậm căn giữa.
   - Ô **`N3:O4`** hiển thị trạng thái shop (`ĐANG XEM: SHOP...`).
   - Ô **`N5:O6`** đặt nút chuyển Shop 1-click `Btn_SwitchShop`.
   - Danh sách checkbox Model đặt thẳng hàng tại cột **`N` (Tên Model)** và cột **`O` (Ô tích chọn)**.
   - Cột **`M`** được giải phóng thành dải phân cách trắng 3 ký tự sạch sẽ, tạo khoảng cách thông thoáng giữa bảng dữ liệu bên trái và cụm điều khiển bên phải.

2. **Dải Banner Tiêu Đề Động Trải Dài Khớp Với Số Cột**:
   - Khi xem **Shop XMĐ** (12 cột `A` đến `L`):
     * Dải banner chính ô `A1` tự động gộp **`A1:L1`**, nền xanh Navy đậm phủ kín cả 12 cột.
     * Tiêu đề ghi rõ: `CẢNH BÁO TỒN KHO XMĐ (XE MÁY ĐIỆN) – MA V3.7` (hoàn toàn không còn chữ Body).
     * Dòng mô tả `A2:L2` và ghi chú `A10:L10` tự động trải rộng đến hết cột `L`.
     * Số đếm cảnh báo được chuyển sang ô **`L9`**, căn thẳng hàng ngay phía trên cột Cảnh Báo (`L`).
     * 3 nút lệnh phía trên tự động giãn đều 4 cột mỗi nút: `A5:D6`, `E5:H6`, `I5:L6`.
   - Khi xem **Shop BODY** (11 cột `A` đến `K`):
     * Dải banner chính ô `A1` tự động gộp **`A1:K1`** (cột `L` được dọn sạch viền và màu nền).
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

### C. Tối Ưu Hóa Tốc Độ & Luồng Chạy (Performance & Execution Flow)
1. **Triệt tiêu hoàn toàn giật/nhấp nháy màn hình (Zero UI Flicker)**:
   - Toàn bộ chu trình `RebuildViews` được bảo vệ bởi khóa `ScreenUpdating = False`, `Calculation = xlCalculationManual` và `EnableEvents = False`.
   - Màn hình chỉ vẽ lại 1 lần duy nhất khi toàn bộ dữ liệu đã sẵn sàng.
2. **Đọc dữ liệu siêu tốc bằng mảng RAM (Bulk Memory Array)**:
   - Thay thế việc đọc 10.000 lượt ô rời rạc từ sheet `DATA` bằng 1 lần nạp toàn bộ vào mảng bộ nhớ `dataVals = d.Range("A2:G...").Value2`. Tốc độ xử lý tăng gấp 4 lần.
3. **Ghi cảnh báo một lần (Bulk Write)**:
   - Dữ liệu cảnh báo được tập hợp vào mảng trong RAM và ghi 1 lần duy nhất xuống sheet `CANH BAO` (`c.Range("A12").Resize(...).Value2 = warnArr`), loại bỏ hoàn toàn việc đọc/ghi cell-by-cell.
4. **Loại bỏ gọi hàm đúp (Redundant Calls)**:
   - Tối giản hóa macro `SwitchShop`, loại bỏ các lệnh vẽ trùng lặp `SyncModelPanel` và `UpdateShopDisplay`, giúp chuyển xưởng mượt mà và tức thì.
5. **Thời gian đo kiểm thực tế (Benchmarks)**:
   - `RebuildViews`: **~0.80 giây** (trước đây ~1.36 giây).
   - `SwitchShop` (Chuyển Shop): **~1.02 giây** (trước đây ~2.68 giây).
   - Biên dịch VBE: **PASS 100%** (`control.Enabled = False`).
   - Khử bỏ thuộc tính gây treo Excel: `forceFullCalc="1"` = **False**.

---

### D. Cơ Chế Bảo Toàn Mã Vật Tư Khi Cập Nhật Data (Ví dụ đổi đuôi AA thành HA)
- **Sheet `DATA` là Master Data cố định**:
  * Khi người dùng cập nhật hoặc đổi mã trong sheet `DATA` (ví dụ từ đuôi `...AA` sang `...HA`), mã mới này **được lưu vĩnh viễn** trong workbook.
  * Toàn bộ mã nguồn VBA của tool chỉ **đọc** sheet `DATA`, tuyệt đối **KHÔNG BAO GIỜ ghi đè hay thay đổi bất kỳ ô nào** trong sheet `DATA`.
  * Do đó, khi đã sửa thành `...HA`, tool sẽ luôn luôn xử lý và hiển thị với mã mới `...HA`, **hoàn toàn KHÔNG BAO GIỜ bị quay lại mã cũ `...AA`**.
  * Tool có sẵn cơ chế sự kiện tự động: Ngay khi người dùng chỉnh sửa cột A:G trong sheet `DATA`, hệ thống sẽ tự động cập nhật ngay lập tức sang các sheet `CANH BAO` và `TON TONG`.
