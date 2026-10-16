# Stock FG MA BODY V3.7 — Báo Cáo & Quản Trị Tồn Kho Body - MA (Bản Chuẩn Hóa Logic Cảnh Báo, Định Dạng US Số Nguyên & UI Toàn Diện)

## 1. Thông Tin Bản Phát Hành & Mã Băm SHA-256 (Cập nhật 16/10/2026)
- **Tên Workbook**: `261008_Stock_FG_MA_BODY_V3_7.xlsm`
  - **Kích thước**: 911,482 bytes
  - **Mã băm SHA-256**: `D0A5154B65798553BF596CF580C9CA3BBBE5A788D9B9304A7583CD7EE14A7793`
- **Gói ZIP đính kèm**: `261008_Stock_FG_MA_BODY_V3_7.zip`
  - **Kích thước**: 816,763 bytes
  - **Mã băm SHA-256**: `C151BFA1262C07CCA5BE9B7BB31F884C3CC6E18B765D78814CAAF7BA57168639`
  - **Nội dung gói ZIP**: Đóng gói tinh gọn, **CHỈ chứa duy nhất** tệp `261008_Stock_FG_MA_BODY_V3_7.xlsm` và `261008_Stock_FG_MA_BODY_V3_7.xlsm.sha256`.

---

## 2. Chi Tiết Các Cải Tiến & Vá Lỗi Mới Nhất

### A. Chuẩn Hóa Định Mức Tiêu Hao (US) Là Số Nguyên Thuần Túy — Không Dấu Chấm Thập Phân
- **Phản hồi người dùng**: Định mức tiêu hao US là số lượng linh kiện cấu thành trên 1 xe (1, 2, 4, 36...), mang bản chất là số nguyên, tuyệt đối không có phần thập phân hay dấu chấm phía sau.
- **Xử lý**:
  * Đặt định dạng NumberFormat cho cột US trên toàn bộ các sheet (`DATA`, `CANH BAO`, `TON TONG`) thành `"0"`.
  * Hiển thị số nguyên gọn gàng, thuần túy (ví dụ: `1`, `2`, `4`, `36`), loại bỏ hoàn toàn dấu chấm thập phân `.`.

### B. Sửa Triệt Để Logic Cảnh Báo Khi Hàng Ở Line Nhưng Kho MA Hết Tồn
- **Vấn đề trước đây**: Khi tổng tồn MA > 0 nhưng Kho MA = 0 và Line MA > 0, tool vẫn đưa ra hành động `Ưu tiên xuất Body` / `Ưu tiên xuất XMĐ`. Đây là thông tin sai lệch vì kho MA không có hàng để xuất sang xưởng xe, hàng đang nằm tại Line MA.
- **Quy tắc mới chuẩn xác**:
  * Nếu **Tồn Kho MA > 0**: `Ưu tiên xuất Body` (hoặc `Ưu tiên xuất XMĐ` khi ở Shop XMĐ).
  * Nếu **Tồn Kho MA = 0 và Tồn Line MA > 0**: Báo chính xác **`Ưu tiên kéo line`** (bắt buộc phải kéo hàng từ Line MA).
  * Nếu **Tồn Kho MA = 0 và Tồn Line MA = 0**: Báo **`MA hết tồn`**.
  * Nếu có lượng Intransit đang chuyển kho: Bổ sung dòng **`Hàng chờ xuất sang Body`** (hoặc `Hàng chờ xuất sang XMĐ`).

### C. Chuẩn Hóa Chuỗi Trạng Thái Chờ Hàng
- Đổi toàn bộ chuỗi cũ *"chờ nhận hàng sang body"* thành chuẩn xác theo nghiệp vụ xuất - nhận:
  * **Shop BODY**: **`Hàng chờ xuất sang Body`**
  * **Shop XMĐ**: **`Hàng chờ xuất sang XMĐ`**

### D. Chuẩn Hóa Font Chữ 100% Times New Roman Toàn Workbook
- Áp dụng thống nhất phông chữ **Times New Roman** trên toàn bộ các thành phần:
  * Style mặc định **Normal** 11pt Times New Roman.
  * Toàn bộ 6 sheet: **`CANH BAO`**, **`TON TONG`**, **`DATA`**, **`CAU HINH`**, **`MB52`**, **`LOG`**.
  * Bảng điều khiển chọn model (Panel `N1:O50`).
  * Tất cả các khối Banner, mô tả, đếm cảnh báo và TextFrame của các nút bấm điều hướng.

### E. Căn Chỉnh & Kẻ Ô Đầy Đủ (Crisp Continuous Borders)
- Áp dụng đường viền kẻ ô liên tục (`xlContinuous`, màu `RGB(218, 225, 233)`) cho tất cả các vùng bảng biểu trên mọi sheet:
  * Sheet `CANH BAO` & `TON TONG`: Kẻ ô toàn diện từ dòng tiêu đề đến dòng dữ liệu cuối cùng, kèm tô màu so le (Zebra striping) trên `CANH BAO`.
  * Căn lề dữ liệu chuẩn xác:
    - Mã chi tiết / Mã trước / Mã sau / Model / Trạm: **Căn giữa (Center)**.
    - Định mức US: **Căn phải (Right)**, định dạng số nguyên thuần túy `"0"`.
    - Số lượng tồn kho (Kho Body, Line Body, Can build, Intransit, Tồn Kho MA, Tồn Line MA): **Căn phải (Right)**, định dạng phân cách hàng nghìn `#,##0`.
    - Cột Cảnh báo: **Căn giữa (Center)**, in đậm (Bold), chữ đỏ nổi bật `RGB(164, 58, 45)`.
  * Sheet `DATA`: Cột A:C & E:G căn giữa, cột US căn phải định dạng `"0"`, kẻ viền đầy đủ 1,768 dòng dữ liệu.
  * Sheet `CAU HINH`: Kẻ ô bảng thông số cấu hình và bảng ngưỡng DOH Model.
  * Sheet `MB52` & `LOG`: Kẻ viền đầy đủ toàn bộ bảng dữ liệu.

---

## 3. Kiến Trúc V3.7: Chuyển Đổi Nhanh Giữa BODY & XMĐ
1. **Nút chuyển đổi 1-Click (`SwitchShop`)**:
   - Nút lệnh tại ô `N5:O6` sheet `CANH BAO` cho phép chuyển tức thì giữa Shop BODY và Shop XMĐ.
2. **Hiển thị tách bạch mã Trước Sơn & Sau Sơn (Chỉ dành cho XMĐ)**:
   - **Cột 1**: `FG trước sơn`
   - **Cột 2**: `FG sau sơn`
   - **Cột 4**: `ESC trước sơn`
   - **Cột 5**: `ESC sau sơn`
   - **Cột 6**: `ESC Carset`
   - Shop BODY hoàn toàn giữ nguyên bảng 11 cột truyền thống, không bị ảnh hưởng bởi mã sau sơn.
3. **Bảo Toàn Master Data Tuyệt Đối**:
   - Khi cập nhật mã mới trong sheet `DATA` (ví dụ đổi đuôi `AA` thành `HA`), dữ liệu mới được bảo toàn vĩnh viễn, tool chỉ đọc và không bao giờ ghi đè hay hoàn tác về mã cũ.
4. **Kiểm Định Kỹ Thuật**:
   - Native VBE Debug Compile: **PASS 100%** (`control.Enabled = False`).
   - Khử bỏ thuộc tính gây treo Excel: `forceFullCalc="1"` = **False**.
