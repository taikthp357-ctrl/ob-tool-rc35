# Stock FG MA BODY — Inventory Shortage & Warning System

Công cụ giám sát tồn kho thành phẩm (FG) và cảnh báo nguy cơ thiếu hàng dành cho Xưởng Thân Vỏ (BODY) và Xưởng Xe Máy Điện (XMĐ).

---

## 📌 Phiên bản phát hành mới nhất: Bản cập nhật (17/10/2026)

* **Tệp Excel chính thức**: [`261008_Stock_FG_MA_BODY.xlsm`](./261008_Stock_FG_MA_BODY.xlsm) (SHA-256: `C49C2F26A00629E9EB6E669799BCBD2D84FBC928D1CEB1A8C5916F3B0690A4AB`)
* **Gói nén ZIP tiêu chuẩn**: [`261008_Stock_FG_MA_BODY.zip`](./261008_Stock_FG_MA_BODY.zip) (SHA-256: `3B84CEFCB890289F43353E5CCF5E3D09B6374ED0D1AF500F30A2AF9509E2BC36`)
* **Bản phiên hiệu V3.7 đồng bộ**: [`261008_Stock_FG_MA_BODY_V3_7.xlsm`](./261008_Stock_FG_MA_BODY_V3_7.xlsm) (SHA-256: `5CE55CFDF3582F55C9B42FD33F9223BFB081B20546D5A3605CA5D2613EDB4B94`) | [`ZIP`](./261008_Stock_FG_MA_BODY_V3_7.zip)
* **Ghi chú kỹ thuật chi tiết**: [`RELEASE_NOTES_V3_7.md`](./RELEASE_NOTES_V3_7.md)

---

## 🛠 Tính năng nổi bật & Cập nhật mới (17/10/2026)
1. **Đổ màu chữ khi Copy vào Microsoft Teams (Thay đổi màu text, giữ nền chuẩn)**:
   - Khi bấm **`COPY CẢNH BÁO`** hoặc xem bảng cảnh báo:
     * Cột **`BD Kho`** (Tồn Kho Body / ESC Kho): Chữ hiển thị **Màu Đỏ** (`RGB(192, 0, 0)`), in đậm.
     * Cột **`BD Carset`** (Khả năng lắp ráp Body / ESC Carset): Chữ hiển thị **Màu Xanh** (`RGB(31, 122, 75)`), in đậm.
     * Chỉ đổi màu text (`.Font.Color`), tuyệt đối không tô màu nền ô, đảm bảo bảng dán vào Teams hiển thị sắc nét, chuyên nghiệp, tương thích cả Dark Mode và Light Mode.
2. **Định dạng US chuẩn số nguyên (Không dấu chấm thập phân)**:
   - Định mức US hiển thị dạng số nguyên thuần túy (`"0"`), không có phần thập phân hay dấu chấm phía sau (`1`, `2`, `4`, `36`...).
3. **Sửa logic cảnh báo vật tư MA**:
   - Khi Kho MA = 0 và Line MA > 0: Cảnh báo chính xác **`Ưu tiên kéo line`** (không còn báo sai "ưu tiên xuất Body/XMĐ" khi kho không có hàng).
   - Khi Kho MA > 0: `Ưu tiên xuất Body` / `Ưu tiên xuất XMĐ`.
   - Khi cả Kho MA và Line MA = 0: `MA hết tồn`.
4. **Đổi chuỗi trạng thái**:
   - Chuyển thành `Hàng chờ xuất sang Body` (đối với Shop BODY) và `Hàng chờ xuất sang XMĐ` (đối với Shop XMĐ).
5. **Chuẩn hóa Typography 100% Times New Roman**:
   - Toàn bộ 6 sheet (`CANH BAO`, `TON TONG`, `DATA`, `CAU HINH`, `MB52`, `LOG`), panel chọn model và nút bấm chuyển sang phông Times New Roman đồng bộ.
6. **Căn chỉnh & kẻ ô đầy đủ (Continuous Borders)**:
   - Kẻ đường viền đầy đủ cho tất cả các bảng dữ liệu trên mọi sheet, căn giữa mã/model/trạm, căn phải số lượng và US, căn giữa in đậm chữ đỏ cho cảnh báo.
7. **Nút chuyển đổi 1-Click (`Shop BODY` <-> `Shop XMĐ`)**:
   - Dễ dàng chuyển đổi trực tiếp giữa 2 xưởng sản xuất ngay trên giao diện mà không cần cấu hình phức tạp.
8. **Tách biệt và hiển thị rõ ràng mã Trước/Sau Sơn (Dành riêng cho XMĐ)**:
   - Cột 1: `FG trước sơn`
   - Cột 2: `FG sau sơn`
   - Cột 4: `ESC trước sơn`
   - Cột 5: `ESC sau sơn`
   - Cột 6: `ESC Carset`
9. **Bảo toàn dữ liệu Master Data**:
   - Sheet `DATA` chỉ đọc, bảo toàn vĩnh viễn mọi cập nhật mã vật tư mới (như chuyển đuôi mã `AA` -> `HA`), không bao giờ bị ghi đè hay hoàn tác về dữ liệu cũ.
