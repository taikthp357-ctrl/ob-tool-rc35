# Stock FG MA BODY — Inventory Shortage & Warning System

Công cụ giám sát tồn kho thành phẩm (FG) và cảnh báo nguy cơ thiếu hàng dành cho Xưởng Thân Vỏ (BODY) và Xưởng Xe Máy Điện (XMĐ).

---

## 📌 Phiên bản phát hành mới nhất: V3.7 (Cập nhật 16/10/2026)

* **Tệp Excel hoàn chỉnh**: [`261008_Stock_FG_MA_BODY_V3_7.xlsm`](./261008_Stock_FG_MA_BODY_V3_7.xlsm) (SHA-256: `D0A5154B65798553BF596CF580C9CA3BBBE5A788D9B9304A7583CD7EE14A7793`)
* **Gói nén ZIP tiêu chuẩn**: [`261008_Stock_FG_MA_BODY_V3_7.zip`](./261008_Stock_FG_MA_BODY_V3_7.zip) (SHA-256: `C151BFA1262C07CCA5BE9B7BB31F884C3CC6E18B765D78814CAAF7BA57168639`)
* **Ghi chú kỹ thuật chi tiết**: [`RELEASE_NOTES_V3_7.md`](./RELEASE_NOTES_V3_7.md)

---

## 🛠 Tính năng nổi bật V3.7
1. **Định dạng US chuẩn số nguyên (Không dấu chấm thập phân)**:
   - Định mức US hiển thị dạng số nguyên thuần túy (`"0"`), không có phần thập phân hay dấu chấm phía sau (`1`, `2`, `4`, `36`...).
2. **Sửa logic cảnh báo vật tư MA**:
   - Khi Kho MA = 0 và Line MA > 0: Cảnh báo chính xác **`Ưu tiên kéo line`** (không còn báo sai "ưu tiên xuất Body/XMĐ" khi kho không có hàng).
   - Khi Kho MA > 0: `Ưu tiên xuất Body` / `Ưu tiên xuất XMĐ`.
   - Khi cả Kho MA và Line MA = 0: `MA hết tồn`.
3. **Đổi chuỗi trạng thái**:
   - Chuyển thành `Hàng chờ xuất sang Body` (đối với Shop BODY) và `Hàng chờ xuất sang XMĐ` (đối với Shop XMĐ).
4. **Chuẩn hóa Typography 100% Times New Roman**:
   - Toàn bộ 6 sheet (`CANH BAO`, `TON TONG`, `DATA`, `CAU HINH`, `MB52`, `LOG`), panel chọn model và nút bấm chuyển sang phông Times New Roman đồng bộ.
5. **Căn chỉnh & kẻ ô đầy đủ (Continuous Borders)**:
   - Kẻ đường viền đầy đủ cho tất cả các bảng dữ liệu trên mọi sheet, căn giữa mã/model/trạm, căn phải số lượng và US, căn giữa in đậm chữ đỏ cho cảnh báo.
6. **Nút chuyển đổi 1-Click (`Shop BODY` <-> `Shop XMĐ`)**:
   - Dễ dàng chuyển đổi trực tiếp giữa 2 xưởng sản xuất ngay trên giao diện mà không cần cấu hình phức tạp.
7. **Tách biệt và hiển thị rõ ràng mã Trước/Sau Sơn (Dành riêng cho XMĐ)**:
   - Cột 1: `FG trước sơn`
   - Cột 2: `FG sau sơn`
   - Cột 4: `ESC trước sơn`
   - Cột 5: `ESC sau sơn`
   - Cột 6: `ESC Carset`
8. **Bảo toàn dữ liệu Master Data**:
   - Sheet `DATA` chỉ đọc, bảo toàn vĩnh viễn mọi cập nhật mã vật tư mới (như chuyển đuôi mã `AA` -> `HA`), không bao giờ bị ghi đè hay hoàn tác về dữ liệu cũ.
