# Stock FG MA BODY — Inventory Shortage & Warning System

Công cụ giám sát tồn kho thành phẩm (FG) và cảnh báo nguy cơ thiếu hàng dành cho Xưởng Thân Vỏ (BODY) và Xưởng Xe Máy Điện (XMĐ).

---

## 📌 Phiên bản phát hành mới nhất: V3.7 (Cập nhật 16/10/2026)

* **Tệp Excel hoàn chỉnh**: [`261008_Stock_FG_MA_BODY_V3_7.xlsm`](./261008_Stock_FG_MA_BODY_V3_7.xlsm) (SHA-256: `84C5F2FFF853E119052CAFC3CDC1609D713C5B4BBF052A1C84F752AE077CD07F`)
* **Gói nén ZIP tiêu chuẩn**: [`261008_Stock_FG_MA_BODY_V3_7.zip`](./261008_Stock_FG_MA_BODY_V3_7.zip) (SHA-256: `7DA6952E3DF9EEDF6665405C531758A1D42B735476A2332706CDCDEE937F0D83`)
* **Ghi chú kỹ thuật chi tiết**: [`RELEASE_NOTES_V3_7.md`](./RELEASE_NOTES_V3_7.md)

---

## 🛠 Tính năng nổi bật V3.7
1. **Sửa logic cảnh báo vật tư MA**:
   - Khi Kho MA = 0 và Line MA > 0: Cảnh báo chính xác **`Ưu tiên kéo line`** (không còn báo sai "ưu tiên xuất Body/XMĐ" khi kho không có hàng).
   - Khi Kho MA > 0: `Ưu tiên xuất Body` / `Ưu tiên xuất XMĐ`.
   - Khi cả Kho MA và Line MA = 0: `MA hết tồn`.
2. **Đổi chuỗi trạng thái**:
   - Chuyển thành `Hàng chờ xuất sang Body` (đối với Shop BODY) và `Hàng chờ xuất sang XMĐ` (đối với Shop XMĐ).
3. **Chuẩn hóa Typography 100% Times New Roman**:
   - Toàn bộ 6 sheet (`CANH BAO`, `TON TONG`, `DATA`, `CAU HINH`, `MB52`, `LOG`), panel chọn model và nút bấm chuyển sang phông Times New Roman đồng bộ.
4. **Căn chỉnh & kẻ ô đầy đủ (Continuous Borders)**:
   - Kẻ đường viền đầy đủ cho tất cả các bảng dữ liệu trên mọi sheet, căn giữa mã/model/trạm, căn phải số lượng và US, căn giữa in đậm chữ đỏ cho cảnh báo.
5. **Nút chuyển đổi 1-Click (`Shop BODY` <-> `Shop XMĐ`)**:
   - Dễ dàng chuyển đổi trực tiếp giữa 2 xưởng sản xuất ngay trên giao diện mà không cần cấu hình phức tạp.
6. **Tách biệt và hiển thị rõ ràng mã Trước/Sau Sơn (Dành riêng cho XMĐ)**:
   - Cột 1: `FG trước sơn`
   - Cột 2: `FG sau sơn`
   - Cột 4: `ESC trước sơn`
   - Cột 5: `ESC sau sơn`
   - Cột 6: `ESC Carset`
7. **Bảo toàn dữ liệu Master Data**:
   - Sheet `DATA` chỉ đọc, bảo toàn vĩnh viễn mọi cập nhật mã vật tư mới (như chuyển đuôi mã `AA` -> `HA`), không bao giờ bị ghi đè hay hoàn tác về dữ liệu cũ.
