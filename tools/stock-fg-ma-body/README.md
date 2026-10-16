# Stock FG MA BODY — Inventory Shortage & Warning System

Công cụ giám sát tồn kho thành phẩm (FG) và cảnh báo nguy cơ thiếu hàng dành cho Xưởng Thân Vỏ (BODY) và Xưởng Xe Máy Điện (XMĐ).

---

## 📌 Phiên bản phát hành mới nhất: V3.7 (16/10/2026)

* **Tệp Excel hoàn chỉnh**: [`261008_Stock_FG_MA_BODY_V3_7.xlsm`](./261008_Stock_FG_MA_BODY_V3_7.xlsm) (SHA-256: `55E014E6D66D3324C571BDC7B7DF35E76CB16F2301B5FDC68ABB61C30D67137F`)
* **Gói nén ZIP tiêu chuẩn**: [`261008_Stock_FG_MA_BODY_V3_7.zip`](./261008_Stock_FG_MA_BODY_V3_7.zip) (SHA-256: `986636358B0ED79DE45CDCF21AA6E99560A6B7B38F4B41E485B631AFCC6DBDDD`)
* **Ghi chú kỹ thuật**: [`RELEASE_NOTES_V3_7.md`](./RELEASE_NOTES_V3_7.md)

---

## 🛠 Tính năng nổi bật V3.7
1. **Nút chuyển đổi 1-Click (`Shop BODY` <-> `Shop XMĐ`)**:
   - Dễ dàng chuyển đổi trực tiếp giữa 2 xưởng sản xuất ngay trên giao diện mà không cần cấu hình phức tạp.
2. **Tách biệt và hiển thị rõ ràng mã Trước/Sau Sơn (Dành riêng cho XMĐ)**:
   - Cột 1: `FG trước sơn` (Ví dụ: `BIW70011341`)
   - Cột 2: `FG sau sơn` (Ví dụ: `BIW70011351`)
   - Cột 4: `ESC trước sơn` (Tồn kho đệm trước sơn)
   - Cột 5: `ESC sau sơn` (Tồn kho đệm sau sơn)
   - Cột 6: `ESC Carset` (Khả năng hoàn thiện carset lắp ráp tổng hợp)
3. **Tiêu đề và thuật ngữ hiển thị động theo xưởng**:
   - Tự động thay đổi tiêu đề lớn, tên cột và nội dung cảnh báo tương ứng với xưởng đang chọn.
4. **Căn chỉnh UI hoàn hảo**:
   - Panel `CHỌN MODEL` cố định tại cột `N:O`, phân tách bằng cột đệm `M`.
   - Nút bấm và cảnh báo tự động cân đối kích thước theo chế độ 11 cột (BODY) hoặc 12 cột (XMĐ).
5. **Bảo toàn dữ liệu Master Data**:
   - Sheet `DATA` chỉ đọc, bảo toàn vĩnh viễn mọi cập nhật mã vật tư mới (như chuyển đuôi mã `AA` -> `HA`), không bao giờ bị ghi đè hay hoàn tác về dữ liệu cũ.
