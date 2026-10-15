# Stock FG MA BODY V3.7 — Báo Cáo & Quản Trị Tồn Kho Body - MA (Bản Chuẩn Hóa: Tách Biệt Tuyệt Đối BODY & XMĐ)

## 1. Thông Tin Bản Phát Hành & Mã Băm SHA-256
- **Tên Workbook**: `261008_Stock_FG_MA_BODY_V3_7.xlsm`
  - **Kích thước**: 832,761 bytes
  - **Mã băm SHA-256**: `F13A4941A2FAAE469F835C60781B24EB1AAC65466E6E66E57FAE09C7DA4A215B`
- **Gói ZIP đính kèm**: `261008_Stock_FG_MA_BODY_V3_7.zip`
  - **Kích thước**: 744,299 bytes
  - **Mã băm SHA-256**: `2873AEED79B9D278690EAEDB180B679BE273E3C969AFD0F96F664EF124157D3A`
  - **Nội dung gói ZIP**: Đóng gói tinh gọn và sạch sẽ, **CHỈ chứa** tệp `261008_Stock_FG_MA_BODY_V3_7.xlsm` và `261008_Stock_FG_MA_BODY_V3_7.xlsm.sha256`, tuyệt đối không có thư mục hay tệp phụ không liên quan.

---

## 2. Các Quy Tắc Kỹ Thuật Cốt Lõi Được Chuẩn Hóa

### A. Shop BODY: Giữ Nguyên 100% Logic V3.6 (Không Khái Niệm Trước/Sau Sơn, Không Plant 2001)
- **Tuyệt đối không áp dụng trước/sau sơn cho BODY**:
  - `Kho Body` = SLoc `1003` (Plant `3001`).
  - `Line Body` = SLoc `3003` (Plant `3001`).
  - `Can build Body` = `Int((Kho Body + Line Body) / Usage)`.
  - `Intransit` = `Stock in Transit` (Plant `3001`).
  - `Tồn Kho MA` = SLoc `3093` (Plant `3093`).
  - `Tồn Line MA` = SLoc `3000` (Plant `3093`).
  - `Model` = 13 Model chuẩn của BODY (`VF8`, `VF5`, `VF6`, `VF7`, `VF9`, `Limo`...).
  - `Trạm` = Dữ liệu Trạm trích xuất từ File 3 (`2026_ Bao cao ton kho_Body-MA.xlsb`).
- **Bảo toàn số lượng cảnh báo**: Shop BODY giữ đúng 24 cảnh báo chuẩn xác của Model VF8, hoàn toàn sạch sẽ và không bị ảnh hưởng bởi dữ liệu hay kho của Plant 2001.

### B. Shop XMĐ: Phân Định Rõ Ràng Mã Trước Sơn & Sau Sơn
- **Cơ chế đặc thù chỉ dành cho XMĐ**:
  - `Kho Body` = ESC trước sơn: SLoc `3002`, `1001`, `600P`, `1D01` (Plant `2001`) của mã trước sơn.
  - `Line Body` = ESC sau sơn: SLoc `4001`, `5001`, `5002`, `600P`, `1D01`, `3001` (Plant `2001`) của mã sau sơn (+ mã trước sơn nếu có).
  - `Can build Body` = `Int((Kho Body + Line Body) / Usage)` phản ánh chuẩn xác tổng linh kiện có thể lắp ráp từ cả trước và sau sơn.
  - `Tồn Kho MA` & `Tồn Line MA` = SLoc `3093`, `3000`, `1005` (Plant `3093`) cộng gộp cả mã trước và sau sơn.
  - `Model` = 21 Model của XMĐ (`VF6,7`, `VF8`, `VF9 LH`, `VFe34`...).
  - `Trạm` = Dữ liệu Trạm lấy từ sheet XMD của File 2 (`260917_ KHSX XMD..xlsb`).

### C. Nút Tổng Chuyển Shop 1-Click: `[ CHUYỂN SHOP: BODY / XMĐ ]`
- Nằm cố định tại ô `M5:N6` trên sheet `CANH BAO`.
- Click chuyển ngay lập tức giữa Shop BODY và XMĐ:
  - Tự động lọc danh sách Checkbox Model phù hợp bên phải (`M7:N...`).
  - Tự động cập nhật bảng `CANH BAO` và bảng `TON TONG` tương ứng.
  - Trạng thái hiển thị rõ ràng tại ô `M3`.

### D. Chuẩn Hóa Thứ Tự Cột 100% Theo Yêu Cầu
- **Sheet `TON TONG` (10 cột)**:
  `Material` | `Usage` | `Kho Body` | `Line Body` | `Can build Body` | `Intransit` | `Tồn Kho MA` | `Tồn Line MA` | `Model` | `Trạm`
- **Sheet `CANH BAO` (11 cột)**:
  `Material` | `Usage` | `Kho Body` | `Line Body` | `Can build Body` | `Intransit` | `Tồn Kho MA` | `Tồn Line MA` | `Model` | `Trạm` | `Cảnh báo`

---

## 3. Kiểm Thử Hệ Thống (Verification Suite)
- **Native VBE Debug Compile**: `PASS 100%` (`control.Enabled = False`).
- **XML Structure**: Thuộc tính `forceFullCalc="1"` bị loại bỏ hoàn toàn, workbook mở mượt mà.
- **BODY Verification**: 1,633 vật tư, đúng 24 cảnh báo VF8.
- **XMĐ Verification**: 134 vật tư, tính toán chính xác tồn trước/sau sơn.
