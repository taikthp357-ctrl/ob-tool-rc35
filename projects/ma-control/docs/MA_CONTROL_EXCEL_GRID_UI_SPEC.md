# MA CONTROL — ĐẶC TẢ GIAO DIỆN BẢNG KIỂU EXCEL (EXCEL-LIKE GRID UI)

- **Mã tài liệu**: `MA-CONTROL-UI-SPEC`
- **Mục tiêu**: Hiện thực hóa 100% bản mockup "MA CONTROL - Điều hành FG" (Màn hình "Điền OB / Ra đơn pick" kiểu Excel).
- **Tác giả**: Astra — Supervisory Architect
- **Ngày lập**: 13/10/2026

---

## 1. THIẾT KẾ TỔNG THỂ & BỐ CỤC MÀN HÌNH (LAYOUT SPECIFICATION)

Giao diện áp dụng bảng màu Dark/Modern Slate chuyên nghiệp kết hợp điểm nhấn xanh lá Microsoft Excel (#107c41) và xanh dương Enterprise (#2563eb).

### 1.1. Bố cục 3 phần chính:
1. **Sidebar bên trái (Navigation Bar)**:
   - Logo thương hiệu: `MA CONTROL` (kèm icon điều hành).
   - Menu điều hướng:
     * `Điều hành FG` (Icon dashboard tổng quan).
     * `Tồn & ưu tiên` (Icon kho hàng, phân bổ tồn).
     * `Điền OB / Đơn pick` (Mục đang hoạt động - Active Highlighted).
     * `Chờ CF` (Icon đồng hồ cát, số lượng badge đơn chờ xác nhận).
     * `BBBG` (Icon chứng từ biên bản bàn giao).
     * `Lịch sử` (Icon nhật ký kiểm toán).
     * `Cấu hình` (Icon bánh răng thiết lập hệ thống).
   - Footer sidebar: Thông tin phiên bản `Build 261013 • Localhost`.

2. **Khu vực Nội dung Chính (Main Work Area)**:
   - **Header Top**:
     * Tiêu đề: `MA CONTROL — Điều hành FG`
     * Huy hiệu trạng thái kết nối:
       - `SAP đã kết nối` (Chấm tròn xanh lá neon `●` + nền pill xanh nhạt `#ecfdf5`, viền `#10b981`)
       - `RC52 • FG` (Tag xám bo góc)
       - Thông tin người dùng: `User: GL_BODY_01`
   - **Thanh tiến trình 5 bước liên hoàn (Horizontal Stepper)**:
     * `1. Tải tồn` (Icon tick xanh tròn `✓`, đã hoàn tất)
     * Đường nối xanh ➔
     * `2. Điền OB` (Vòng tròn số 2 nền xanh dương nổi bật, trạng thái Active)
     * Đường nối xám ➔
     * `3. Ra đơn pick` (Vòng tròn số 3 viền xám, bước kế tiếp)
     * Đường nối xám ➔
     * `4. Người dùng CF` (Vòng tròn số 4 kèm nhãn "Bắt buộc kiểm tra thủ công")
     * Đường nối xám ➔
     * `5. BBBG` (Vòng tròn số 5 hoàn tất hồ sơ & GOS)

3. **Thanh công cụ & Phân nhánh (Toolbar & Sub-Tabs)**:
   - **2 Tab phân nhánh nghiệp vụ**:
     * `Body • 3001` (Active: line-border dưới màu xanh dương, quản lý draft treo và bù trừ batch)
     * `EDC • 4007` (Quản lý STO, VL10D và khóa chứng từ)
   - **Các nút thao tác dữ liệu nhanh**:
     * Nút `Lấy từ tồn`: Modal chọn các mã ưu tiên xuất từ snapshot tồn kho mới nhất.
     * Nút `Dán từ Excel`: Hỗ trợ mở ô dán clipboard trực tiếp nếu trình duyệt chặn paste tự động.
     * Nút `+ Thêm dòng`: Thêm dòng trống vào cuối bảng.
   - **Thanh tìm kiếm và bộ lọc nhanh**:
     * Ô nhập: `Tìm mã / OB...` (Icon kính lúp)
     * Dropdown lọc Model: `Tất cả model` (VF5, VF8, VF89 LIMO, VF7...)
     * Dropdown lọc Trạng thái: `Tất cả trạng thái` (Sẵn sàng điền, Chưa kiểm tra, Thiếu SL OB, Chờ pick / CF)
   - **Bộ 3 nút hành động chính (Action Buttons)**:
     * `Kiểm tra` (Nút viền xám / outline, icon kính lúp kiểm tra số lượng và tính sẵn sàng)
     * `Điền OB` (Nút xanh dương đậm / primary button, icon SAP play)
     * `Ra đơn pick` (Nút xanh lục đậm / success button, icon clipboard)

---

## 2. ĐẶC TẢ LƯỚI THAO TÁC TRUNG TÂM (EXCEL-LIKE DATA GRID)

Lưới dữ liệu trung tâm mang lại cảm giác thao tác tự nhiên tương đương Microsoft Excel:

### 2.1. Cấu trúc Cột và Ký hiệu:
| Ký hiệu Excel | Tên cột hiển thị | Loại dữ liệu | Tính chất chỉnh sửa | Căn lề | Ghi chú |
| :---: | :--- | :--- | :--- | :---: | :--- |
| **A** | `[x]` (Checkbox) | Boolean | Checkbox click | Giữa | Chọn nhiều dòng để thao tác gộp |
| **B** | `Mã hàng` | Text (Alphanumeric) | Cho phép sửa trực tiếp | Trái | Mã chi tiết (BIW73012711CA...) |
| **C** | `Model` | Text | Dropdown / Tự nhận diện | Giữa | Tự động điền theo mã hoặc nhập tay |
| **D** | `SL yêu cầu` | Number (Integer/Decimal) | Cho phép sửa trực tiếp | Phải | Số lượng PCS cần xuất |
| **E** | `Batch` | Text | Cho phép sửa trực tiếp | Giữa | Số lô xuất kho (ví dụ: `26101301`) |
| **F** | `Shop` | Text | Cho phép sửa trực tiếp | Giữa | `BODY`, `MA`, `EDC` |
| **G** | `STO` | Text | Cho phép sửa trực tiếp | Giữa | Số STO (dành cho nhánh EDC 4007) |
| **H** | `OB` | Text | Tự động điền sau khi tạo | Giữa | Số chứng từ Outbound Delivery SAP |
| **I** | `Trạng thái` | Status Chip | Read-only (Hệ thống tính toán) | Giữa | Badge màu thể hiện tiến trình dòng |

### 2.2. Quy cách Status Chip (Huy hiệu trạng thái):
- **Sẵn sàng điền**: Nền xanh lá nhạt (`#dcfce7`), chữ xanh lá đậm (`#166534`), viền `#86efac`.
- **Chưa kiểm tra**: Nền xám nhạt (`#f3f4f6`), chữ xám đậm (`#4b5563`), viền `#d1d5db`.
- **Thiếu SL OB**: Nền vàng/cam nhạt (`#fef3c7`), chữ cam đậm (`#92400e`), viền `#fcd34d`, kèm icon cảnh báo tam giác `⚠`.
- **Chờ pick / CF**: Nền xanh dương nhạt (`#dbeafe`), chữ xanh dương đậm (`#1e40af`), viền `#93c5fd`.

### 2.3. Trải nghiệm Tương tác Kiểu Excel (Keyboard & Clipboard Engine):
1. **Dán từ Clipboard (Ctrl+V)**:
   - Người dùng sao chép vùng bảng từ Excel (ví dụ 5 cột x 10 dòng) và ấn `Ctrl+V` vào ô bất kỳ trên lưới.
   - Trình duyệt bắt sự kiện `paste`, tách dữ liệu theo dòng (`\r\n` hoặc `\n`) và cột (`\t`).
   - Tự động điền giá trị vào các ô tương ứng bắt đầu từ ô đang được chọn (Active Cell). Nếu số dòng paste vượt quá số dòng hiện tại, lưới tự động sinh thêm dòng mới.
2. **Điều hướng bằng bàn phím**:
   - Phím `Tab`: Chuyển sang ô liền kề bên phải (nếu ở cuối dòng thì xuống ô đầu dòng kế tiếp).
   - Phím `Shift + Tab`: Chuyển sang ô liền kề bên trái.
   - Phím `Enter`: Chuyển xuống ô ngay bên dưới trên cùng một cột.
   - Phím `Mũi tên Up/Down/Left/Right`: Di chuyển ô chọn sang 4 hướng.
3. **Hiển thị ô đang chọn (Cell Focus)**:
   - Ô đang chọn có viền màu xanh lá cây Microsoft Excel dày 2px (`outline: 2px solid #107c41`).
   - Khi double-click hoặc bắt đầu gõ phím, ô chuyển sang chế độ inline edit input để người dùng nhập văn bản trực tiếp.
4. **Thanh Footer Tổng kết Lưới**:
   - Phía trái: `Đã chọn: X dòng • Tổng SL yêu cầu: Y PCS`.
   - Phía phải: Hướng dẫn nhanh `[Ctrl+V]: Dán nhiều ô/dòng • [Tab/Enter]: Di chuyển • [Phím cách]: Bật/tắt chọn dòng`.

---

## 3. KHU VỰC "ĐƠN PICK ĐÃ TẠO" & CỔNG KIỂM TRA CF THỦ CÔNG

Nằm ngay bên dưới lưới chính, thể hiện giai đoạn chuyển giao giữa Bước 3 và Bước 4:
- **Tiêu đề section**: `Đơn pick đã tạo (Chờ người dùng CF)`
- **Bảng danh sách đơn pick**:
  * `STT`: 1, 2, 3...
  * `Số đơn pick`: Ví dụ `PICK-20261013-001` (in đậm)
  * `Ngày tạo`: `13/10/2026 08:30`
  * `Shop`: `BODY`
  * `Model`: `VF8`
  * `Số mã hàng`: `4 mã`
  * `Tổng SL (PCS)`: `1,250`
  * `Trạng thái`: Chip màu vàng/xanh dương `Chờ người dùng CF` (chữ nhấp nháy nhẹ hoặc biểu tượng đồng hồ cát)
  * `Thao tác`:
    - Nút **"Kiểm tra CF"** (Primary Button, màu xanh dương): Nhấp vào sẽ mở modal đối chiếu chi tiết hàng hóa thực tế và cho phép người dùng click "Xác nhận CF thành công".
    - Nút **"..."** (Menu phụ): Xem chi tiết, In phiếu pick, Hủy đơn pick.
- **Dòng nhắc nhở nghiệp vụ (Business Callout)**:
  > ℹ️ **Lưu ý bắt buộc**: Sau khi người dùng kiểm tra thực tế và bấm xác nhận CF thành công, đơn pick mới được phép chuyển sang bước tạo **BBBG (Biên bản bàn giao)**. Tuyệt đối không tự động xác nhận ngầm.
