# OB Tool & Stock Management Releases — Public Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)** và **Stock FG MA BODY**.

---

## 📊 Bản phát hành Công Cụ Quản Trị Tồn Kho: Stock FG MA BODY V3.7 (16/10/2026)

Bản cập nhật hoàn thiện: Tách biệt tuyệt đối giữa Shop BODY và XMĐ, tích hợp dữ liệu Trạm Line MA và KHSX XMĐ.

### 📦 Các tệp đính kèm V3.7

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`261008_Stock_FG_MA_BODY_V3_7.zip`** | 744,299 bytes | `2873AEED79B9D278690EAEDB180B679BE273E3C969AFD0F96F664EF124157D3A` | **Gói nén ZIP V3.7 (CHỈ chứa .xlsm và .sha256)** |
| **`261008_Stock_FG_MA_BODY_V3_7.xlsm`** | 832,761 bytes | `F13A4941A2FAAE469F835C60781B24EB1AAC65466E6E66E57FAE09C7DA4A215B` | Tệp Excel Macro V3.7 hoàn chỉnh (Compile PASS 100%) |
| **`RELEASE_NOTES_V3_7.md`** | — | — | Ghi chú kỹ thuật chi tiết bản V3.7 |

### 🚀 Điểm cốt lõi bản V3.7
1. **Shop BODY giữ nguyên 100% logic V3.6**:
   - Không áp dụng khái niệm trước sơn hay sau sơn cho BODY.
   - Không lấy tồn Plant 2001 (kho sơn 4001, thành phẩm 5001).
   - Bảo toàn chuẩn xác 24 cảnh báo của Model VF8.
2. **Shop XMĐ quản lý tách biệt mã trước & sau sơn**:
   - `Kho Body` = ESC trước sơn (SLoc 3002, 1001, 600P, 1D01 tại Plant 2001).
   - `Line Body` = ESC sau sơn (SLoc 4001, 5001, 5002, 600P, 1D01, 3001 tại Plant 2001).
   - `Tồn Kho MA` & `Tồn Line MA` = Plant 3093 cho cả trước và sau sơn.
3. **Nút bấm 1-click Chuyển Shop `[ CHUYỂN SHOP: BODY / XMĐ ]`**:
   - Tọa độ `M5:N6` trên sheet `CANH BAO`.
   - Tự động hoán đổi danh sách Model tương ứng (13 Model BODY hoặc 21 Model XMĐ).
4. **Chuẩn hóa thứ tự cột**:
   - **TỒN TỔNG (10 cột)**: `Material` | `Usage` | `Kho Body` | `Line Body` | `Can build Body` | `Intransit` | `Tồn Kho MA` | `Tồn Line MA` | `Model` | `Trạm`
   - **CẢNH BÁO (11 cột)**: `Material` | `Usage` | `Kho Body` | `Line Body` | `Can build Body` | `Intransit` | `Tồn Kho MA` | `Tồn Line MA` | `Model` | `Trạm` | `Cảnh báo`

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC64 (16/10/2026)

### 📦 Các tệp đính kèm RC64

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC64_HANDOFF.zip`** | 935,918 bytes | `80DDDC1D787B7058A4F530285080E316C853D15E5FB5F67472C9332A1D86DA9F` | **Gói bàn giao đầy đủ RC64 (Streamlined VL10D & ALL IN ONE)** |
| **`261010_OB_TOOL_V1_RC64.xlsm`** | 949,326 bytes | `9A9D9A3875D037C5DF5771932C1622C993780AE41131D52B6ADA90A8FC199FC5` | Tệp Excel Macro phát hành RC64 (Đã tối ưu & Compile PASS 100%) |
| **`VBA_RC63_TO_RC64.diff`** | 10,677 bytes | `679885A193EE90D303B0725D20F0A046D6C8A4F844B79116342110A6619D427A` | Chi tiết bản vá mã nguồn VBA từ RC63 lên RC64 |
| **`CHANGELOG_RC63_TO_RC64.md`** | 4,323 bytes | `7EE3DC060A907B4F5DDF09C6F74059617689981C9B4C1D437A4A3BFB599CDD5E` | Nhật ký thay đổi và giải pháp kỹ thuật RC64 |

### 🚀 Điểm cải tiến chính RC64
1. **Kiểm tra mức giảm STO Due trực tiếp trên VL10D (F8 Refresh `btn[8]`)**:
   - Sau khi bấm `btn[19]`, bấm ngay `btn[8]` để làm mới Due list.
   - Nếu số lượng STO đã giảm đủ: **COI LÀ DONE**, bỏ qua hoàn toàn việc quét/xóa 3001, chuyển thẳng sang VL02N.
   - Chỉ khi STO giữ nguyên chưa giảm số lượng, mới rẽ sang 3001 xóa Draft OB nhả tồn.
2. **Khắc phục triệt để hiển thị nút ALL IN ONE**:
   - Tạo đủ 3 nút bấm độc lập: Nút 1 (Teal), Nút 2 (Cam), Nút 3 ALL IN ONE (Royal Blue `RGB(30, 64, 175)`) tại `P17:R19`.
   - Đảm bảo hiển thị tự động khi mở workbook qua `Workbook_Open` và tự phục hồi trong `OB_Check`.

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC63 (15/10/2026)
* **`RC63_HANDOFF.zip`**: SHA-256 `B69601E855CD1DECE016C0D597A2E9F6DFE73E29EB334D972DF79FDBC4AE5980`
* **`261010_OB_TOOL_V1_RC63.xlsm`**: SHA-256 `60EF66BDD0E44BA13E1A635CD6EB255749A5F8990F605A7FFB5254A112108D64`

### OB Tool V1 RC62 (15/10/2026)
* **`RC62_HANDOFF.zip`**: SHA-256 `DB39C4BD0CD87E48DA41213DF0E305BD2B90010CD30BD3295EDAA82681C1C00D`
* **`261010_OB_TOOL_V1_RC62.xlsm`**: SHA-256 `4947174A06C0D0017BD17924803C14699B6305F6A1253C9ADECD66E0A97C6593`
