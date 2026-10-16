# OB Tool & Stock Management Releases — Public Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)** và **Stock FG MA BODY**.

---

## 📊 Bản phát hành Công Cụ Quản Trị Tồn Kho: Stock FG MA BODY V3.7 (16/10/2026)

Bản cập nhật hoàn thiện: Tách biệt tuyệt đối giữa Shop BODY và XMĐ, tiêu đề và bảng cột động theo shop, hiển thị rõ ràng mã trước/sau sơn, căn chỉnh UI hoàn hảo.

### 📦 Các tệp đính kèm V3.7

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`261008_Stock_FG_MA_BODY_V3_7.zip`** | 773,300 bytes | `986636358B0ED79DE45CDCF21AA6E99560A6B7B38F4B41E485B631AFCC6DBDDD` | **Gói nén ZIP V3.7 (CHỈ chứa .xlsm và .sha256)** |
| **`261008_Stock_FG_MA_BODY_V3_7.xlsm`** | 863,451 bytes | `55E014E6D66D3324C571BDC7B7DF35E76CB16F2301B5FDC68ABB61C30D67137F` | Tệp Excel Macro V3.7 hoàn chỉnh (Compile PASS 100%) |
| **`RELEASE_NOTES_V3_7.md`** | — | — | Ghi chú kỹ thuật chi tiết bản V3.7 |

### 🚀 Điểm cốt lõi bản V3.7
1. **Căn chỉnh UI thẳng hàng & thẩm mỹ**:
   - Khối `CHỌN MODEL` chuyển sang `N1:O2` căn thẳng tuyệt đối với nút `Btn_SwitchShop` (`N5:O6`) và danh sách Model (`N7:O...`).
   - Cột `M` làm dải phân cách trắng 3 ký tự sạch sẽ giữa bảng dữ liệu và panel điều khiển.
   - Banners `A1:L1`, `A2:L2`, `A10:L10` tự động mở rộng bao trọn 12 cột khi ở shop XMĐ và thu về `A1:K1` khi ở shop BODY.
2. **Tiêu đề và thuật ngữ động theo Shop**:
   - Khi xem **BODY**: Tiêu đề hiển thị `CẢNH BÁO TỒN KHO BODY – MA V3.7`, các cột mang tên `Kho Body`, `Line Body`, `Can build Body`, cảnh báo `Ưu tiên xuất Body`.
   - Khi bấm chuyển **XMĐ**: Tiêu đề tự động chuyển thành `CẢNH BÁO TỒN KHO XMĐ (XE MÁY ĐIỆN) – MA V3.7`, các cột mang tên `ESC trước sơn`, `ESC sau sơn`, `ESC Carset`, cảnh báo `Ưu tiên xuất XMĐ` (không còn chữ Body).
3. **Hiển thị rõ ràng & tách biệt mã trước và sau sơn cho XMĐ**:
   - Cột 1: `FG trước sơn` (Mã trước sơn, e.g. `BIW70011341`)
   - Cột 2: `FG sau sơn` (Mã sau sơn, e.g. `BIW70011351`)
   - Cột 4: `ESC trước sơn` (Tồn của mã trước sơn tại ESC)
   - Cột 5: `ESC sau sơn` (Tồn của mã sau sơn tại ESC)
   - Cột 6: `ESC Carset` (Khả năng lắp ráp xe tổng từ cả trước và sau sơn)
4. **Bảo toàn dữ liệu cập nhật (đổi đuôi AA sang HA)**:
   - Sheet `DATA` là Master Data, tool chỉ đọc và không bao giờ ghi đè. Mọi thay đổi mã vật tư trong sheet `DATA` đều được bảo toàn vĩnh viễn và không bị quay lại mã cũ.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC64 (16/10/2026)

### 📦 Các tệp đính kèm RC64

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC64_HANDOFF.zip`** | 935,918 bytes | `80DDDC1D787B7058A4F530285080E316C853D15E5FB5F67472C9332A1D86DA9F` | **Gói bàn giao đầy đủ RC64 (Streamlined VL10D & ALL IN ONE)** |
| **`261010_OB_TOOL_V1_RC64.xlsm`** | 949,326 bytes | `9A9D9A3875D037C5DF5771932C1622C993780AE41131D52B6ADA90A8FC199FC5` | Tệp Excel Macro phát hành RC64 (Đã tối ưu & Compile PASS 100%) |
| **`VBA_RC63_TO_RC64.diff`** | 10,677 bytes | `679885A193EE90D303B0725D20F0A046D6C8A4F844B79116342110A6619D427A` | Chi tiết bản vá mã nguồn VBA từ RC63 lên RC64 |
| **`CHANGELOG_RC63_TO_RC64.md`** | 4,323 bytes | `7EE3DC060A907B4F5DDF09C6F74059617689981C9B4C1D437A4A3BFB599CDD5E` | Nhật ký thay đổi và giải pháp kỹ thuật RC64 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC63 (15/10/2026)
* **`RC63_HANDOFF.zip`**: SHA-256 `B69601E855CD1DECE016C0D597A2E9F6DFE73E29EB334D972DF79FDBC4AE5980`
* **`261010_OB_TOOL_V1_RC63.xlsm`**: SHA-256 `60EF66BDD0E44BA13E1A635CD6EB255749A5F8990F605A7FFB5254A112108D64`

### OB Tool V1 RC62 (15/10/2026)
* **`RC62_HANDOFF.zip`**: SHA-256 `DB39C4BD0CD87E48DA41213DF0E305BD2B90010CD30BD3295EDAA82681C1C00D`
* **`261010_OB_TOOL_V1_RC62.xlsm`**: SHA-256 `4947174A06C0D0017BD17924803C14699B6305F6A1253C9ADECD66E0A97C6593`
