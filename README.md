# OB TOOL V1 RC46

Canonical XLSM: `261010_OB_TOOL_V1_RC46.xlsm`  
Canonical SHA256: `D3121B65845EBDFDA4AEF1DA84C9BB7B92D7E4E496D8A18918B98FD7C2F5D5A7`  
Handoff Archive: `RC46_HANDOFF.zip`  
Archive SHA256: `AA012032530A6DE17D10AAA5A4C2DA8BD4A1F4798B678E3BFAC577F31A39CDAC`

## Key Improvements in RC46
- **Nguyên tắc "Lượt nào ra lượt đó" - Bảo vệ trọn vẹn Candidate OB có sẵn**:
  - Khắc phục triệt để lỗi đánh dấu `BLOCK_OB: EDC STO/MATERIAL SCOPE BLOCK` lây nhiễm toàn bộ Delivery Order (ví dụ DO `1012699408` chứa 7 mã sẵn có).
  - Phân tách độc lập phạm vi: chỉ chặn từng dòng vi phạm (`BLOCK_ITEM`), tuyệt đối không lây nhiễm sang toàn bộ OB hay các mặt hàng hợp lệ khác.
  - Sau khi Nút 1 quét và gộp VL06O, toàn bộ ứng viên hợp lệ của lượt trước được bảo toàn nguyên vẹn để người dùng bấm Nút 2 (`OB_Execute`) chia lô ngay lập tức.
- **Cơ chế "Tự can và tự tạo lại" (Auto-Cancel 3001 Draft & Immediate Auto-Retry VL10D)**:
  - Khi SAP thiếu hụt số lượng do Plant 3001 chiếm giữ tồn tại 3093, tool tự động tìm và xóa draft OB 3001 (chưa gán SLoc, chưa gán Batch, sát nhất với lượng thiếu).
  - Ngay sau khi xóa thành công giải phóng tồn, hệ thống **tự động gọi tạo lại ngay trên VL10D (`AUTO_RETRY_VL10D`) trong cùng một phiên chạy**, tự động nhận diện DO mới sinh ra và commit vào kế hoạch!
- **Khắc phục lỗi xóa Draft OB trên VL02N**:
  - Chuyển đổi từ phím ảo `Shift+F12` (bị vô hiệu hóa trên Dynpro 4001 gây lỗi `The virtual key is not enabled`) sang mã hàm SAP chuẩn `LOES` trong OKCode field kết hợp toolbar `btn[14]` và menu, bảo đảm 100% xóa sạch draft OB trên SAP GUI.
- **Tải lại toàn diện danh mục hàng trong nhóm (`EDCRefreshGroupOB`)**:
  - Nạp đầy đủ `InputProducts()` khi refresh VL06O thay vì chỉ lọc theo danh sách thiếu hụt, tránh xóa nhầm cache các mã đã có trên sheet `VL06O_DATA`.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC46 (11/10/2026)

### 📦 Các tệp đính kèm RC46

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC46_HANDOFF.zip`** | 763,579 bytes | `AA012032530A6DE17D10AAA5A4C2DA8BD4A1F4798B678E3BFAC577F31A39CDAC` | **Gói bàn giao đầy đủ RC46 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC46.xlsm`** | 788,278 bytes | `D3121B65845EBDFDA4AEF1DA84C9BB7B92D7E4E496D8A18918B98FD7C2F5D5A7` | Tệp Excel Macro phát hành RC46 (Đã dọn dẹp & Compile) |
| **`VBA_RC45_TO_RC46.diff`** | 7,817 bytes | `5A62FB3F30774CF786C6C07605FBA3966D53C5FBD34FEC47F5A24308D47A715C` | Chi tiết bản vá mã nguồn VBA từ RC45 lên RC46 |
| **`CHANGELOG_RC46.md`** | 4,571 bytes | `A4301DF2390FD6445F9E66A4ADF7DC91D48262232E8C85049CA7C3DFA206CB2C` | Nhật ký thay đổi và giải pháp kỹ thuật RC46 |

### 🚀 Điểm cải tiến cốt lõi trên RC46:
1. **[Lượt nào ra lượt đó - Bảo tồn Candidate OB có sẵn]:** Không bị ngộ nhận, block chùm hay xóa nhầm ứng viên OB có sẵn khi có một hoặc vài mã còn thiếu due hoặc chưa tạo được trên VL10D.
2. **[Tự can và tự tạo lại trong cùng phiên]:** Tự động xóa draft OB 3001 chiếm giữ tồn tại 3093 và **lập tức gọi tạo lại ngay trên VL10D**, sinh đủ DO và phân bổ trọn vẹn mà không cần thao tác thủ công.
3. **[Vá lỗi xóa Draft OB bằng LOES]:** Loại bỏ hoàn toàn lỗi `The virtual key is not enabled` trên VL02N.
4. **[Nghiệm thu & VBE Debug Compile PASS 100%]:**  
   - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
   - Gắn macro nút bấm chuẩn xác (`OB_Check` và `OB_Execute`).

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC45 (11/10/2026)
* **`RC45_HANDOFF.zip`**: SHA-256 `06CBFF1A5F26AAA33A8B89893C0C7895FE63BC690BDA7857614AF68B5589FAF8`
* **`261010_OB_TOOL_V1_RC45.xlsm`**: SHA-256 `FB6EBCC36DC181408E0DB6488B79C44E0CB69F5157EC749AFB99A1AB64CC2C8A`
* Chấp nhận từng phần khi tạo DO trên VL10D, theo dõi lượng thiếu hụt do 3001 chiếm giữ tồn, quét xóa thông minh OB nháp 3001.

### OB Tool V1 RC44 (11/10/2026)
* **`RC44_HANDOFF.zip`**: SHA-256 `B38CE54348DDF3D469767E10AB8FD84E440BDC8755FDBF420280614445803C38`
* **`261010_OB_TOOL_V1_RC44.xlsm`**: SHA-256 `FF3873E87F37F05C5D4952EF8AEE22D4B11BE73F4992F200A463F9DAF063AE12`
* Mở khóa bộ lọc nhiều vật tư trên VL10D, bảo toàn số lượng Due trên Classic View, và khớp trọn vẹn lô nguyên theo WholeBatchGate & FIFO.

### OB Tool V1 RC43 (10/10/2026)
* **`RC43_HANDOFF.zip`**: SHA-256 `FDCF2FFE3354D6F3037D8415C83D5FB294BC0743928F2C417B11DEFBC6A595CC`
* **`261010_OB_TOOL_V1_RC43.xlsm`**: SHA-256 `A406F9265328781F6BBDB2F63F7158666257976D9333ECFBD6CC4445F2FA4F6F`
* Trực tiếp xác thực OB từ VL06O, tái sử dụng màn hình Due VL10D chạy 1 lần duy nhất, thông báo phản hồi thành công trên Nút 1.
