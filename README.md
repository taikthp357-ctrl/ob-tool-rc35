# OB TOOL V1 RC47

Canonical XLSM: `261010_OB_TOOL_V1_RC47.xlsm`  
Canonical SHA256: `B90CA4209A13806027C2234D32608E2F2C583F76074979BC048C3C450C3604EC`  
Handoff Archive: `RC47_HANDOFF.zip`  
Archive SHA256: `A9F845EC80F4DB7D0C03B78E9E5060964A3559E90A0838BDE2EAD022D002ACA8`

## Key Improvements in RC47
- **Cam kết & Kích hoạt Nút 2 Tức thì (`Atomic Commit & mChecked = True`)**:
  - Khi SAP tạo DO cho một phần danh sách mặt hàng (ví dụ 3/4 mã), các mặt hàng thành công được commit và ghi nhận biên nhận ngay lập tức.
  - Thiết lập `mChecked = True` tức thì, bảo đảm người dùng bấm Nút 2 (`OB_Execute`) là chạy chia lô và lưu ngay mà không bao giờ bị báo "chưa có kế hoạch"!
- **Bảo tồn Tuyệt đối Phân bổ các Mặt hàng đã có OB (`EDCPlanOne`)**:
  - Khi một mã đã có DO với một phần số lượng (`capacity > 0`), hệ thống không ném lỗi `OB_ALREADY_CREATED_WAIT`, cho phép phân bổ trọn vẹn số lượng đó và ghi nhận vào kế hoạch Nút 2.
- **Cách ly Hoàn toàn `mEDCWaitByMat` khỏi `freshBlocked` (`OB_Execute`)**:
  - Gỡ bỏ việc gộp danh sách chờ tạo VL10D vào `freshBlocked`, ngăn chặn tình trạng các mã hợp lệ bị loại khỏi nhóm thực thi của Delivery Order.
- **Bất biến Bảo vệ Hàng trên Delivery Order (`RemoveUnplannedItems`)**:
  - Áp dụng nguyên tắc bất khả xâm phạm: Mọi dòng hàng trên DO có mã vật tư nằm trong danh mục `InputProducts()` tuyệt đối không bị xóa! Chấm dứt triệt để lỗi xóa nhầm các mã khác trên DO.
- **Nâng cấp Cơ chế Xóa OB 3001 & Bắt lỗi Chặt chẽ (`SAPX_DeleteDraftOB`)**:
  - Quét tìm lệnh Delete từ menu `Outbound Delivery` và OKCode `DELE`.
  - Bắt buộc kiểm tra `sType <> "E"`, không bao giờ ngộ nhận xóa thành công khi SAP trả về lỗi, và loại bỏ hoàn toàn việc gọi xóa lặp lại trong khối lỗi.
- **Xử lý An toàn Mảng Kế hoạch (`EDCEnsureGroupInSnapshot`)**:
  - Kiểm tra `IsArray(planned)` loại bỏ lỗi `Type mismatch (13)`, tự động tách các mã có số lượng bằng 0 khỏi nhóm commit mà không làm gián đoạn các mã đã tạo thành công.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC47 (11/10/2026)

### 📦 Các tệp đính kèm RC47

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC47_HANDOFF.zip`** | 765,842 bytes | `A9F845EC80F4DB7D0C03B78E9E5060964A3559E90A0838BDE2EAD022D002ACA8` | **Gói bàn giao đầy đủ RC47 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC47.xlsm`** | 792,283 bytes | `B90CA4209A13806027C2234D32608E2F2C583F76074979BC048C3C450C3604EC` | Tệp Excel Macro phát hành RC47 (Đã dọn dẹp & Compile) |
| **`VBA_RC46_TO_RC47.diff`** | 10,812 bytes | `7DB5C925A1AE7C7E13FD0DCE75CBA6EC9E3AE80164A496033D352F942475143F` | Chi tiết bản vá mã nguồn VBA từ RC46 lên RC47 |
| **`CHANGELOG_RC47.md`** | 6,854 bytes | `2C240A2F3E75A9C0CEE83D0CE8C18C8B72D71BE20129B854BA6E56CEFBD58804` | Nhật ký thay đổi và giải pháp kỹ thuật RC47 |

### 🚀 Điểm cải tiến cốt lõi trên RC47:
1. **[Lần 1 có kế hoạch $ightarrow$ Bấm Nút 2 chạy ngay]:** Cam kết từng phần và đặt `mChecked = True`, người dùng bấm Nút 2 là thực thi mượt mà.
2. **[Không xóa nhầm mã trên DO]:** Bảo vệ 100% các dòng hàng thuộc `InputProducts()` trong `RemoveUnplannedItems`, điền đủ SLoc/Batch cho tất cả các mã trên DO.
3. **[Không đóng băng mã đã có OB]:** `EDCPlanOne` và `freshBlocked` bảo tồn trọn vẹn số lượng OB đã tạo để phân bổ ngay.
4. **[Chuẩn hóa xóa OB 3001]:** Menu delete + OKCode `DELE`, chặn đứng lỗi lặp và kiểm tra nghiêm ngặt kết quả.
5. **[Nghiệm thu & VBE Debug Compile PASS 100%]:**  
   - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
   - Gắn macro nút bấm chuẩn xác (`OB_Check` và `OB_Execute`).

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC46 (11/10/2026)
* **`RC46_HANDOFF.zip`**: SHA-256 `AA012032530A6DE17D10AAA5A4C2DA8BD4A1F4798B678E3BFAC577F31A39CDAC`
* **`261010_OB_TOOL_V1_RC46.xlsm`**: SHA-256 `D3121B65845EBDFDA4AEF1DA84C9BB7B92D7E4E496D8A18918B98FD7C2F5D5A7`

### OB Tool V1 RC45 (11/10/2026)
* **`RC45_HANDOFF.zip`**: SHA-256 `06CBFF1A5F26AAA33A8B89893C0C7895FE63BC690BDA7857614AF68B5589FAF8`
* **`261010_OB_TOOL_V1_RC45.xlsm`**: SHA-256 `FB6EBCC36DC181408E0DB6488B79C44E0CB69F5157EC749AFB99A1AB64CC2C8A`

### OB Tool V1 RC44 (11/10/2026)
* **`RC44_HANDOFF.zip`**: SHA-256 `B38CE54348DDF3D469767E10AB8FD84E440BDC8755FDBF420280614445803C38`
* **`261010_OB_TOOL_V1_RC44.xlsm`**: SHA-256 `FF3873E87F37F05C5D4952EF8AEE22D4B11BE73F4992F200A463F9DAF063AE12`
