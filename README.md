# OB TOOL V1 RC45

Canonical XLSM: `261010_OB_TOOL_V1_RC45.xlsm`  
Canonical SHA256: `FB6EBCC36DC181408E0DB6488B79C44E0CB69F5157EC749AFB99A1AB64CC2C8A`  
Handoff Archive: `RC45_HANDOFF.zip`  
Archive SHA256: `06CBFF1A5F26AAA33A8B89893C0C7895FE63BC690BDA7857614AF68B5589FAF8`

## Key Improvements in RC45
- **Chấp nhận từng phần khi tạo DO trên VL10D (`EDCDiscoverReceiptsFromVL06O` & `EDCVerifyGroup`)**:
  - Khắc phục triệt để lỗi fatal `GROUP_RECEIPT_MISSING` và `GROUP_OB_QTY_MISMATCH` khi SAP chỉ tạo được một phần danh sách vật tư (ví dụ 7/8 mã) hoặc tạo thiếu số lượng (ví dụ cần 361 nhưng SAP chỉ tạo 200).
  - Tự động tiếp nhận, cập nhật số lượng thực tế đã tạo cho các dòng hàng thành công, commit giải phóng pending và kích hoạt Nút 2 (`OB_Execute`) chia lô ngay lập tức!
- **Theo dõi lượng thiếu hụt do 3001 chiếm giữ tồn (`deficits`)**:
  - Ghi nhận chính xác số lượng thiếu hụt cho các mã bị SAP bỏ qua hoặc tạo thiếu: `Deficit = Planned_Qty - Created_Qty`.
  - Cập nhật ghi chú trực quan trên Sheet `INPUT` cột D: `WAIT_STO_DUE_4007: Thieu ... pcs do 3001 giu ton`.
- **Tự động quét và xóa OB nháp 3001 bằng `Shift + F12` (`SAPX_FindAndReleaseDraft3001OB`)**:
  - Sau khi DO 4007 đã tạo xong và biết chính xác lượng thiếu, tool mở VL06O quét danh sách OB của Plant 3001 tại kho xuất 3093.
  - Áp dụng bộ lọc an toàn tuyệt đối: Chỉ nhắm vào các OB nháp có `SLoc = ""` và `Batch = ""` (chưa điền kho, chưa gán lô, GM status = A).
  - Thuật toán chọn tối ưu: Tìm OB có số lượng gần sát nhất với lượng thiếu ($\min |OB\_Qty - Deficit|$).
  - Mở `VL02N`, gửi lệnh `Shift + F12` (`sendVKey 24`), tự động xác nhận popup xóa để giải phóng lượng tồn bị chiếm dụng tại 3093.
  - Cập nhật ghi chú trên Sheet `INPUT`: `DA XOA OB 3001: [OB] ([Qty] pcs) DE GIAI PHONG TON; NHAN NUT 1 DE TAO LAI VL10D`. Người dùng chỉ cần nhấn lại Nút 1 để tạo tiếp phần còn lại!

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC45 (11/10/2026)

### 📦 Các tệp đính kèm RC45

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC45_HANDOFF.zip`** | 761,744 bytes | `06CBFF1A5F26AAA33A8B89893C0C7895FE63BC690BDA7857614AF68B5589FAF8` | **Gói bàn giao đầy đủ RC45 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC45.xlsm`** | 783,972 bytes | `FB6EBCC36DC181408E0DB6488B79C44E0CB69F5157EC749AFB99A1AB64CC2C8A` | Tệp Excel Macro phát hành RC45 (Đã dọn dẹp & Compile) |
| **`VBA_RC44_TO_RC45.diff`** | 17,504 bytes | `470EF7F3AD925D3D0AED92F2CF901C4F3E1CB7A4B206CCBDA0FD30B2820B39D5` | Chi tiết bản vá mã nguồn VBA từ RC44 lên RC45 |
| **`CHANGELOG_RC45.md`** | 6,023 bytes | `651BEF4B7E72FC8B137E096DAB3DFFF11106D29A7A4965060605C1BF58EF0B7E` | Nhật ký thay đổi và giải pháp kỹ thuật RC45 |

### 🚀 Điểm cải tiến cốt lõi trên RC45:
1. **[Chấp nhận từng phần khi tạo DO trên VL10D]:** Nhận diện và tiếp nhận toàn bộ các dòng hàng SAP đã tạo (7/8 mã hoặc partial qty), không bị hủy toàn bộ mẻ hay văng lỗi.
2. **[Kích hoạt ngay Nút 2 cho các mã đã tạo]:** Commit dữ liệu ngay cho các mã thành công, sẵn sàng chia lô và gán SLoc 3093.
3. **[Quét xóa thông minh OB nháp 3001 bằng Shift+F12]:** Tự động tìm OB nháp chưa có SLoc/Batch gần sát nhất với lượng thiếu để xóa và giải phóng tồn ATP.
4. **[Nghiệm thu & VBE Debug Compile PASS 100%]:**  
   - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
   - Gắn macro nút bấm chuẩn xác (`OB_Check` và `OB_Execute`).

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC44 (11/10/2026)
* **`RC44_HANDOFF.zip`**: SHA-256 `B38CE54348DDF3D469767E10AB8FD84E440BDC8755FDBF420280614445803C38`
* **`261010_OB_TOOL_V1_RC44.xlsm`**: SHA-256 `FF3873E87F37F05C5D4952EF8AEE22D4B11BE73F4992F200A463F9DAF063AE12`
* Mở khóa bộ lọc nhiều vật tư trên VL10D, bảo toàn số lượng Due trên Classic View, và khớp trọn vẹn lô nguyên theo WholeBatchGate & FIFO.

### OB Tool V1 RC43 (10/10/2026)
* **`RC43_HANDOFF.zip`**: SHA-256 `FDCF2FFE3354D6F3037D8415C83D5FB294BC0743928F2C417B11DEFBC6A595CC`
* **`261010_OB_TOOL_V1_RC43.xlsm`**: SHA-256 `A406F9265328781F6BBDB2F63F7158666257976D9333ECFBD6CC4445F2FA4F6F`
* Trực tiếp xác thực OB từ VL06O (bỏ VL03N gây lỗi layout), tái sử dụng màn hình Due VL10D chạy 1 lần duy nhất, thông báo phản hồi thành công trên Nút 1.

### OB Tool V1 RC42 (10/10/2026)
* **`RC42_HANDOFF.zip`**: SHA-256 `957A35CDA3BE65AE8F2A5C889F54D42F99110C2B6FCFFE828ACB46234BA70155`
* **`261010_OB_TOOL_V1_RC42.xlsm`**: SHA-256 `03F4B9AD3120901197B49D8A3DE6EE1A9D6C517E9659A3FB4D88A2E82A9432ED`
* Mở khóa tạo Delivery VL10D ở chế độ APPLY (Nhánh 4007), chuẩn hóa hiển thị ngày tháng và UI sheet LOG.
