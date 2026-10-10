# OB TOOL V1 RC44

Canonical XLSM: `261010_OB_TOOL_V1_RC44.xlsm`  
Canonical SHA256: `FF3873E87F37F05C5D4952EF8AEE22D4B11BE73F4992F200A463F9DAF063AE12`  
Handoff Archive: `RC44_HANDOFF.zip`  
Archive SHA256: `B38CE54348DDF3D469767E10AB8FD84E440BDC8755FDBF420280614445803C38`

## Key Improvements in RC44
- **Mở khóa lọc danh sách nhiều vật tư trên VL10D (`EDCSetGroupMaterialFilter`)**:
  - Khắc phục lỗi lặp lại `-2147217243; Cannot verify pasted Material selection table.` khi nhập $> 1$ mã vật tư.
  - Chuẩn hóa: Sau khi dán clipboard (`btn[24]`), tool bấm F8 (`btn[8]`) xác nhận và đóng popup, quay lại màn hình chính điền Plant 4007, Supplying Plant 3093 và nhấn F8 truy vấn VL10D.
- **Bảo toàn số lượng Due đọc được trên Classic View VL10D (`EDC_ClassicDue`)**:
  - Khắc phục lỗi `eligible rows = 0` khi STO thực tế có Due (ví dụ 399 pcs) do xô lệch toạ độ cột `Open` làm triệt tiêu số lượng.
  - Thêm điều kiện an toàn: Chỉ giới hạn `openQty` khi `openQty > 0`. Bổ sung quét lân cận $\pm 6$ cột nếu ô số lượng bị dịch toạ độ.
- **Thuật toán tạo OB khớp trọn vẹn theo WholeBatchGate & FIFO (`EDCWholeBatchFitQty`)**:
  - Khi $\text{STO Due} < \text{need}$, tool tự động duyệt danh sách lô FIFO để gom các lô nguyên có tổng $\le \text{STO Due}$ (ví dụ: Due = 399 $\implies$ gom 2 lô $161 + 200 = 361$ pcs).
  - Đặt số lượng tạo OB đúng bằng 361 pcs, giúp Nút 2 mở VL02N chia vừa khít 2 lô nguyên, không bị dư thừa 38 pcs, bảo toàn 100% WholeBatchGate!

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC44 (11/10/2026)

### 📦 Các tệp đính kèm RC44

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC44_HANDOFF.zip`** | 741,894 bytes | `B38CE54348DDF3D469767E10AB8FD84E440BDC8755FDBF420280614445803C38` | **Gói bàn giao đầy đủ RC44 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC44.xlsm`** | 769,211 bytes | `FF3873E87F37F05C5D4952EF8AEE22D4B11BE73F4992F200A463F9DAF063AE12` | Tệp Excel Macro phát hành RC44 (Đã dọn dẹp & Compile) |
| **`VBA_RC43_TO_RC44.diff`** | 15,156 bytes | `973F6B9006D3BB13C1C65C09D30FAECB5D00CEBD0DE8A2A93CA91F7994C4B3ED` | Chi tiết bản vá mã nguồn VBA từ RC43 lên RC44 |
| **`CHANGELOG_RC44.md`** | 5,310 bytes | `2B6C76E1ADEC33EF84604EA2A3F97371F646C9A254B8F09C2CAA2CBAD05D9394` | Nhật ký thay đổi và giải pháp kỹ thuật RC44 |

### 🚀 Điểm cải tiến cốt lõi trên RC44:
1. **[Mở khóa Multiple Material Filter trên VL10D]:** Dán danh sách vật tư và bấm F8 đóng popup mượt mà, không bị văng lỗi kiểm tra bảng.
2. **[Bảo toàn số lượng Due trên Classic View]:** Nhận diện chính xác Due 399 pcs, không bị lỗi cột Open ép số lượng về 0.
3. **[Khớp số lượng tạo OB theo WholeBatchGate & FIFO]:** Tính toán số lượng tạo OB bằng tổng các lô nguyên khả dụng (361 pcs), giúp Nút 2 gán lô vừa khít 100%.
4. **[Nghiệm thu & VBE Debug Compile PASS 100%]:**  
   - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
   - Gắn macro nút bấm chuẩn xác (`OB_Check` và `OB_Execute`).

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC43 (10/10/2026)
* **`RC43_HANDOFF.zip`**: SHA-256 `FDCF2FFE3354D6F3037D8415C83D5FB294BC0743928F2C417B11DEFBC6A595CC`
* **`261010_OB_TOOL_V1_RC43.xlsm`**: SHA-256 `A406F9265328781F6BBDB2F63F7158666257976D9333ECFBD6CC4445F2FA4F6F`
* Trực tiếp xác thực OB từ VL06O (bỏ VL03N gây lỗi layout), tái sử dụng màn hình Due VL10D chạy 1 lần duy nhất, thông báo phản hồi thành công trên Nút 1.

### OB Tool V1 RC42 (10/10/2026)
* **`RC42_HANDOFF.zip`**: SHA-256 `957A35CDA3BE65AE8F2A5C889F54D42F99110C2B6FCFFE828ACB46234BA70155`
* **`261010_OB_TOOL_V1_RC42.xlsm`**: SHA-256 `03F4B9AD3120901197B49D8A3DE6EE1A9D6C517E9659A3FB4D88A2E82A9432ED`
* Mở khóa tạo Delivery VL10D ở chế độ APPLY (Nhánh 4007), chuẩn hóa hiển thị ngày tháng và UI sheet LOG.

### OB Tool V1 RC41 (10/10/2026)
* **`RC41_HANDOFF.zip`**: SHA-256 `C058C5604AB6E1EB880717A711C63A2E50A178667A2E2514E1B65B9332ECE552`
* **`261010_OB_TOOL_V1_RC41.xlsm`**: SHA-256 `6D2B1E08ED2C8CD3C3287420B755FEAF79C1ECBF2C09D281604E7ED178644D3A`
* Hỗ trợ phân bổ STO Due từng phần $\text{CreateQty} = \min(\text{missing}, \text{Due})$, tự động lập kế hoạch tạo 280 pcs cho nhánh 4007.

### OB Tool V1 RC40 (10/10/2026)
* **`RC40_HANDOFF.zip`**: SHA-256 `6E3B0C9B59815808EFEF63B5AFC1843C70708623F8947069C4BE9708A5E56574`
* **`261010_OB_TOOL_V1_RC40.xlsm`**: SHA-256 `191D37DAD2E884A5B26026B9332F2BF01CB0C64C8427D9985B5F267DEFAAF4C7`
* Bổ sung Smart Lock Backoff (2 nhịp retry 6s/5s) cho hàng đợi deferred bị kẹt `DO_AUTO`, mở khóa màn hình chọn ban đầu VL10D cho nhánh 4007, tự động dọn dẹp cờ pending cũ.

### OB Tool V1 RC39 (10/10/2026)
* **`RC39_HANDOFF.zip`**: SHA-256 `36BAF1F6154731D8B363D2FEB462746042E32175DA6D18C3ED2870B27A9BAF7E`
* **`261010_OB_TOOL_V1_RC39.xlsm`**: SHA-256 `1C374619A87D3A888DC8CA82E307FE8784747DE63AED9AE3FB6ABB30E3A27E32`
* Dừng quét ngay khi hết item (Early Stop), tối ưu hóa xóa item và định vị dòng trực tiếp không cuộn trang xuống đáy.

### OB Tool V1 RC38 (10/10/2026)
* **`RC38_HANDOFF.zip`**: SHA-256 `0864D60669766357A17001A4C52BEDD8D2F42453E4F11315D58CE1A532D411D6`
* **`261010_OB_TOOL_V1_RC38.xlsm`**: SHA-256 `6100917C4BAD6F68A3DED74A2E74CFEBA7A9006790458055DDD500D61CA51DC8`
* Khôi phục Fast Save từ RC34 (bỏ bắt buộc VL03N), khai thông đọc VL10D live Due Qty cho nhánh 4007 STO, hỗ trợ VL02N session reuse.

### OB Tool V1 RC37 (10/10/2026)
* **`RC37_HANDOFF.zip`**: SHA-256 `C762540D94D7366F13AE4CE650C2BB1DF5314947BE3A22B598A42EFC560E9EE5`
* **`261010_OB_TOOL_V1_RC37.xlsm`**: SHA-256 `E66F327F3CC98A830E6C9037C735D5A7FE2FAC2FCBC1950DB26DC0840D84FB33`
* Vá lỗi bẫy rỗng VL06O cho mọi Plant, dọn dẹp lịch sử LOG/LEDGER, tự động bắt lại OB từ VL06O, cảnh báo tồn kho khóa QI/QU/QT.

### OB Tool V1 RC36 (10/10/2026)
* **`RC36_HANDOFF.zip`**: SHA-256 `35D034EC709FAB0724FDD50C3E0C525D8B5D362F7C3486D2F6929EE02D14B6B7`
* **`261010_OB_TOOL_V1_RC36.xlsm`**: SHA-256 `5BA00C8A1D6744A9CB2291A2DB924E26DA890B16EF1F43143796761EB1D609D7`
* Tự động bắt số OB từ VL06O theo STO + Material + Người tạo (ERNAM).
