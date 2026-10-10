# OB TOOL V1 RC43

Canonical XLSM: `261010_OB_TOOL_V1_RC43.xlsm`  
Canonical SHA256: `A406F9265328781F6BBDB2F63F7158666257976D9333ECFBD6CC4445F2FA4F6F`  
Handoff Archive: `RC43_HANDOFF.zip`  
Archive SHA256: `FDCF2FFE3354D6F3037D8415C83D5FB294BC0743928F2C417B11DEFBC6A595CC`

## Key Improvements in RC43
- **Xác thực trực tiếp OB từ VL06O (Direct VL06O Receipt Verification - Nhánh 4007)**:
  - Khắc phục lỗi dừng sau khi tạo thành công OB mới: `VL10D_TARGET_READ_WAIT: OB=...; VL03N; Thieu field LIPS-LFIMG row 0. Can MAP layout VL02N` dẫn tới mất kế hoạch tạo OB.
  - Khi OB được hệ thống tự động nhận diện từ VL06O (`VL06O_AUTO_DISCOVERED`), RC43 trực tiếp xác thực số lượng và thông tin từ bộ nhớ `VL06O_DATA` mà không cần mở VL03N.
  - Chuyển trường `LIPS-MATNR` và `LIPS-LFIMG` sang đọc an toàn (`EDCOptionalTableField`), triệt tiêu lỗi 4205 do layout SAP cá nhân.
- **Tái sử dụng màn hình Due VL10D - Chạy VL10D duy nhất 1 lần (Single VL10D Run)**:
  - Không mở lại màn hình chọn và truy vấn VL10D lần 2 khi đã có sẵn danh sách Due hợp lệ trong bộ nhớ đệm (`mEDCLastDue`).
- **Phản hồi trạng thái thành công rõ ràng trên Nút 1 (User Feedback)**:
  - Cập nhật thông báo rõ ràng trên ô trạng thái: `TAO OB THANH CONG: ... ALLOC; SAN SANG BAM NUT 2; MODE=...` khi Nút 1 tạo xong OB mới.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC43 (10/10/2026)

### 📦 Các tệp đính kèm RC43

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC43_HANDOFF.zip`** | 742,689 bytes | `FDCF2FFE3354D6F3037D8415C83D5FB294BC0743928F2C417B11DEFBC6A595CC` | **Gói bàn giao đầy đủ RC43 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC43.xlsm`** | 769,253 bytes | `A406F9265328781F6BBDB2F63F7158666257976D9333ECFBD6CC4445F2FA4F6F` | Tệp Excel Macro phát hành RC43 (Đã dọn dẹp & Compile) |
| **`VBA_RC42_TO_RC43.diff`** | 23,958 bytes | `7AFD27F647795F57A89A487E37F3FD7440609FF9A79BAC78E66C35867C9C4E7B` | Chi tiết bản vá mã nguồn VBA từ RC42 lên RC43 |
| **`CHANGELOG_RC43.md`** | 4,779 bytes | `401D8D0DA146FA9B1A29A1E0B86BDCA5898EF7D0087CABC10CF14BD5643BCBE0` | Nhật ký thay đổi và giải pháp kỹ thuật RC43 |

### 🚀 Điểm cải tiến cốt lõi trên RC43:
1. **[Xác thực trực tiếp OB từ VL06O - Bỏ mở VL03N gây lỗi Layout]:**  
   - Tự động lấy số lượng và tham chiếu từ bảng dữ liệu VL06O trong RAM, triệt tiêu lỗi thiếu field `LIPS-LFIMG` trên VL03N.
2. **[Tối ưu hóa phiên VL10D - Không mở lặp lại lần 2]:**  
   - Tái sử dụng màn hình Due đang mở để thực hiện Background Delivery, tăng tốc xử lý và loại bỏ các bước thao tác thừa trên SAP GUI.
3. **[Nghiệm thu & VBE Debug Compile PASS 100%]:**  
   - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
   - Gắn macro nút bấm chuẩn xác (`OB_Check` và `OB_Execute`).
   - Bảo toàn trọn vẹn 4 Invariants, WholeBatchGate và nguyên tắc FIFO.

---

## 📜 Lịch sử bản phát hành trước đó

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
