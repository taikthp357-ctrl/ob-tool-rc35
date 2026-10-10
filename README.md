# OB TOOL V1 RC42

Canonical XLSM: `261010_OB_TOOL_V1_RC42.xlsm`  
Canonical SHA256: `03F4B9AD3120901197B49D8A3DE6EE1A9D6C517E9659A3FB4D88A2E82A9432ED`  
Handoff Archive: `RC42_HANDOFF.zip`  
Archive SHA256: `957A35CDA3BE65AE8F2A5C889F54D42F99110C2B6FCFFE828ACB46234BA70155`

## Key Improvements in RC42
- **Mở khóa tạo Delivery VL10D ở chế độ APPLY (Nhánh 4007)**:
  - Khắc phục lỗi dừng trước khi bấm nút tạo Delivery: `VL10D_GROUP_WAIT: GROUP_PREPARE; -2147217243; Group due missing Material/STO/Item; no Create`.
  - Nguyên nhân do cột Supplying Plant (`RESWK`) không hiển thị trên danh sách Classic View của VL10D khiến `a(10)` rỗng và trượt điều kiện lọc nguồn 3093.
  - RC42 gán trực tiếp nguồn 3093 và nới lỏng điều kiện kiểm tra cột nguồn thành `(CStr(a(10)) = "" Or CStr(a(10)) = "3093")`, cho phép tiến thẳng tới bước gửi lệnh tạo Delivery sang SAP.
  - Bổ sung kiểm tra tooltip hỗ trợ cả nhãn `"background"` và `"hintergrund"` cho nút lệnh tạo Delivery nền.
- **Chuẩn hóa giao diện và định dạng ngày tháng sheet LOG (UI & Formatting)**:
  - Khắc phục lỗi cột `TIMESTAMP` hiển thị số thực thô `46305,95369`: Đặt chuẩn `NumberFormat = "m/d/yy h:mm AM/PM"` và `.Value = Now`, đảm bảo mọi dòng ghi log luôn hiển thị đúng ngày giờ (ví dụ: `10/10/26 10:54 PM`).
  - Mở rộng độ rộng các cột sheet LOG để triệt tiêu hoàn toàn hiện tượng tràn chữ, đè cột:
    - Cột B (`RUN ID`): Mở rộng từ 14 lên **20**, căn giữa.
    - Cột C (`STEP`): Mở rộng từ 18 lên **30**, căn trái.
    - Cột D (`OB / REF`): Mở rộng lên **18**, căn giữa.
    - Cột F (`MESSAGE`): Độ rộng **90**, căn trái.
- **Đồng bộ hóa sheet CONFIG & LEDGER**:
  - Cập nhật ô `CONFIG!B2` thành `OB_V1_RC42` (khắc phục nhãn cũ `OB_V1_RC34`).
  - Tối ưu độ rộng cột sheet `LEDGER` (`UPDATED AT` = 20, `MESSAGE` = 35).

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC42 (10/10/2026)

### 📦 Các tệp đính kèm RC42

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC42_HANDOFF.zip`** | 731,125 bytes | `957A35CDA3BE65AE8F2A5C889F54D42F99110C2B6FCFFE828ACB46234BA70155` | **Gói bàn giao đầy đủ RC42 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC42.xlsm`** | 762,168 bytes | `03F4B9AD3120901197B49D8A3DE6EE1A9D6C517E9659A3FB4D88A2E82A9432ED` | Tệp Excel Macro phát hành RC42 (Đã dọn dẹp & Compile) |
| **`VBA_RC41_TO_RC42.diff`** | 4,331 bytes | `4348CEBC908E931DE040563F87B772D415C6252BE1F02763F9963A105DD8E3A2` | Chi tiết bản vá mã nguồn VBA từ RC41 lên RC42 |
| **`CHANGELOG_RC42.md`** | 5,249 bytes | `9DF3002466DF71FF57676943F4B96EFFC272A400A542DFB8649204C1FDCDEF47` | Nhật ký thay đổi và giải pháp kỹ thuật RC42 |

### 🚀 Điểm cải tiến cốt lõi trên RC42:
1. **[Mở khóa tạo Delivery VL10D ở chế độ APPLY - Nhánh 4007]:**  
   - Gỡ bỏ điểm nghẽn `GROUP_PREPARE: Group due missing Material/STO/Item; no Create`.
   - Nhận diện đúng STO Due trên Classic View của VL10D và gửi lệnh tạo Delivery sang SAP.
2. **[Chuẩn hóa hiển thị ngày tháng và UI sheet LOG]:**  
   - Cột A (`TIMESTAMP`) luôn hiển thị định dạng ngày giờ chuẩn (`10/10/26 10:54 PM`), không còn bị hiển thị số thực thô.
   - Cột B (`RUN ID`) và Cột C (`STEP`) được mở rộng độ rộng, căn lề khoa học, chống tràn chữ.
3. **[Nghiệm thu & VBE Debug Compile PASS 100%]:**  
   - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
   - Gắn macro nút bấm chuẩn xác (`OB_Check` và `OB_Execute`).
   - Bảo toàn trọn vẹn 4 Invariants, WholeBatchGate và nguyên tắc FIFO.

---

## 📜 Lịch sử bản phát hành trước đó

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
