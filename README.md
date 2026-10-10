# OB TOOL V1 RC41

Canonical XLSM: `261010_OB_TOOL_V1_RC41.xlsm`  
Canonical SHA256: `6D2B1E08ED2C8CD3C3287420B755FEAF79C1ECBF2C09D281604E7ED178644D3A`  
Handoff Archive: `RC41_HANDOFF.zip`  
Archive SHA256: `C058C5604AB6E1EB880717A711C63A2E50A178667A2E2514E1B65B9332ECE552`

## Key Improvements in RC41
- **Phân bổ STO Due từng phần (Partial STO Due Allocation & Dynamic Create - Nhánh 4007)**:
  - Cho phép chọn STO có số lượng Due nhỏ hơn tổng nhu cầu cần bù (`missing`), thay vì đòi hỏi cứng `>= missing`.
  - Tự động tính số lượng tạo OB: $\text{CreateQty} = \min(\text{missing}, \text{Due})$.
  - Tạo mới OB ngay qua VL10D trong Nút 1, tự động refresh nạp OB vào VL06O, phân bổ đủ lô cho số lượng khả dụng và báo cáo rõ ràng phần còn thiếu do hạn mức STO.
- **Bảo toàn thành quả 100% của Nhánh 3001**:
  - 19/19 OBs thành công tuyệt đối trên dữ liệu thực tế.
  - Vận hành siêu tốc với Fast Save Mode (lưu không cần qua VL03N) và Direct Item Locate (không cuộn trang xuống đáy, không quét dư thừa).
  - Tự động giải phóng khóa ngầm SAP `DO_AUTO` với cơ chế Smart Lock Backoff (2 nhịp retry 6s/5s).

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC41 (10/10/2026)

### 📦 Các tệp đính kèm RC41

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC41_HANDOFF.zip`** | 738,739 bytes | `C058C5604AB6E1EB880717A711C63A2E50A178667A2E2514E1B65B9332ECE552` | **Gói bàn giao đầy đủ RC41 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC41.xlsm`** | 787,464 bytes | `6D2B1E08ED2C8CD3C3287420B755FEAF79C1ECBF2C09D281604E7ED178644D3A` | Tệp Excel Macro phát hành RC41 (Đã dọn dẹp & Compile) |
| **`VBA_RC40_TO_RC41.diff`** | 3,933 bytes | `8264CEC4FA8F31FB686439EE69C4C17BF5EC6622D1332672F3B556474A75958A` | Chi tiết bản vá mã nguồn VBA từ RC40 lên RC41 |
| **`CHANGELOG_RC41.md`** | 4,200 bytes | `70C88BDEDBB5FB676CFC82AEFEC4F3454532EAB323660C803F4CE2DFF284D54F` | Nhật ký thay đổi và giải pháp kỹ thuật RC41 |

### 🚀 Điểm cải tiến cốt lõi trên RC41:
1. **[Partial STO Due Allocation & Dynamic Create - Nhánh 4007]:**  
   - Bỏ điều kiện cứng `candidate(4) >= missing`. Chấp nhận mọi STO còn Due > 0.
   - Tính toán lượng tạo động: $\text{CreateQty} = \min(\text{missing}, \text{Due})$.
   - Khi có STO Due 280 pcs mà nhu cầu là 680 pcs $\rightarrow$ Hệ thống vẫn duyệt tạo OB 280 pcs, nạp vào danh sách và phân bổ ngay lập tức cho 7 batch ($7 \times 40 = 280$), đồng thời báo cáo rõ 400 pcs còn lại thiếu STO.
2. **[Tự động đồng bộ và nạp VL06O sau khi tạo OB]:**  
   - Ngay sau khi VL10D sinh số OB mới, tự động refresh VL06O và nạp OB vào RAM để sẵn sàng cho Nút 2 xử lý điền SLoc 3093 và Batch.
3. **[Nghiệm thu & VBE Debug Compile PASS 100%]:**  
   - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
   - Gắn macro nút bấm chuẩn xác (`OB_Check` và `OB_Execute`).
   - Bảo toàn trọn vẹn 4 Invariants, WholeBatchGate và nguyên tắc FIFO.

---

## 📜 Lịch sử bản phát hành trước đó

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

### OB Tool V1 RC35 R5 (10/10/2026)
* **`RC35_R5_HANDOFF.zip`**: SHA-256 `BBAAD4ABAF3C91C3B428C5CF177CEEC0D8FE94585F070ACBFFC4839CDD122A74`
* **`261010_OB_TOOL_V1_RC35_R5.xlsm`**: SHA-256 `98EE42AD70097067436EFD674AE069F1B9C2A72D66BEC3D4856394E2A833FF53`
* Khắc phục lỗi Active Error Handler (`Resume RetryWait`) và chốt chặn `EDCEnsureGroupInSnapshot`.
