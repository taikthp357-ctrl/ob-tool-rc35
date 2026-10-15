# OB TOOL V1 RC64 (Streamlined 4007 STO Due & Permanent ALL IN ONE Button)

Canonical XLSM: `261010_OB_TOOL_V1_RC64.xlsm`  
Canonical SHA256: `9A9D9A3875D037C5DF5771932C1622C993780AE41131D52B6ADA90A8FC199FC5`  
Handoff Archive: `RC64_HANDOFF.zip`  
Archive SHA256: `80DDDC1D787B7058A4F530285080E316C853D15E5FB5F67472C9332A1D86DA9F`

## 🌟 Điểm Cải Tiến Đột Phá Trên RC64

### 1. Cơ chế Đánh giá STO Due trực tiếp trên VL10D theo VBS tham chiếu
- **Xác thực trực tiếp mức giảm số lượng STO qua Refresh F8 (`btn[8]`)**:
  * Khi tạo OB trên VL10D, tool bấm `btn[19]`, ngay sau đó bấm `btn[8]` để làm mới danh sách Due.
  * Đọc lại trực tiếp số lượng Out/Open của STO ngay trên màn hình hiện hành:
    `consumedQty = initialDue - remainingDue`
  * **Nếu STO đã giảm đủ số lượng yêu cầu (`consumedQty >= requestedQty`)**: **ĐƯỢC COI LÀ DONE!**
    Tool bỏ qua 100% bước quét và xóa nháp 3001 (`RELEASE_3001_BYPASS`), rút ngắn tối đa thời gian thực thi và loại bỏ các bước kiểm tra thừa.
  * **Chỉ khi số lượng STO sau khi tạo vẫn giữ nguyên** (chứng tỏ 3001 giữ tồn khả dụng / thiếu ATP), tool mới rẽ nhánh sang 3001 để xóa Draft OB nhả tồn, sau đó quay lại VL10D refresh và tạo bù phần thiếu.

### 2. Khắc phục triệt để lỗi nút "ALL IN ONE" không hiện khi mở file
- **3 Nút bấm đồ họa chuẩn độc lập trên sheet `INPUT`**:
  * Nút 1: `1. KIỂM TRA & LẬP KẾ HOẠCH` (`P9:R11`) - Teal `RGB(15, 118, 110)` -> `OB_Check`
  * Nút 2: `2. THỰC HIỆN KẾ HOẠCH` (`P13:R15`) - Cam `RGB(166, 94, 15)` -> `OB_Execute`
  * Nút 3: `ALL IN ONE - TỰ ĐỘNG CHẠY TOÀN BỘ (1 + 2)` (`P17:R19`) - Xanh Hoàng Gia `RGB(30, 64, 175)` -> `OB_RunAll`
- **Luôn tự động hiển thị ổn định**: Cập nhật vòng lặp `For i = 0 To 2` trong `OB_SetupButtons`, kết nối trong `Workbook_Open` và kiểm tra phòng vệ trong `OB_Check` đảm bảo nút bấm không bao giờ bị biến mất.
