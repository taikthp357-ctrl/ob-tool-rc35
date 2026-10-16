# CHANGELOG - OB TOOL V1 RC64

## 1. Tổng quan Bản phát hành RC64
Bản phát hành RC64 tập trung vào 2 mục tiêu kỹ thuật cốt lõi:
1. **Tối ưu hóa quy trình kiểm tra tạo OB nhánh 4007 trên VL10D theo tài liệu tham chiếu `vl10d1(1).vbs`**:
   - Tận dụng cơ chế refresh F8 (`btn[8]`) trực tiếp trên màn hình Due List của VL10D.
   - Kiểm tra mức giảm số lượng Out/Open của STO ngay trên màn hình VL10D:
     * Nếu số lượng STO đã giảm đủ theo số lượng yêu cầu (`consumedQty >= requestedQty`): **COI LÀ DONE!** Bỏ qua hoàn toàn bước quét và xóa nháp 3001, rút ngắn tối đa thời gian chờ và tránh các thao tác kiểm tra rà soát thừa.
     * Chỉ khi số lượng STO sau khi tạo và refresh vẫn giữ nguyên (hoặc giảm không đủ, chứng tỏ 3001 đang giữ tồn khả dụng / thiếu ATP), tool mới rẽ nhánh sang 3001 để xóa Draft OB nhả tồn, sau đó quay lại VL10D refresh và tạo bù đúng phần thiếu.
2. **Khắc phục triệt để lỗi nút "ALL IN ONE" không hiện khi mở file**:
   - Hoàn thiện hàm `OB_SetupButtons` trong `ModuleMain.bas` tạo đủ cả 3 nút bấm độc lập:
     * Nút 1: `1. KIỂM TRA & LẬP KẾ HOẠCH` (Rows 9-11: `P9:R11`) - Màu Teal `RGB(15, 118, 110)`.
     * Nút 2: `2. THỰC HIỆN KẾ HOẠCH` (Rows 13-15: `P13:R15`) - Màu Cam `RGB(166, 94, 15)`.
     * Nút 3: `ALL IN ONE - TỰ ĐỘNG CHẠY TOÀN BỘ (1 + 2)` (Rows 17-19: `P17:R19`) - Màu Xanh Hoàng Gia `RGB(30, 64, 175)`.
   - Kết nối tự động trong sự kiện `Workbook_Open` và kiểm tra phòng thủ trong `OB_Check` để đảm bảo nút luôn hiển thị ổn định.

---

## 2. Chi tiết Kỹ thuật Các Thay Đổi

### 2.1. Cơ chế Đánh giá STO Due trực tiếp trên VL10D (`ModuleSAP.bas` & `ModuleMain.bas`)
- Trong `ModuleSAP.bas`:
  * Cập nhật `EDC4007_CreateGroup`:
    - Sau khi nhấn `btn[19]` tạo Delivery, thực hiện ngay lệnh Refresh `btn[8]` trực tiếp trên Due List (`EDC4007_RefreshDueScreen`).
    - Gọi hàm mới `EDC4007_EvaluateDueDecrease(sto)` để đọc lại danh sách Due hiện hành (`EDC4007_ReadDueCurrentScreen`) mà không cần thoát ra `/nVL10D`.
    - Tính toán chính xác:
      `consumedQty = initialDue - remainingDue`
      - Nếu `consumedQty >= requestedQty` -> Ghi nhận `DONE` (deficit = 0).
      - Nếu `consumedQty < requestedQty` -> Ghi nhận `DEFICIT = requestedQty - consumedQty`.
  * Cung cấp hàm công khai `EDC4007_GetDueDeficits() As Object`.
- Trong `ModuleMain.bas`:
  * Trong `EDCCreateGroup`:
    - Đọc `dueDeficits` từ `ModuleSAP.EDC4007_GetDueDeficits()`.
    - Nếu tất cả các mã đều đã giảm đủ số lượng (`deficits.count = 0`):
      Kích hoạt `RELEASE_3001_BYPASS`, **bỏ qua 100% bước gọi `EDCReleaseDrafts3001`**, không can thiệp vào bất kỳ OB nào ở 3001.
    - Chỉ kích hoạt `EDCReleaseDrafts3001` khi thực sự có mã bị kẹt tồn (`deficits.count > 0`).

### 2.2. Khắc phục Hiển thị Nút Bấm Sheet INPUT
- Trong `ModuleMain.bas` (`OB_SetupButtons`):
  * Mở rộng mảng cấu hình nút bấm lên 3 phần tử (`For i = 0 To 2`):
    - `actions = Array("OB_Check", "OB_Execute", "OB_RunAll")`
    - `rows = Array(9, 13, 17)`
  * Nút ALL IN ONE được đặt tại `P17:R19`, chữ trắng, nền Royal Blue, căn giữa chuẩn đồ họa.
  * Trong `OB_Check`: Tự động phục hồi nút bấm nếu phát hiện thiếu shape `OBV1__2`.

---

## 3. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: PASS 100% (`before=True`, `after=False`).
- **Shape Macro Bindings**:
  * `INPUT!OBV1__0` -> `'261010_OB_TOOL_V1_RC64.xlsm'!OB_Check`
  * `INPUT!OBV1__1` -> `'261010_OB_TOOL_V1_RC64.xlsm'!OB_Execute`
  * `INPUT!OBV1__2` -> `'261010_OB_TOOL_V1_RC64.xlsm'!OB_RunAll`
- **Version Labels**:
  * `INPUT!A1` = `"OB TOOL V1 RC64"`
  * `CONFIG!B2` = `"RC64"`
  * `INPUT!R22` = `"RC64 ALL-IN-ONE & STREAMLINED 4007 STO DUE PIPELINE"`
  * `INPUT!R25` = `"64.0"`
- **CalcPr Optimization**: Thuộc tính `forceFullCalc="1"` hoàn toàn không tồn tại.
