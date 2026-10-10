# CHANGELOG - OB TOOL V1 RC45

## 1. Bối cảnh & Nguyên nhân từ Thực tế Vận hành (Root Cause)
- Trong phiên chạy thực tế RC44: Người dùng nhập 10 mã vật tư, công cụ tính toán lập kế hoạch nhóm VL10D cho 8 mã vật tư thỏa mãn WholeBatchGate.
- Khi gửi lệnh tạo DO trên SAP VL10D, SAP tạo thành công Delivery `1012699408` chứa 7 mã hàng, nhưng có 1 mã bị bỏ qua (0 pcs) hoặc có mã bị giảm số lượng (ví dụ cần 361 pcs nhưng chỉ tạo được 200 pcs).
- **Nguyên nhân nghiệp vụ**: Các đơn Outbound Delivery (DO) nháp của Plant 3001 tại kho xuất 3093 chưa điền SLoc và chưa chia Batch (draft reservations) đang chiếm giữ hạn mức ATP tồn kho. Khi 4007 gọi VL10D tạo DO, SAP kiểm tra ATP và tự động cắt giảm hoặc không tạo dòng hàng cho mã bị 3001 chiếm giữ.
- **Hạn chế của RC44**:
  - `EDCDiscoverReceiptsFromVL06O` yêu cầu `matchedCount = neededCount` (phải đủ 8/8 mã). Khi SAP chỉ tạo 7 mã, hàm trả về rỗng và văng lỗi fatal `GROUP_RECEIPT_MISSING: inspect VL10D creation log read-only; no Create retry.`
  - `EDCVerifyGroup` so sánh ngặt nghèo `totals(mat) = planned(3)`. Nếu có mã bị tạo thiếu số lượng, hàm văng lỗi fatal `GROUP_OB_QTY_MISMATCH`.
  - Hậu quả: Toàn bộ 7 mã hàng đã tạo thành công trên SAP DO `1012699408` bị kẹt ở trạng thái pending, không được commit, không bật Nút 2 (`OB_Execute`) để chia lô.

---

## 2. Các Nâng cấp Đột phá trong RC45 (V1 RC45)

### 2.1. Chấp nhận Từng phần Khi Tạo DO (Partial Group Acceptance & Reconciliation)
- Trong `ModuleMain.bas` (`EDCDiscoverReceiptsFromVL06O`):
  - Chuyển điều kiện kiểm tra từ `matchedCount = neededCount` sang `matchedCount > 0`.
  - Cho phép công cụ nhận diện và khớp toàn bộ các dòng hàng thực tế mà SAP đã tạo ra trên DO mới.
- Trong `ModuleMain.bas` (`EDCVerifyGroup`):
  - Xóa bỏ bẫy lỗi ngặt nghèo `GROUP_OB_QTY_MISMATCH`.
  - Phân loại xử lý tự động cho từng mã vật tư:
    1. **Mã tạo đủ 100%**: Xác nhận và ghi log `VL10D_ITEM_FULL_VERIFIED`.
    2. **Mã tạo một phần (Partial)**: Cập nhật số lượng kế hoạch theo số lượng thực tế SAP đã tạo (`planned(3) = actualQty`), ghi nhận lượng thiếu vào từ điển `deficits`, ghi chú cột D: `WAIT_STO_DUE_4007: Thieu [X] pcs do 3001 giu ton (da tao [Y] pcs)`.
    3. **Mã bị SAP bỏ qua (0 pcs)**: Ghi nhận lượng thiếu toàn phần vào `deficits`, tách mã khỏi danh sách DO hiện tại (`plans.Remove mat`), ghi chú cột D: `WAIT_STO_DUE_4007: SAP khong tao duoc DO (thieu [X] pcs do 3001 giu ton)`.
- Các mã đã tạo thành công được chuyển qua `EDCEnsureGroupInSnapshot` và `EDCCommitGroup` để giải phóng pending, lưu trữ biên nhận, và kích hoạt Nút 2 (`OB_Execute`) sẵn sàng chia lô ngay lập tức!

### 2.2. Cơ chế Tự động Quét và Xóa OB Nháp 3001 bằng `Shift + F12` (Targeted Draft 3001 Release)
- **Quy trình chuẩn hóa sau khi tạo DO**: Sau khi DO 4007 đã tạo xong và xác định chính xác số lượng thiếu (`deficitQty`):
  1. Nếu `AUTO_DELETE_DRAFT_3001 = TRUE` (mặc định BẬT trong CONFIG): Gọi `EDCReleaseDrafts3001`.
  2. Mở `VL06O` cho Ship-to = `3001`, Shipping Point = `3093`, khoảng ngày 60 ngày, GM status = `A`.
  3. Lọc danh sách dòng hàng của mã thiếu tồn.
  4. **Tiêu chí an toàn tuyệt đối cho OB nháp 3001**:
     - `SLoc` PHẢI RỖNG (`LGORT = ""`).
     - `Batch` PHẢI RỖNG (`CHARG = ""`).
     - `GM status` = `"A"` (chưa PGI).
     - `Parent Item` = `""` hoặc `0` (không phải dòng chia lô con).
     - Tuyệt đối KHÔNG chạm vào bất kỳ OB nào đã có SLoc hoặc đã có Batch!
  5. **Thuật toán chọn OB tối ưu**:
     - Tính khoảng cách lệch số lượng: `diff = Abs(OB_Qty - deficitQty)`.
     - Tìm OB có `diff` nhỏ nhất ($\min |OB\_Qty - Deficit|$), ưu tiên OB có số lượng gần sát nhất với lượng thiếu để giải phóng đúng lượng tồn ATP cần thiết.
  6. **Thực thi xóa an toàn trong VL02N**:
     - Điều hướng tới `VL02N`, mở mã OB nháp 3001 đã chọn.
     - Gửi lệnh `Shift + F12` (`sendVKey 24` trong SAP GUI Scripting).
     - Tự động bắt popup xác nhận `wnd[1]`, nhấn nút Yes / Option 1 (`btnSPOP-OPTION1`).
     - Kiểm tra thanh trạng thái `wbar`, xác nhận thông báo xóa thành công.
  7. **Cập nhật giao diện**:
     - Ghi chú chi tiết vào cột D sheet `INPUT`:
       `DA XOA OB 3001: [OB] ([Qty] pcs) DE GIAI PHONG TON; NHAN NUT 1 DE TAO LAI VL10D`
     - Người dùng chỉ cần nhấn lại Nút 1 ở chu kỳ tiếp theo, VL10D sẽ lập tức tạo tiếp lượng DO còn lại do tồn kho đã được giải phóng!

---

## 3. Bảo toàn Tuyệt đối Các Bất biến Hệ thống (System Invariants)
- **Invariant 1 (FIFO)**: Tiếp tục duy trì 100% nguyên tắc phân bổ lô cũ nhất trước.
- **Invariant 2 (Safety Lock)**: Bảo vệ dữ liệu, không ghi đè, phát hiện conflict.
- **Invariant 3 (WholeBatchGate)**: Đảm bảo nguyên lô khớp với STO Due.
- **Invariant 4 (Fast Save & UI)**: Dọn dẹp sạch sẽ LOG/LEDGER, giữ file siêu nhẹ (~780 KB), mở và lưu tức thì.

---

## 4. Bằng chứng Kiểm chứng Kỹ thuật (Verification Evidence)
- **VBE Debug Compile**:
  - `compile_enabled_before`: `True`
  - `compile_enabled_after`: `False` (Biên dịch native 100% sạch lỗi, 0 syntax error, 0 type mismatch).
- **File đầu ra chuẩn hóa**: `261010_OB_TOOL_V1_RC45.xlsm`
- **Button Macro Bindings**:
  - `INPUT!OBV1__0` -> `'261010_OB_TOOL_V1_RC45.xlsm'!OB_Check`
  - `INPUT!OBV1__1` -> `'261010_OB_TOOL_V1_RC45.xlsm'!OB_Execute`
