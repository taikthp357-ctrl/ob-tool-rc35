# CHANGELOG - OB TOOL V1 RC63 (ALL IN ONE)

## 1. Tổng quan Bản phát hành RC63 — ALL IN ONE
Bản phát hành RC63 tập trung giải quyết triệt để 2 lỗi cốt lõi phát hiện từ các đợt chạy thực tế cuối của RC62, bảo toàn toàn vẹn số lượng nguyên Batch, gia cố bảo vệ chứng từ đa mã và cung cấp tính năng **ALL IN ONE** (Chạy Tất Cả) với 1 nút bấm duy nhất.

---

## 2. Chi tiết Kỹ thuật Các Thay Đổi Cốt Lõi

### 2.1. Khắc phục Triệt để Root Cause A: Vòng lặp Xóa Giải Phóng Tồn 3001 Cam Kết 2 Pha (`SAPX_FindAndReleaseDraft3001OB_Batch`)
- **Vấn đề trong RC62**: Số lượng thiếu (`remDeficits`) bị trừ ngay khi chọn OB ứng viên trước khi gửi lệnh xóa sang VL02N. Khi OB bị khóa bởi người dùng khác hoặc chứng từ STO bị khóa (`WAIT_STO_LOCK`, ví dụ OB `1012750166` giữ 83 PCS bị khóa bởi STO `3601042690`), tool vẫn tưởng đã giải phóng xong 100 PCS và dừng tìm kiếm, dẫn đến VL10D chỉ tạo được 17 PCS (thiếu 83 PCS).
- **Giải pháp RC63**:
  * Chuyển đổi sang cơ chế **Confirmed 2-Phase Release Commit**: Chỉ trừ `remDeficits` và tích lũy `confirmedRelease` khi SAP trả về kết quả `DELETE_VERIFIED`.
  * Nếu OB bị khóa (`WAIT_STO_LOCK`, `WAIT_OB_LOCK`) hoặc xóa thất bại: Tuyệt đối **không trừ số lượng thiếu**, đánh dấu đã thử OB này và tiếp tục tìm tổ hợp ứng viên tiếp theo trong danh sách cho đến khi giải phóng đủ 100% nhu cầu hoặc hết ứng viên.
  * Tích hợp tự động trừ dồn ảnh hưởng đa mã ngay khi từng OB được xác nhận xóa thành công.

### 2.2. Khắc phục Triệt để Root Cause B: Đối soát Trạng thái Delivery VL03N & Phục hồi Thông minh Khi Lỗi WMS (`VL589`)
- **Vấn đề trong RC62**: Khi lưu chứng từ đa mã có dòng item chưa được phân bổ SLoc/Batch, SAP báo lỗi WMS `VL589: Outbound delivery could not be distributed` và hủy/rollback chứng từ. Hàm `ApplyOneOB` bắt trạng thái `UNKNOWN` và kích hoạt `Err.Raise vbObjectError + 4121`, làm sập toàn bộ tiến trình và hủy hoại cả những OB đã lưu thành công trước đó (như `1012754414`).
- **Giải pháp RC63**:
  * Bổ sung hàm công khai `SAPX_CheckDeliveryStatus(ob)` tra cứu trực tiếp qua giao dịch `VL03N`.
  * Nhận diện chính xác các trạng thái: `NOT_FOUND` (chứng từ không tồn tại trong DB/archive do rollback), `LOCKED`, `EXISTS`, hoặc `ERROR`.
  * Trong `ApplyOneOB`: Khi gặp `UNKNOWN`, `SAVED_UNVERIFIED` hoặc lỗi phân phối `VL589`, tool tự động tra cứu `VL03N`:
    - Nếu `NOT_FOUND`: Xác nhận `OB_ABSENT_CONFIRMED`, loại trừ OB lỗi khỏi kế hoạch (`mExcludedOB(ob) = True`), ghi nhận `UNSAVED` cho các item đó, đưa vào hàng đợi `QueueRecovery` để lập kế hoạch bù mà **tuyệt đối không tung `Err.Raise` làm sập tool**.
    - Bảo toàn 100% kết quả của các OB đã lưu thành công trước đó.

### 2.3. Bảo toàn Tính Toàn vẹn Nguyên Batch (Whole-Batch Integrity)
- Không bao giờ chia nhỏ Batch (ví dụ: Batch 50 PCS tuyệt đối không được gán vào item chỉ có 17 PCS).
- Phân định rõ ràng giữa `createdObQty`, `usableWholeBatchQty` và `savedWholeBatchQty`.
- Nếu số lượng tạo trên VL10D không đủ dung lượng nguyên Batch (như 17 PCS < 50 PCS), tool ghi nhận trạng thái `WAIT_FULL_BATCH_CAPACITY` và đưa vào kế hoạch tạo bù, không gán lửng lơ gây lỗi `VL589`.

### 2.4. Gia cố Xóa Cấp Dòng vs Xóa Toàn Bộ OB (`SAPX_DeleteDraftOB`)
- **Nhánh giải phóng 3001**: Giữ nguyên cơ chế xóa nhanh toàn bộ OB qua `btn[14]` (`isDraft3001Release = True`), không mở từng item.
- **Nhánh dọn dẹp cấp dòng 4007** (`targetMat <> ""` hoặc `targetItem <> ""`):
  * Kiểm tra status bar của SAP sau khi chọn dòng: Nếu SAP báo `"cannot be selected"` hoặc `"Item cannot be deleted"` -> Trả về `DELETE_ITEM_UNSUPPORTED` và dừng lại an toàn.
  * Nếu không tìm thấy dòng cần xóa (`targetRow < 0`): Trả về `DELETE_ITEM_NOT_FOUND`.
  * **Tuyệt đối không bao giờ rơi (fall-through) sang `DeleteWholeOB` khi đang yêu cầu xóa dòng lẻ**, bảo vệ an toàn cho các dòng item hợp lệ còn lại trên chứng từ đa mã.

### 2.5. Điều phối Tổng thể ALL IN ONE (`OB_RunAll` & `ExecuteAllocationsCore`)
- Bổ sung nút bấm mới: **`⚡ CHẠY TẤT CẢ (ALL IN ONE)`** trên sheet `INPUT`.
- Macro `OB_RunAll` tự động:
  1. Kiểm tra cấu hình và tự động kích hoạt `MODE = "APPLY"`.
  2. Điều hướng tự động theo Ship-to (`INPUT!R5`):
     * Ship-to `4007`: Tự động chạy toàn bộ pipeline EDC (Kiểm tra -> Quét giải phóng 3001 nếu thiếu Due -> Tạo OB qua VL10D -> Xác minh OB -> Lập kế hoạch VL02N -> Thực thi ghi SLoc/Batch qua VL02N -> Tự động phục hồi nếu có lỗi).
     * Ship-to `2001`: Tự động điều phối qua nhánh EWM Shop XMD (`OB2001_Check` -> `OB2001_Execute`).
     * Ship-to tiêu chuẩn khác: Chạy Check -> Execute liên hoàn.
  3. Tách biệt `ExecuteAllocationsCore(isAllInOne)` cho phép cả 3 nút (Nút 1 Check, Nút 2 Execute, Nút 3 Run All) hoạt động độc lập và liền mạch.

---

## 3. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: PASS 100% (`before=True`, `after=False`).
- **Shape Macro Bindings**:
  * `INPUT!OBV1__ALL` -> `'261010_OB_TOOL_V1_RC63.xlsm'!OB_RunAll`
  * `INPUT!OBV1__0` -> `'261010_OB_TOOL_V1_RC63.xlsm'!OB_Check`
  * `INPUT!OBV1__1` -> `'261010_OB_TOOL_V1_RC63.xlsm'!OB_Execute`
- **Version Labels**:
  * `INPUT!A1` = `"OB TOOL V1 RC63"`
  * `CONFIG!B2` = `"RC63"`
  * `INPUT!R22` = `"RC63 ALL-IN-ONE & CONFIRMED 3001 RELEASE PIPELINE"`
  * `INPUT!R25` = `"63.0"`
- **CalcPr Optimization**: Thuộc tính `forceFullCalc="1"` hoàn toàn không tồn tại.
