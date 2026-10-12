# CHANGELOG - OB TOOL V1 RC50

## 1. Tổng quan Bản phát hành RC50 (Sửa Nhánh 4007, Giải phóng 3001, Xử lý triệt để UoM & Dominant STO Clustering)
- **Mục tiêu**: Xử lý triệt để 4 nhóm vấn đề cốt lõi (P0 - Xác minh Receipt & Deficit, P1 - Giải phóng Draft 3001, P2 - Đọc UoM & Hỗ trợ đơn vị tương đương/thập phân, P3 - Dominant STO Clustering & Loại bỏ Popup).
- **Bảo toàn Invariants**: Giữ nguyên vẹn 100% logic, luồng chạy và độ ổn định của Nhánh 3001 Body Shop và Nhánh 2001 Shop XMĐ (/SCWM/MON EWM Dispatcher).
- **Kiểm định Thực thi**: VBE Debug Compile PASS 100%, khử bỏ `forceFullCalc="1"`, liên kết macro nút bấm chính xác.

---

## 2. Chi tiết Cải tiến & Nghiệp vụ Mới

### 2.1. P0 — Sửa Xác minh Receipt và Dữ liệu Thiếu Trước Khi Cho Phép Cancel
- **Lưu Create Intent Bất biến (`VL10D_INTENT_<mat>`)**:
  * Lưu trữ cấu trúc bất biến gồm: Request ID, Material, Qty định tạo, STO/Item, Timestamp.
  * Luôn dọn sạch receipt cũ và pending key (`VL10D_PENDING_KEY`, `VL10D_GROUP_PENDING`, `VL10D_GROUP_OB_RECEIPTS`) trước mỗi đợt tạo nhóm mới để loại bỏ triệt để rủi ro nhận nhầm receipt cũ từ request trước.
- **Loại bỏ Hoàn toàn `On Error Resume Next` Nuốt Lỗi**:
  * Thay thế các khối bắt lỗi mù quáng bằng xử lý ngoại lệ có kiểm soát `On Error GoTo Failed`, bảo toàn thông tin lỗi và mã trạng thái chi tiết.
- **Xóa Bỏ Lỗi Thay Thế Số lượng Kế hoạch (`obQty <= 0`)**:
  * Loại bỏ hoàn toàn dòng mã lỗi `If obQty <= 0 Then obQty = CDbl(planned(3))` trong `EDCVerifyGroup`.
  * Nếu SAP không sinh hoặc đọc về 0/âm, xác định chính xác số lượng thực tế là `0#`, không tự ngộ nhận số lượng kế hoạch đã hoàn thành.
- **Lọc OB Baseline trong Tự động Khám phá VL06O (`EDCDiscoverReceiptsFromVL06O`)**:
  * Loại trừ tuyệt đối các OB đã tồn tại từ trước (`planned(8)`) thông qua kiểm tra `InStr(baselineIds, "|" & ob & "|") = 0`.
- **Phân loại Kết quả Rõ ràng Từng Mã**:
  * Định danh và ghi log chi tiết: `VERIFIED_FULL`, `VERIFIED_PARTIAL`, `VERIFIED_ZERO`, `UNKNOWN_READBACK`.
- **Commit Riêng Phần Đã Xác minh — Loại bỏ Hoàn toàn "COMMITTED Materials=0"**:
  * Hàm `EDCCommitGroup` chỉ tiếp nhận và commit `verifiedPlans` (`verifiedPlans.Count > 0`).
  * Nếu số lượng xác minh = 0, bỏ qua commit, bảo toàn trạng thái pending và ghi log `VL10D_GROUP_ZERO_VERIFIED`, tuyệt đối không báo `COMMITTED Materials=0` như một thành công giả.

### 2.2. P1 — Đưa Recovery 3001 Vào Đúng Các Đường Quyết định
- **Kích hoạt Giải phóng 3001 Trên Thâm hụt Đã Xác minh**:
  * Chỉ gọi `EDCReleaseDrafts3001` khi Due đọc hợp lệ và có thâm hụt thực tế đã được xác minh (`deficits.Count > 0`).
- **Chẩn đoán An toàn Khi Due = 0**:
  * Khi `Due = 0`, chẩn đoán STO mà không tự hủy draft 3001 bừa bãi, thông báo `WAIT_STO_DUE_4007`.
- **Thẩm định Ứng viên Draft 3001 Trước Khi Hủy**:
  * Kiểm tra đúng Ship-to `3001`, đúng Material, Batch và SLoc trống, chưa GI. Kiểm tra live ngay trước thao tác xóa trong `SAPX_FindAndReleaseDraft3001OB`.
- **Đồng bộ ATP SAP Sau Hủy (`WAIT_RELEASE_SYNC`)**:
  * Chờ 2 giây (`Application.Wait Now + TimeSerial(0, 0, 2)`) sau khi xóa draft 3001 để hệ thống SAP ATP giải phóng tồn kho trước khi kích hoạt quy trình tự động thử lại (`AUTO_RETRY_VL10D`).

### 2.3. P2 — Sửa Đọc Classic UoM và Xử lý Triệt để Đơn vị Khác EA
- **Nhật ký Đọc UoM Chi tiết Toàn diện (`VL10D_UOM_READ`)**:
  * Ghi nhận đầy đủ: `raw value`, `normalized value`, `expected UoM`, `Material`, `STO/item`, `page`, `row`, `x coordinate` và `control ID`.
- **Tách bạch Rõ ràng Trạng thái Lỗi UoM**:
  * Phân biệt rành mạch giữa `UOM_EMPTY` (ô trống trên giao diện) và `UOM_MISMATCH` (sai lệch đơn vị tính thực sự).
- **Hàm Chuẩn hóa `NormalizeUoM`**:
  * Loại bỏ hoàn toàn khoảng trắng ẩn: non-breaking space `Chr(160)`, tab `Chr(9)`, dòng mới `Chr(10)`, `Chr(13)` và chuyển thành chữ hoa.
- **Hỗ trợ Nhóm Đơn vị Đếm Tương đương (`IsCompatibleUoM`)**:
  * Công nhận tính tương thích qua lại giữa các đơn vị đếm cơ sở chuẩn SAP: `EA`, `UN`, `PC`, `ST`, `PCE`, không văng lỗi giả khi hệ thống SAP cấu hình `UN` hoặc `EA`.
- **Hỗ trợ Định dạng Số Thập phân Cho Đơn vị Khác EA (`FormatSAPQty`)**:
  * Hàm `FormatSAPQty` định dạng số lượng nguyên (`10`) và số lượng thập phân chuẩn SAP (`10.5` hoặc `10.250`), hỗ trợ phân bổ lô trong VL02N mượt mà.

### 2.4. P3 — Dominant STO Clustering & Loại bỏ Hoàn toàn Popup
- **Loại bỏ Toàn bộ Hộp thoại Popup**:
  * Loại bỏ triệt để `Application.InputBox` và `MsgBox` trong `EDCPlanOne` làm gián đoạn tự động hóa của người dùng.
- **Thuật toán Dominant STO Clustering (`EDCBuildCreatePlans`)**:
  * Tự động quét toàn bộ danh sách vật tư cần tạo trong đợt chạy và tính toán tần suất xuất hiện của từng STO trong phạm vi R31.
  * Tự động chọn STO chủ đạo (Dominant STO) gom được nhiều vật tư nhất cùng nhau để tối đa hóa việc gộp chung vào 1 Outbound Delivery trên VL10D.
  * Cơ chế Tie-breaker: Nếu hai STO gom được số lượng mã ngang nhau, ưu tiên STO đứng trước trong phạm vi R31 và có Due Qty lớn hơn.

---

## 3. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: PASS 100% (`control.Enabled` chuyển từ `True` sang `False`, không cảnh báo hay lỗi cú pháp).
- **Shape Macro Bindings**:
  * `INPUT!OBV1__0` $ightarrow$ `'261010_OB_TOOL_V1_RC50.xlsm'!OB_Check`
  * `INPUT!OBV1__1` $ightarrow$ `'261010_OB_TOOL_V1_RC50.xlsm'!OB_Execute`
- **Khử bỏ `forceFullCalc`**: Bảo đảm `forceFullCalc="1"` bị xóa sạch khỏi `xl/workbook.xml` (`calcPr calcId="181029"/>`).
- **Mã băm SHA-256**: Tạo bảng kê SHA-256 chính xác cho workbook và handoff zip package.
