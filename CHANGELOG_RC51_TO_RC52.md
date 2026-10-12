# CHANGELOG - OB TOOL V1 RC52

## 1. Tổng quan Bản phát hành RC52 (VBS Xóa OB 3001, Khôi phục Sau Khi OB Bị Xóa, Tách Khóa STO & Tối ưu Tốc độ VL10D)
- **Mục tiêu**: Xử lý triệt để 4 nhóm vấn đề trọng tâm theo hồ sơ kỹ thuật `RC51/PLAN_FIX_DELETE_RESUME_VL10D_SPEED.md` và bản ghi chuẩn `RC51/deleteob.vbs`:
  1. Tích hợp chuẩn xác quy trình xóa draft 3001 qua VL02N theo đúng bản ghi `deleteob.vbs`.
  2. Khôi phục lập kế hoạch và gỡ bỏ bẫy `OB_ALREADY_CREATED_WAIT` khi OB tạo thiếu đã bị người dùng xóa khỏi SAP.
  3. Tách biệt lỗi khóa STO (`WAIT_STO_LOCK`) khỏi thiếu tồn/ATP; không gọi xóa 3001 khi STO bị khóa bởi người dùng khác.
  4. Tối ưu hóa tốc độ chuẩn bị VL10D trước Background Create (gom theo trang, bỏ scroll lặp thừa, loại bỏ lần stage dư).
  5. Loại bỏ triệt để hộp thoại `InputBox`/`MsgBox` còn sót trong `EDCExistingCapacity`.
- **Bảo toàn Invariants**: Giữ nguyên vẹn 100% logic của Branch 3001 Body Shop, Branch 2001 Shop XMĐ (/SCWM/MON EWM Dispatcher) và toàn bộ các tính năng P0, P1, P2 từ RC50/RC51.

---

## 2. Chi tiết Cải tiến & Nghiệp vụ Mới

### 2.1. P0 — Triển khai Hàm Xóa Draft 3001 Theo Đúng `deleteob.vbs` (`ModuleSAP.bas`)
- Trong `SAPX_DeleteDraftOB`:
  * Mở giao dịch `VL02N`, điền số OB và gõ Enter (`sendVKey 0`).
  * Kiểm tra lỗi mở ban đầu: Phát hiện chứng từ đang bị khóa hoặc không tồn tại.
  * Thao tác trực tiếp theo đúng bản ghi `deleteob.vbs`:
    1. Nhấn nút Xóa trên Toolbar: `ses.findById("wnd[0]/tbar[1]/btn[14]").press`.
    2. Chờ popup xác nhận xóa và nhấn: `ses.findById("wnd[1]/usr/btnSPOP-OPTION1").press`.
  * Hậu kiểm live: Mở lại tra cứu số OB trong `VL02N` để xác nhận thông báo `does not exist`/không còn tồn tại trên SAP trước khi công nhận xóa thành công.
  * Phân định trạng thái trả về: `DELETE_VERIFIED`, `WAIT_OB_LOCK`, `WAIT_STO_LOCK`, `DELETE_FAILED`.
  * Trong `SAPX_FindAndReleaseDraft3001OB`: Xử lý rành mạch các trạng thái trả về, tránh báo `NOT_FOUND` giả khi thực tế là xóa thất bại hoặc bị khóa.

### 2.2. P0 — Khôi phục Chạy Tiếp Khi OB Sinh Thiếu Đã Bị Xóa (`ModuleMain.bas`)
- Trong `EDCPlanOne`:
  * Xử lý trường hợp người dùng xóa Outbound Delivery sinh thiếu (ví dụ cần 200, tạo 102 bị xóa):
    - Khi phát hiện receipt tồn tại trong `VL10D_DIRECT_DONE_<mat>` nhưng `capacity = 0`:
      + Xác nhận trạng thái `OB_DELETED_CONFIRMED` (OB không còn trong VL06O).
      + Xóa receipt cũ (`PutConfig "VL10D_DIRECT_DONE_" & mat, ""`), ghi log `OB_DELETED_REPLAN`.
      + Không ném lỗi `OB_ALREADY_CREATED_WAIT` nữa!
      + Lập lại kế hoạch cho toàn bộ số lượng thực tế cần tạo theo nguồn live (`missing = need`, ví dụ: lập lại đủ 200 pcs).

### 2.3. P0 — Tách Biệt Lỗi Khóa STO (`WAIT_STO_LOCK`)
- Nhận diện chính xác thông báo `Purchase order ... is currently being processed by ...`:
  * Gán nhãn trạng thái `WAIT_STO_LOCK`.
  * Tuyệt đối không gọi quy trình hủy draft 3001 vì đây là lỗi do người khác đang mở chứng từ, không phải do thiếu tồn ATP.
  * Ghi chú rõ ràng lên ô Note: `WAIT_STO_LOCK: Chung tu bi khoa boi nguoi khac; vui long yeu cau nha khoa`.

### 2.4. P1 — Tối Ưu Hóa Tốc Độ Chuẩn Bị VL10D Trước Background (`ModuleSAP.bas`)
- Trong `EDCStageGroup`:
  * Staging theo trang: Gom nhóm các vật tư cần thao tác theo từng trang (`VerticalScrollbar.Position`).
  * Chỉ cuộn tới mỗi trang một lần, điền số lượng và tick chọn toàn bộ các dòng trên trang đó rồi mới chuyển trang kế tiếp, loại bỏ hoàn toàn việc cuộn đi cuộn lại về dòng đầu cho từng vật tư.
  * Bổ sung log đo thời gian thực tế `VL10D_TIMING`.
- Trong `EDC4007_CreateGroup`:
  * Loại bỏ lệnh gọi `EDCStageGroup False` lặp lại dư thừa ngay trước khi bấm Create Delivery, tiết kiệm đáng kể thời gian thao tác màn hình SAP.

### 2.5. P3 — Loại Bỏ Hoàn Toàn Hộp Thoại Popup Còn Sót Trong `EDCExistingCapacity`
- Trong `EDCExistingCapacity`:
  * Loại bỏ hoàn toàn `Application.InputBox` và `MsgBox` khi phát hiện vật tư có OB thuộc nhiều STO khác nhau.
  * Tự động ưu tiên STO chủ đạo (`mEDCDominantSTO`) hoặc STO có sức chứa (capacity) lớn nhất.

---

## 3. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: PASS 100% (`control.Enabled` chuyển từ `True` sang `False`, không cảnh báo hay lỗi cú pháp).
- **Shape Macro Bindings**:
  * `INPUT!OBV1__0` $ightarrow$ `'261010_OB_TOOL_V1_RC52.xlsm'!OB_Check`
  * `INPUT!OBV1__1` $ightarrow$ `'261010_OB_TOOL_V1_RC52.xlsm'!OB_Execute`
- **Khử bỏ `forceFullCalc`**: Bảo đảm `forceFullCalc="1"` bị xóa sạch khỏi `xl/workbook.xml` (`calcPr calcId="181029"/>`).
- **Mã băm SHA-256**: Tạo bảng kê SHA-256 chính xác cho workbook và handoff zip package.
