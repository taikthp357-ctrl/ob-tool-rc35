# NHẬT KÝ THAY ĐỔI VÀ NÂNG CẤP KỸ THUẬT: OB TOOL V1 RC40

**Ngày phát hành:** 10/10/2026  
**Phiên bản:** OB TOOL V1 RC40  
**Tập tin Workbook:** `261010_OB_TOOL_V1_RC40.xlsm`  
**Gói bàn giao:** `RC40_HANDOFF.zip`  

---

## 1. NGUYÊN NHÂN GỐC RỄ & GIẢI PHÁP TỐI ƯU HÓA TRÊN RC40

Dựa trên dữ liệu LOG chạy thực tế tại hai nhánh **3001** và **4007**, bản RC40 giải quyết triệt để 2 vấn đề lớn:

---

### 1.1. [P0 - Nhánh 3001] Cơ chế chờ thông minh giải phóng khóa SAP (`DO_AUTO`) cho hàng đợi Deferred
* **Nguyên nhân gốc rễ:**
  - Ở lượt chạy 1 trên nhánh 3001, các đơn hàng gặp lỗi:
    `Purchase order ... is currently being processed by DO_AUTO`
    đã được đưa vào hàng đợi `deferred.keys` (`DEFER_OB`).
  - Tuy nhiên, vòng lặp thử lại `For Each k In deferred.keys` đã kích hoạt **ngay lập tức (sau 0.38s - 0.43s, không có thời gian nghỉ)**.
  - Tiến trình ngầm `DO_AUTO` của SAP (interface/RFC xử lý PO) thường giữ khóa từ 5 đến 15 giây. Khi tool thử lại sau 0.4 giây, SAP vẫn đang giữ khóa, dẫn tới lượt thử lại duy nhất bị dập tắt chớp nhoáng (1.2 giây cho cả 3 OB). Tool dừng lại và người dùng phải bấm chạy lại lần 2 thủ công (khi SAP đã xong).
* **Giải pháp trên RC40:**
  - Bổ sung hàm API Windows chuẩn `SleepMS` trong `ModuleHelper.bas`.
  - Trong `OB_Execute` (`ModuleMain.bas`), trước khi duyệt hàng đợi `deferred.keys`:
    + Kiểm tra nếu có OB bị hoãn (`deferred.count > 0`), kích hoạt cơ chế **Smart Lock Backoff (2 nhịp retry)**.
    + Nhịp 1: Tạm dừng 6 giây và hiển thị đồng hồ đếm ngược trên `Application.StatusBar`:  
      `"SAP DANG XU LY NGAM (DO_AUTO). DANG CHO 6s DE HE THONG NHA KHOA (NHIP 1/2)..."`
    + Ghi log minh bạch `WAIT_LOCK_BACKOFF` và gọi `DoEvents` để giao diện Excel không bị treo.
    + Thử lại các OB. Nếu lưu thành công $\rightarrow$ Ghi log `DONE_DEFERRED; pass=1` và tăng biến `done`.
    + Nếu sau nhịp 1 vẫn còn OB bị kẹt khóa SAP $\rightarrow$ Tiếp tục nhịp 2 với thời gian chờ 5 giây.
    + Chỉ khi sau 2 nhịp (tổng thời gian chờ ~11 giây) mà SAP vẫn bị khóa thì mới ghi nhận dừng an toàn.

---

### 1.2. [P0 - Nhánh 4007] Mở khóa màn hình chọn VL10D (Khôi phục chuẩn RC34)
* **Nguyên nhân gốc rễ:**
  - Trên nhánh 4007, khi thiếu OB và mở VL10D để kiểm tra/tạo mới, tool điền đủ thông tin (STO trên Tab 5, Material và Plant 4007 trên Tab 3) nhưng **không bấm F8 để vào màn hình Due của VL10D**, mà dừng lại với lỗi:
    `VL10D_PLAN_WAIT: ... -2147217243; VL10D supplying Plant field missing; no Create.`
  - Trong giao dịch VL10D, trường Kho xuất / Supplying Plant `ST_RESWK` không nằm trên Tab 3. Một câu lệnh chặn cứng:
    `If ctl Is Nothing Then Err.Raise vbObjectError + 4261, , "VL10D supplying Plant field missing; no Create."`
    vô tình bị thêm vào từ các bản sau RC34. Do trường này không có trên Tab 3, lệnh `Err.Raise` bị văng ra **trước khi kịp gửi phím F8 (`sendVKey 8`)**.
* **Giải pháp trên RC40:**
  - **Khôi phục hoàn toàn chuẩn RC34**: Bỏ câu lệnh chặn cứng `Err.Raise` khi `ctl Is Nothing`.
    Trường `ST_RESWK-LOW` trở lại là tùy chọn: nếu tìm thấy thì điền 3093, không thấy thì bỏ qua và tiếp tục bấm F8 (`sendVKey 8`) để vào danh sách Due của VL10D.
  - Tại lưới kết quả ALV của VL10D, chuẩn hóa lại điều kiện kiểm tra cột `RESWK`: nếu cột có hiển thị giá trị thì xác thực đúng 3093, nếu layout không hiển thị cột này thì không bị văng lỗi.
  - Việc kiểm tra Supplying Plant = 3093 vẫn được bảo vệ độc lập 100% qua ME2N (`EDC4007_CheckSTO`).

---

### 1.3. [P1] Tự động dọn dẹp cờ Pending cũ khi đổi Material hoặc bắt đầu Request mới
* **Nguyên nhân:**
  - Trong cấu hình nội bộ (`CONFIG`), `VL10D_PENDING_KEY` vẫn giữ thông tin từ phiên chạy trước của mã khác, dẫn tới cảnh báo:
    `VL10D_PENDING_GROUP_WAIT: Pending group Material outside INPUT.`
* **Giải pháp trên RC40:**
  - Trong `AssertRunIdentity`: Khi `NEW_REQUEST` kích hoạt (INPUT thay đổi hoặc Request ID mới), tự động xóa sạch `VL10D_PENDING_KEY` và `VL10D_GROUP_PENDING`.
  - Trong `EDCResolvePendingGroup`: Khi phát hiện mã pending cũ không nằm trong danh sách INPUT hiện tại, tự động dọn sạch cờ pending cũ và lưu lại cấu hình.

---

## 2. KẾT QUẢ KIỂM THỬ VÀ BIÊN DỊCH VBE

- **VBE Debug Compile:** PASS 100% (Control `Compi&le VBAProject` chuyển từ `Enabled = True` sang `Enabled = False`, không có lỗi cú pháp).
- **Sheet LOG & LEDGER:** Đã dọn sạch lịch sử cũ, workbook nhẹ mở/lưu siêu tốc.
- **Button Bindings:** Đã liên kết chuẩn xác với file `261010_OB_TOOL_V1_RC40.xlsm`.
- **Bảo toàn 4 Invariants:**
  1. Bảo toàn item ngoài kế hoạch (Non-target item preservation) $\rightarrow$ Giữ nguyên.
  2. WholeBatchGate (Không tách dở dang 1 batch) $\rightarrow$ Giữ nguyên.
  3. FIFO Allocation $\rightarrow$ Giữ nguyên.
  4. Fast Save Mode & Direct Item Locate $\rightarrow$ Hoạt động ổn định.
