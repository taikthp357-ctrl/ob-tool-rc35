# NHẬT KÝ THAY ĐỔI VÀ NÂNG CẤP KỸ THUẬT: OB TOOL V1 RC48

**Ngày phát hành:** 13/10/2026  
**Phiên bản:** OB TOOL V1 RC48  
**Tập tin Workbook:** `261010_OB_TOOL_V1_RC48.xlsm`  
**Gói bàn giao:** `RC48_HANDOFF.zip`  

---

## 1. NÂNG GIỚI HẠN STO TỪ 2 LÊN TỐI ĐA 4 (PLANT 4007 STO)
* **Quy chuẩn định dạng & Giới hạn STO:**
  - Trong `ModuleMain.bas` (`EDCNormalizeSTOs`), hỗ trợ chuẩn hóa phân tách đa dạng: dấu chấm phẩy (`;`), dấu phẩy (`,`), xuống dòng (`vbCrLf`, `vbCr`, `vbLf`), và tab (`vbTab`).
  - Nâng trần kiểm tra STO từ tối đa 2 lên **tối đa 4 STO** (`uniqueSTOs.count > 4 Then Err.Raise ... "STO R31: toi da 4 STO cho mot lan kiem tra."`).
  - Cảnh báo nhập sai định dạng hoặc vượt quá 4 STO với thông điệp rõ ràng, ngăn chặn lỗi nhập liệu.
* **Giao diện bảng tính:**
  - Trên sheet `INPUT`, cập nhật nhãn ô `P31` từ `STO 4007 (tối đa 2)` thành **`STO 4007 (tối đa 4)`**.
  - Ô cấu hình `R31` tiếp nhận từ 1 đến 4 STO, tự động làm sạch và đồng bộ vào quy trình kiểm tra và lập kế hoạch.

---

## 2. GOM MÃ ĐỌC VL10D TỐI ƯU CHO NHIỀU STO (BULK MULTI-STO READ)
* **Vấn đề trên các phiên bản trước:**
  - Khi có nhiều mã và nhiều STO, hệ thống lặp mở VL10D cho từng mã × từng STO (ví dụ: 10 mã × 4 STO = 40 lần mở VL10D), gây tiêu tốn thời gian (2-5 phút) và quá tải kết nối SAP GUI.
* **Cơ chế tối ưu hóa một lượt (Single Bulk Query) trên RC48:**
  - **Lọc nhiều STO trên Tab 5:** Bổ sung hàm `EDCSetGroupSTOFilter` trong `ModuleSAP.bas`. Khi có từ 2 đến 4 STO, tự động kích hoạt Multiple Selection popup `%_ST_EBELN_%_APP_%-VALU_PUSH`, dán danh sách STO và xác nhận. Khi có 1 STO, điền trực tiếp vào `ST_EBELN-LOW`.
  - **Lọc nhiều Material trên Tab 3:** Tiếp tục kết hợp bộ lọc danh sách mã qua Multiple Selection `%_ST_MATNR_%_APP_%-VALU_PUSH`.
  - **Truy vấn VL10D gộp một lượt:** Hàm `EDC4007_FindDueMultiSTO` thực thi mở VL10D đúng **1 lần duy nhất**, đọc toàn bộ các dòng Due của tất cả mã và tất cả STO nằm trong phạm vi chỉ định.
  - **Phân loại bộ nhớ đệm RAM theo Material & STO:** Trong `ModuleMain.bas` (`EDCCachedDue`), toàn bộ dữ liệu Due được phân loại ngay vào từ điển RAM `byMat(mat)`. Các bước kiểm tra tiếp theo (`EDCPlanOne`, `EDCCheckOne`) truy xuất tức thì từ RAM (0.00s), loại bỏ 100% vòng lặp mở lại VL10D!
  - **Bảo toàn quy tắc nghiệp vụ:** Giữ nguyên vẹn 100% logic trừ lượng OB đã có, bảo vệ nghiêm ngặt quy tắc Whole Batch Fit và phân bổ FIFO. Khi tạo OB (`EDCExecuteCreatePlans`), hệ thống tự động gom các mã theo STO đã chọn và kiểm tra đúng dòng trước khi gửi lệnh Create.

---

## 3. QUẢN LÝ CHẾ ĐỘ TÍNH TOÁN EXCEL (CALCULATION MODE & PERFORMANCE GUARD)
* **Gỡ bỏ triệt để `forceFullCalc="1"` trong `xl/workbook.xml`:**
  - Loại bỏ hoàn toàn thuộc tính `forceFullCalc="1"` khỏi cấu trúc XML của workbook.
  - Đặt thuộc tính `wb.ForceFullCalculation = False` ở cấp độ ứng dụng.
  - Giải quyết triệt để tình trạng Excel bị đơ/lag hoặc bắt buộc tính toán lại toàn bộ khi người dùng mở đồng thời các tệp Excel khác.
* **Quản lý trạng thái Application thông minh trong VBA:**
  - Trong `ModuleMain.bas`, lưu trữ trạng thái gốc của phiên làm việc người dùng:
    `mSavedCalculation = Application.Calculation`, `mSavedEnableEvents = Application.EnableEvents`, `mSavedScreenUpdating = Application.ScreenUpdating`.
  - Tại `BeginRun`: Chuyển `Calculation` sang `xlCalculationManual` (-4135), tắt `EnableEvents` và `ScreenUpdating` để quá trình xử lý macro đạt tốc độ tối đa.
  - Tại `RestoreApplicationSettings` / `EndRun`: Khôi phục chính xác trạng thái ban đầu của người dùng (`Application.Calculation = mSavedCalculation`); tuyệt đối không đơn phương ép `xlCalculationAutomatic`.
* **Tối ưu hóa thao tác bảng tính:**
  - Loại bỏ lệnh gọi `rows.AutoFit` lặp từng dòng trong `UpdateItemStatus`, triệt tiêu độ trễ render lại giao diện khi cập nhật trạng thái đơn hàng.

---

## 4. KẾT QUẢ BIÊN DỊCH & BẢO TOÀN HỆ THỐNG
* **VBE Debug Compile:**
  - `compile_enabled_before = True`
  - `compile_enabled_after = False` (PASS 100% trên Excel Native COM).
* **Gắn Macro nút bấm:**
  - Nút 1 (Kiểm tra): `'261010_OB_TOOL_V1_RC48.xlsm'!OB_Check`
  - Nút 2 (Thực hiện): `'261010_OB_TOOL_V1_RC48.xlsm'!OB_Execute`
* **Dọn dẹp Bảng:**
  - Sheet `LOG`: Giữ đúng 5 dòng tiêu đề (max row: 5).
  - Sheet `LEDGER`: Giữ đúng 1 dòng tiêu đề (max row: 1).
  - Sheet `CONFIG`: Phiên bản `OB_V1_RC48`.
* **Bảo toàn Invariants:**
  - 4 Invariants Shop Body nhánh 3001, WholeBatchGate, và nguyên tắc FIFO được bảo toàn nguyên vẹn 100%.
