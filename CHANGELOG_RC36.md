# OB Tool V1 RC36 Changelog

## 🚀 Tính năng đột phá: Tự động bắt số OB từ VL06O theo STO + ERNAM (Người tạo)

### 1. Bối cảnh & Vấn đề giải quyết
* Ở các bản RC trước (RC34, RC35 R1–R5), sau khi bấm tạo Delivery ở VL10D, SAP chạy background và không trả trực tiếp số OB trên màn hình.
* Bản RC35 R5 phải dừng lại yêu cầu người dùng tra Creation Log thủ công và dán số OB vào cấu hình (`VL10D_GROUP_OB_RECEIPTS`) để đối soát.
* Trên RC36, hệ thống đã giải quyết triệt để nút thắt này: Tự động liên kết và trích xuất số OB từ VL06O dựa trên sự kết hợp độc nhất:
  - **Mã STO** (`VGBEL` trùng với STO cần tạo)
  - **Mã Vật tư** (`MATNR` trùng với mã INPUT)
  - **Người tạo** (`ERNAM` lấy tự động từ `session.Info.User` của phiên SAP đang chạy)
  - **Trạng thái xuất** (`CANDIDATE`, GM = 'A', chưa gán SLoc/Batch)

### 2. Chi tiết kỹ thuật nâng cấp
* **`ModuleSAP.bas`:**
  - Bổ sung nhận diện 2 trường ALV Grid mới: `colCreator` (`ERNAM` - Người tạo) và `colDate` (`ERDAT` - Ngày tạo).
  - Tự động lấy SAP User ID qua `ses.Info.User` và lưu vào cấu hình `SAP_CURRENT_USER`.
  - Mở rộng cấu trúc bảng `VL06O_DATA` sang 16 cột (Cột 15: `CREATED_BY`, Cột 16: `CREATED_ON`).
* **`ModuleMain.bas`:**
  - Sau khi bấm `EDC4007_CreateGroup` trong VL10D, tool tự động kích hoạt `EDCRefreshGroupOB` làm mới danh sách `VL06O_DATA`.
  - Tích hợp hàm `EDCDiscoverReceiptsFromVL06O`: Tự động tìm kiếm các dòng thỏa mãn STO + Material + Creator trùng khớp và tổng hợp thành biên nhận chuẩn `VL10D_CREATION_LOG`.
  - Nếu `VL10D_GROUP_OB_RECEIPTS` trống, tool tự động điền và giải phóng khóa pending ngay lập tức, chuyển thẳng sang lập kế hoạch phân bổ.
* **`ModuleHelper.bas`:**
  - Bổ sung `Public Sub PutConfig` hỗ trợ ghi cấu hình tập trung.
* **Bảo toàn 100% Invariants:**
  - Nhánh 3001 Shop Body không bị ảnh hưởng (WholeBatchGate, giữ nguyên 4 Invariants).
  - Không gây nguy cơ double Create: Nếu không tìm thấy dòng khớp STO+Creator, tool vẫn giữ pending an toàn (`fail-closed`).
