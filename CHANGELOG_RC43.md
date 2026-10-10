# NHẬT KÝ THAY ĐỔI VÀ NÂNG CẤP KỸ THUẬT: OB TOOL V1 RC43

**Ngày phát hành:** 10/10/2026  
**Phiên bản:** OB TOOL V1 RC43  
**Tập tin Workbook:** `261010_OB_TOOL_V1_RC43.xlsm`  
**Gói bàn giao:** `RC43_HANDOFF.zip`  

---

## 1. NGUYÊN NHÂN GỐC RỄ & GIẢI PHÁP TỐI ƯU HÓA TRÊN RC43

Dựa trên dữ liệu LOG thực tế từ phiên chạy 11:19 PM của RC42:

### 1.1. [P0 - Nhánh 4007] Xác thực trực tiếp OB từ VL06O (Direct VL06O Receipt Verification)
* **Hiện tượng:**
  - Nhánh 4007 trên VL10D đã tạo thành công Outbound Delivery mới (OB `1012698826` với số lượng 280 pcs cho vật tư `BIW70004090AA`).
  - `EDCRefreshGroupOB` làm mới VL06O và nạp thành công OB `1012698826` vào bộ nhớ với trạng thái `CANDIDATE`.
  - Khâu nhận diện biên lai `EDCDiscoverReceiptsFromVL06O` đã khớp đúng `OB=1012698826/10; Creator=TAIPT8; STO=3601040480` và đánh dấu `VL06O_AUTO_DISCOVERED`.
  - Tuy nhiên, sau đó hàm `EDCVerifyGroup` gọi `EDC4007_ReadOBReference` mở màn hình hiển thị delivery `VL03N`. Tại giao diện SAP của người dùng, bảng chi tiết item của VL03N bị thiếu cột số lượng giao hàng `LIPS-LFIMG` ở chế độ hiển thị hiện tại, dẫn tới lỗi runtime 4205:
    `VL10D_TARGET_READ_WAIT: OB=1012698826; attempt=1/2; VL03N; Thieu field LIPS-LFIMG row 0. Can MAP layout VL02N.`
  - Lỗi này làm huỷ quy trình ghi nhận OB vừa tạo, dẫn đến tình trạng tool dừng lại và báo thiếu OB: `WAIT_FULL_BATCH_CAPACITY: ... THIEU OB (VL06O) ... PLAN: 0 item; 0 allocations`.
* **Giải pháp trên RC43:**
  - **Bỏ qua bước mở VL03N không cần thiết:** Khi biên lai được đánh dấu `VL06O_AUTO_DISCOVERED`, hàm `EDCVerifyGroup` trực tiếp xác thực số lượng và thông tin tham chiếu từ bảng dữ liệu `VL06O_DATA` trong RAM qua hàm mới `EDCGetVL06OQty(ob, item, mat)`.
  - **Cơ chế đọc an toàn (Safe Fallback):** Trong trường hợp buộc phải đọc qua VL03N, chuyển từ `TableField` nghiêm ngặt sang `EDCOptionalTableField` cho các trường `LIPS-MATNR` và `LIPS-LFIMG`, triệt tiêu hoàn toàn rủi ro văng lỗi 4205 do layout SAP cá nhân của người dùng.
  - **Bổ sung bộ nhớ đệm tham chiếu (Reference Cache):** Bổ sung hàm `EDCSetReferenceCache` và xử lý an toàn trong `EDCEnsureGroupInSnapshot` khi cột STO/Item trên VL06O_DATA tạm thời để trống.

---

### 1.2. [P0 - Hiệu năng VL10D] Tái sử dụng màn hình Due - Chỉ mở VL10D đúng 1 lần (Single VL10D Run)
* **Hiện tượng:**
  - Trong lượt chạy Nút 1 trước đây, tool mở màn hình VL10D lần thứ nhất để kiểm tra số lượng Due khả dụng. Sau khi lập xong kế hoạch tạo OB, tool lại gọi lại lệnh `SAPX_Navigate ses, "VL10D"` để mở lại VL10D lần thứ hai trước khi bấm Background Delivery.
* **Giải pháp trên RC43:**
  - Bổ sung biến lưu bộ nhớ đệm `mEDCLastDue` và dấu thời gian `mEDCLastDueAt`.
  - Trong hàm `EDC4007_PrepareGroup`, kiểm tra điều kiện `canReuseScreen`: nếu màn hình danh sách Due của VL10D đang hiển thị hợp lệ cho cùng STO và chứa đầy đủ các vật tư cần tạo trong vòng 30 giây, tool sẽ tái sử dụng ngay danh sách hiện tại (`Set due = mEDCLastDue`) mà không tải lại màn hình VL10D từ đầu.

---

### 1.3. [P0 - Trải nghiệm người dùng] Phản hồi trạng thái thành công rõ ràng trên Nút 1
* **Hiện tượng:**
  - Sau khi Nút 1 tạo thành công OB, người dùng không nhận được thông báo phản hồi rõ ràng là đã tạo xong hay chưa, dẫn đến băn khoăn về bước tiếp theo.
* **Giải pháp trên RC43:**
  - Trong `ModuleMain.OB_Check`, khi hoàn tất kiểm tra và tạo mới OB thành công cho nhánh 4007, hệ thống cập nhật thông báo rõ ràng trên ô trạng thái:
    `TAO OB THANH CONG: [số lượng alloc] ALLOC; SAN SANG BAM NUT 2; MODE=[chế độ]`
  - Giúp người dùng biết chắc chắn OB mới đã sẵn sàng và an tâm bấm Nút 2 để thực hiện phân bổ lô và kho xuất.

---

## 2. KẾT QUẢ KIỂM THỬ VÀ BIÊN DỊCH VBE

- **VBE Debug Compile:** PASS 100% (`compile_enabled_before = True` $\rightarrow$ `compile_enabled_after = False`).
- **Bảo toàn toàn vẹn:** 4 Invariants, WholeBatchGate, FIFO Allocation, Fast Save Mode.
- **Tập tin phát hành:**
  - `261010_OB_TOOL_V1_RC43.xlsm` (SHA256: `A406F9265328781F6BBDB2F63F7158666257976D9333ECFBD6CC4445F2FA4F6F`)
  - `RC43_HANDOFF.zip` kèm mã băm SHA256 và diff chi tiết.
