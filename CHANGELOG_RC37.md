# NHẬT KÝ THAY ĐỔI VÀ NÂNG CẤP KỸ THUẬT: OB TOOL V1 RC37

**Ngày phát hành:** 10/10/2026  
**Phiên bản:** OB TOOL V1 RC37  
**Tập tin Workbook:** `261010_OB_TOOL_V1_RC37.xlsm`  
**Gói bàn giao:** `RC37_HANDOFF.zip`  

---

## 1. TỔNG QUAN CÁC NÂNG CẤP CỐT LÕI TRÊN RC37

### 1.1. Vá lỗi bẫy rỗng VL06O cho mọi Plant (Cả 3001 và 4007)
* **Hiện trạng cũ (RC36):** Bẫy "no deliveries" chỉ áp dụng khi `ship = "4007"`. Khi Plant 3001 hoặc các Plant khác không có delivery, tool cố gắng tìm ALV grid dẫn đến treo 20 giây và văng Runtime Error 4202 (`VL06O Execute timeout`).
* **Giải pháp RC37:** Xóa bỏ điều kiện giới hạn `ship = "4007"`. Áp dụng bẫy trạng thái rỗng ("no deliveries", "no data", "keine lieferungen") chung cho tất cả các Plant. Xóa sạch vùng dữ liệu 16 cột (`A6:P`) và thoát êm ngay lập tức, triệt tiêu hoàn toàn lỗi runtime 4202.

### 1.2. Dọn dẹp siêu tốc lịch sử LOG và LEDGER (Tối ưu hiệu năng)
* Dọn dẹp sạch 4,298 dòng log cũ trên sheet `LOG` (bảo toàn nguyên vẹn 5 dòng tiêu đề & header).
* Dọn dẹp sạch 250 dòng lịch sử cũ trên sheet `LEDGER` (bảo toàn nguyên vẹn dòng 1 header).
* Giảm kích thước workbook từ **1,070,522 bytes xuống 784,358 bytes** (giảm hơn 26.7%), giúp Excel mở, tính toán và lưu siêu tốc.

### 1.3. Cơ chế tự động bắt lại OB từ VL06O khi OB cũ bị xóa hoặc không tồn tại
* Khi thực hiện OB (VL02N) gặp lỗi không tồn tại hoặc bị xóa trên SAP (`does not exist`, `deleted`):
  1. Đánh dấu loại trừ OB lỗi: `mExcludedOB(ob) = True`.
  2. Kích hoạt làm mới VL06O (`OB_ReadVL06O`) để lấy dữ liệu thực tế mới nhất từ SAP.
  3. Bảo toàn đầy đủ 16 cột (`A:P`) trong các hàm xử lý dữ liệu: `MergeUntouched`, `EDCRefreshOB`, `EDCRefreshGroupOB`, `PrepareRecoveryBatch`, `ConfirmMissingOB`.
  4. Nếu SAP đã sinh ra OB mới từ cùng STO hoặc có OB ứng viên hợp lệ, hệ thống tự động gán và phân bổ sang OB mới.
  5. Nếu không còn OB nào trong VL06O: ghi rõ cảnh báo trên sheet `INPUT` tại Cột D: `"THIEU OB (VL06O) - Nhan nut 1 de tao lai qua VL10D hoac mo VL10D kiem tra"` và popup cảnh báo nhắc người dùng mở VL10D / bấm nút 1 để kiểm tra tạo mới.

### 1.4. Cơ chế Retry 1 lần tinh gọn cho OB bị Lock / Xung đột
* Giữ nguyên cơ chế retry 1 lượt: OB bị lock (`locked`, `being processed`, `WAIT_REMOTE_SYNC`) được đưa về cuối hàng đợi `deferred` và thử lại đúng 1 lần duy nhất, tránh vòng lặp vô tận.

### 1.5. Cảnh báo tồn kho khóa / kiểm định (QI / QU / QT)
* Trong `PrepareStock` và `AllocateRequest`: Tự động nhận diện các dòng tồn thuộc trạng thái `QT`, `QI`, `QU`.
* Nếu số lượng khả dụng nhỏ hơn nhu cầu cần xuất (`available < qty`): tự động ghi chú rõ ràng vào cột D sheet `INPUT`:  
  `; (CO TON BI KHOA/QI/QU - XEM POSTING J)`.

### 1.6. Tập trung giao diện chính (Sheet INPUT)
* Mọi cảnh báo, điều phối tập trung trực tiếp trên sheet `INPUT`, không làm xáo trộn sheet `BATCH_PLAN`.
* Bảo toàn 100% 4 Invariants nhánh 3001 Shop Body, WholeBatchGate và nguyên tắc FIFO.

---

## 2. KẾT QUẢ KIỂM THỬ NATIVE & BIÊN DỊCH VBE
* **VBE Debug Compile:** `PASS 100%` (`control.Enabled = False`).
* **Button Bindings:** Đã re-bind chuẩn xác:
  * Nút 1: `'261010_OB_TOOL_V1_RC37.xlsm'!OB_Check`
  * Nút 2: `'261010_OB_TOOL_V1_RC37.xlsm'!OB_Execute`
* **VBA Modules:**
  * `ModuleLog.bas`: 23 dòng
  * `ModuleHelper.bas`: 497 dòng
  * `ModuleSAP.bas`: 5,160 dòng
  * `ModuleMain.bas`: 4,067 dòng
