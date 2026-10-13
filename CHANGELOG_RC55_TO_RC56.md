# CHANGELOG - OB TOOL V1 RC56

## 1. Tổng quan Bản phát hành RC56 (Vá Schema Pending Key 6 Trường, Nhánh Tạo 1 Nửa STO / Tự Động Xóa OB Thiếu, Gỡ Presave Extra Item & Đối Chiếu An Toàn Draft 3001)
- **Mục tiêu**: Xử lý triệt để 6 nhóm vấn đề kỹ thuật và vận hành theo hồ sơ chỉ đạo `BRIEF_RC56_PARTIAL_STO_PENDING_SCHEMA.md` trên baseline `261010_OB_TOOL_V1_RC55.xlsm`:
  1. Chuẩn hóa schema `VL10D_PENDING_KEY` 6 trường (`Request | Material | Qty | STO | Item | Timestamp`), khắc phục lỗi deserializer `Invalid pending item identity` và không xóa nhầm pending key khi chưa xác minh.
  2. Bảo toàn `VL10D_PENDING_KEY` trong quy trình `AUTO_RETRY_VL10D`, ngăn chặn việc xóa pending sớm trong `EDCCommitGroup` trước khi lượt tạo bổ sung hoàn tất.
  3. Bổ sung nhánh xử lý khi VL10D chỉ tạo được một phần STO ("tạo 1 nửa STO"): Nếu nhánh 3001 không giữ tồn hoặc không thể sinh thêm để đạt nguyên batch, tự động xóa OB tạo thiếu ("xóa tạo bừa đó") qua `SAPX_DeleteDraftOB` per `deleteob.vbs` để hoàn trả Due về STO.
  4. Gỡ bỏ bẫy `PRESAVE_GATE_FAIL: EXTRA_ITEM` trong VL02N và `SavedSnapshotMatchesPlan`: Cho phép lưu OB khi trên chứng từ có các item khác chưa có kế hoạch trong lượt chạy hiện tại.
  5. Sửa lỗi `TableField` đọc tràn dòng gây lỗi runtime 4205 (`Thieu field LIPS-POSNR row 3`).
  6. Đối chiếu an toàn khi tìm xóa Draft 3001 (`VGBEL` & `ERNAM`), mở rộng phạm vi ngày Planned GI tương lai trong VL06O (+60 ngày), và ưu tiên STO nhận đủ nguyên batch.
- **Bảo toàn Invariants**: Giữ nguyên vẹn 100% logic của Branch 3001 Body Shop, Branch 2001 Shop XMĐ (/SCWM/MON EWM Dispatcher), tối ưu tốc độ VL10D staging và toàn bộ các tính năng từ RC50-RC55.

---

## 2. Chi tiết Cải tiến & Nghiệp vụ Mới

### 2.1. P0 — Vá lỗi Schema Deserializer `VL10D_PENDING_KEY` 6 trường (`ModuleMain.bas`)
- Trong `EDCResolvePendingCreate`:
  * Nâng cấp bộ giải mã chấp nhận cả định dạng 5 trường cũ và 6 trường mới (`If UBound(parts) < 4 Then Err.Raise ...`).
- Trong `EDCResolvePendingGroup`:
  * Chấp nhận deserializer 6 trường: `parts(0)`: RequestId, `parts(1)`: Material, `parts(2)`: Qty, `parts(3)`: STO, `parts(4)`: Item, `parts(5)`: Timestamp.
  * Cơ chế bảo vệ Pending Key: Tuyệt đối không xóa pending key trong khối `Waiting:` do lỗi tạm thời; chỉ xóa khi đã xác minh thành công trên SAP hoặc vật tư nằm ngoài INPUT.

### 2.2. P0 — Bảo toàn Pending Key trong quy trình `AUTO_RETRY_VL10D` (`ModuleMain.bas`)
- Trong `EDCCommitGroup`:
  * Bổ sung tham số `Optional ByVal keepPendingKey As Boolean = False`.
  * Khi có đợt retry đang chuẩn bị thực hiện (`releasedDraft And deficits.count > 0`), giữ nguyên `VL10D_PENDING_KEY` và `VL10D_GROUP_PENDING`.
  * Cơ chế fallback linh hoạt nếu pending identity bị gián đoạn: Tự động khôi phục từ danh sách `plans` đang xử lý.
- Trong `EDCCreateGroup`:
  * Tái tạo và lưu trữ đầy đủ `VL10D_GROUP_PENDING` và `VL10D_PENDING_KEY` trước khi gọi SAP Create trong lượt `AUTO_RETRY_VL10D`.
  * Trong `EDCVerifyGroup`: Bổ sung fallback xác định STO/Request từ `plans` để chống lỗi `GROUP_RECEIPT_MISSING`.

### 2.3. P0 — Bổ sung Nhánh Xử Lý Tạo 1 Nửa STO & Tự Động Xóa OB Thiếu (`ModuleMain.bas`)
- Khi VL10D tạo xong nhóm hàng và có mã chỉ tạo được một phần số lượng (`actualQty < plannedQty`):
  * Đã thử giải phóng qua 3001 và retry nhưng vẫn không thể tạo đủ.
  * Kiểm tra cấu hình `AUTO_DELETE_PARTIAL_STO_OB` (mặc định = True / "1"):
    - Kiểm tra xem OB mới tạo có chứa mã vật tư nào khác đã hoàn thành đủ số lượng hay không (`obHasCompletedMat`).
    - Nếu OB chứa mã khác đã hoàn thành: Giữ OB, ghi chú `"VERIFIED_PARTIAL (DA TAO " & actualQty & "/" & plannedQty & " PCS)"`.
    - Nếu OB chỉ chứa mã bị thiếu: Tự động gọi `ModuleSAP.SAPX_DeleteDraftOB(ses, partialOB)` để xóa chứng từ dở dang, trả Due về lại STO.
    - Ghi log `VL10D_PARTIAL_OB_CLEANUP` và cập nhật ghi chú trên sheet INPUT: `"DA XOA OB THIEU (TAO " & actualQty & "/" & plannedQty & " PCS): HET TON & 3001 KO GIU TON; TRA LAI STO DUE"`.

### 2.4. P1 — Gỡ bỏ bẫy `PRESAVE_GATE_FAIL: EXTRA_ITEM` trong VL02N (`ModuleSAP.bas`)
- Trong `VerifyDeliveryPlan`:
  * Không ném ngoại lệ dừng lưu khi phát hiện các item ngoài kế hoạch hiện tại (ví dụ: chứng từ có item 20 nhưng lượt chạy chỉ làm item 10, 30, 40).
  * Chuyển sang ghi log thông tin `PRESAVE_EXTRA_ITEM_INFO` và tập trung kiểm tra tính hợp lệ 100% của các item có trong `itemMap`.
- Trong `SavedSnapshotMatchesPlan`:
  * Bỏ kiểm tra cấm `EXTRA_ITEM` đối với các item không nằm trong kế hoạch; chỉ xác nhận tất cả item thuộc `itemMap` và con của chúng đã khớp chính xác.

### 2.5. P1 — Khắc phục lỗi đọc `TableField` trong VL02N (`ModuleSAP.bas`)
- Trong `TableField`:
  * Khi `visibleRow` vượt quá số dòng control thực tế hoặc không tìm thấy trường, hàm trả về `Nothing` an toàn thay vì ném lỗi runtime 4205.
- Trong `OverviewSnapshot` và `VisibleItemRow`:
  * Kiểm tra `If fld Is Nothing Then GoTo finished` / `Exit For` an toàn, ngăn ngừa hoàn toàn lỗi Runtime 91 và 4205.

### 2.6. P1 — Đối chiếu An toàn Draft 3001 & Mở rộng Phạm vi Ngày Planned GI (`ModuleSAP.bas`, `ModuleMain.bas`)
- Trong `SAPX_FindAndReleaseDraft3001OB`:
  * Đọc thêm cột STO (`cSTO` / `VGBEL`) và Creator (`cCreator` / `ERNAM`).
  * Hệ thống tính điểm ứng viên: Ưu tiên ứng viên do chính người dùng hiện tại tạo (+1000 điểm) và tham chiếu đúng STO mục tiêu (+500 điểm). Tuyệt đối tránh xóa nhầm draft của người khác.
- Trong `OB_ReadVL06O` và `SAPX_FindAndReleaseDraft3001OB`:
  * Mở rộng `IT_WADAT-HIGH` từ `Date` thành `DateAdd("d", 60, Date)` để bao quát cả các OB có Planned GI trong tương lai 60 ngày.
- Trong `EDCPlanOne`:
  * Ưu tiên chọn STO có Due đáp ứng trọn vẹn nguyên batch (`EDCWholeBatchFitQty > 0`) trước khi chọn theo tiêu chí STO chủ đạo.
  * Nếu mã đã có OB tạo một phần nhưng `missing > 0`: Cho phép tiếp tục lập kế hoạch tạo phần còn thiếu khi STO có thêm Due, không bỏ dở.
- Trong `PrepareStock`:
  * Đồng bộ kiểm tra loại kho: Chỉ chặn khi không nằm trong `POST_ALLOWED_STOCK_TYPES`.

---

## 3. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
- **Shape Macro Bindings**:
  * `INPUT!OBV1__0` -> `'261010_OB_TOOL_V1_RC56.xlsm'!OB_Check`
  * `INPUT!OBV1__1` -> `'261010_OB_TOOL_V1_RC56.xlsm'!OB_Execute`
- **Khử bỏ `forceFullCalc`**: Bảo đảm `forceFullCalc="1"` bị xóa sạch khỏi `xl/workbook.xml` (`calcPr calcId="181029" calcMode="manual"/>`).
- **Mã băm SHA-256**: Tạo bảng kê SHA-256 chính xác cho workbook và handoff zip package.
