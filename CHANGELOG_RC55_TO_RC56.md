# CHANGELOG - OB TOOL V1 RC56

## 1. Tổng quan Bản phát hành RC56 (Vá Schema Pending Key 6 Trường, Nhánh Tạo 1 Nửa STO / Tự Động Xóa OB Thiếu, Gỡ Bỏ Presave Extra Item VL02N)
- **Mục tiêu**: Khắc phục triệt để các lỗi ngầm và hoàn thiện quy trình nghiệp vụ ghi nhận trong các lượt chạy live thực tế từ 13/10 đến 14/10/2026:
  1. Chuẩn hóa schema `VL10D_PENDING_KEY` 6 trường (`Request | Material | Qty | STO | Item | Timestamp`), sửa deserializer `EDCResolvePendingGroup` để không bị từ chối `Invalid pending item identity` và không tự xóa pending khi chưa xác minh.
  2. Bảo toàn `VL10D_PENDING_KEY` trong quy trình `AUTO_RETRY_VL10D`, ngăn chặn việc xóa pending sớm trong `EDCCommitGroup` trước khi lượt tạo bổ sung hoàn tất.
  3. Bổ sung nhánh xử lý khi VL10D chỉ tạo được một phần STO ("tạo 1 nửa STO"): Nếu nhánh 3001 không giữ tồn hoặc không thể sinh thêm để đạt nguyên batch, tự động xóa OB tạo thiếu ("xóa tạo bừa đó") qua `SAPX_DeleteDraftOB` per `deleteob.vbs` để hoàn trả Due về STO.
  4. Gỡ bỏ bẫy `PRESAVE_GATE_FAIL: EXTRA_ITEM` trong VL02N và `SavedSnapshotMatchesPlan`: Cho phép lưu OB khi trên chứng từ có các item khác chưa có kế hoạch trong lượt chạy hiện tại.
  5. Sửa lỗi `TableField` đọc tràn dòng gây lỗi runtime 4205 (`Thieu field LIPS-POSNR row 3`).
  6. Đối chiếu an toàn khi tìm xóa Draft 3001 (`VGBEL` & `ERNAM`), mở rộng phạm vi ngày Planned GI tương lai trong VL06O, và ưu tiên STO nhận đủ nguyên batch.
  7. Native VBE Debug Compile PASS 100%, commit Git tag `v1.0.0-rc56`, push repo và xuất bản GitHub Release `v1.0.0-rc56`.
- **Bảo toàn Invariants**: Giữ nguyên vẹn 100% logic của Branch 3001 Body Shop, Branch 2001 Shop XMĐ (/SCWM/MON EWM Dispatcher), WholeBatchGate và FIFO allocation.

---

## 2. Chi tiết Cải tiến & Nghiệp vụ Mới

### 2.1. P0 — Vá lỗi Schema Deserializer `VL10D_PENDING_KEY` 6 trường (`ModuleMain.bas`)
- Trong `EDCCreateGroup`:
  * Chuẩn hóa dữ liệu pending 6 trường: `reqId & "|" & mat & "|" & planned(3) & "|" & sto & "|" & planned(2) & "|" & timestamp`.
- Trong `EDCResolvePendingGroup` và `EDCResolvePendingCreate`:
  * Chấp nhận `UBound(parts) >= 4` thay vì bắt buộc `= 4`, tương thích 100% với cả format 5 trường và 6 trường có timestamp.
  * Trích xuất an toàn các trường và không bao giờ ném lỗi `Invalid pending item identity`.
- Cơ chế bảo vệ Pending Key:
  * Tuyệt đối không xóa `VL10D_PENDING_KEY` khi gặp lỗi đọc tạm thời; chỉ xóa khi đã xác minh thành công trên SAP hoặc vật tư nằm ngoài danh sách INPUT.

### 2.2. P0 — Bảo toàn Pending Key trong quy trình `AUTO_RETRY_VL10D` (`ModuleMain.bas`)
- Trong `EDCCreateGroup`:
  * Khi gọi `EDCCommitGroup verifiedPlans`: Thêm cờ `keepPendingKey`, không xóa pending key của nhóm khi vẫn còn vật tư thiếu (`deficits.count > 0` và có retry).
  * Trước khi gọi `EDC4007_CreateGroup(sto)` trong `AUTO_RETRY_VL10D`: Thiết lập lại `VL10D_GROUP_PENDING` và `VL10D_PENDING_KEY` cho danh sách `retryPlans`.
  * Trong `EDCVerifyGroup`: Thêm cơ chế fallback xác định STO/Request từ danh sách kế hoạch, loại bỏ hoàn toàn nguy cơ ném lỗi `GROUP_RECEIPT_MISSING: pending identity invalid`.

### 2.3. P0 — Bổ sung Nhánh Xử lý Tạo 1 Nửa STO & Tự Động Xóa OB Thiếu (`ModuleMain.bas`)
- Logic nghiệp vụ:
  * Khi VL10D chỉ tạo được một phần số lượng STO (`actualQty < plannedQty`):
    - Đã cố gắng giải phóng qua `EDCReleaseDrafts3001(deficits)`.
    - Nếu không tìm thấy draft 3001 nào (`releasedDraft = False`, nhánh 3001 không giữ) HOẶC sau khi retry mà SAP vẫn không thể tạo thêm (hết tồn kho thực tế / hết ATP):
    - Kiểm tra cấu hình `AUTO_DELETE_PARTIAL_STO_OB` (mặc định = 1):
      + Tự động gọi `ModuleSAP.SAPX_DeleteDraftOB(ses, partialOB)` để xóa bỏ OB dở dang này theo đúng `deleteob.vbs`.
      + Ghi log `VL10D_PARTIAL_OB_CLEANUP`: *"Material=...; Successfully deleted partial OB=... (... pcs); STO ... Due restored"*.
      + Loại bỏ mã khỏi danh sách `verifiedPlans` để không ghi nhận trạng thái ảo.
      + Ghi chú rõ trên sheet INPUT: `"DA XOA OB THIEU (TAO " & actualQty & "/" & plannedQty & " PCS): HET TON & 3001 KO GIU TON; TRA LAI STO DUE"`.

### 2.4. P1 — Gỡ bỏ Bẫy `PRESAVE_GATE_FAIL: EXTRA_ITEM` trong VL02N (`ModuleSAP.bas`)
- Hiện tượng thực tế trên OB `1012736970`: Chứng từ chứa 4 line items (10, 20, 30, 40) do STO tạo ra. Lượt chạy hiện tại chỉ phân bổ cho 10, 30, 40 (item 20 chưa có kế hoạch hoặc không thuộc yêu cầu lần này). Hàm `VerifyDeliveryPlan` ném lỗi `PRESAVE_GATE_FAIL: EXTRA_ITEM 20` và dừng lưu.
- Khắc phục:
  * Trong `VerifyDeliveryPlan`: Chuyển việc kiểm tra sang log thông tin `PRESAVE_EXTRA_ITEM_INFO`, không ném ngoại lệ dừng lưu; kiểm tra nghiêm ngặt 100% tất cả các item có trong kế hoạch `itemMap`.
  * Trong `SavedSnapshotMatchesPlan`: Cho phép lưu khi tồn tại item chưa có kế hoạch trên delivery overview.

### 2.5. P1 — Khắc phục Lỗi Đọc `TableField` trong VL02N (`ModuleSAP.bas`)
- Trong `TableField`:
  * Trả về `Nothing` an toàn thay vì ném ngoại lệ runtime 4205 (`Thieu field ... Can MAP layout VL02N.`).
  * Trong `OverviewSnapshot`: Khi `TableField` trả về `Nothing`, dừng vòng lặp duyệt dòng (`GoTo finished`) an toàn.

### 2.6. P1 — Đối Chiếu An Toàn Draft 3001 & Mở Rộng Phạm Vi Ngày Planned GI (`ModuleSAP.bas`, `ModuleMain.bas`)
- Trong `SAPX_FindAndReleaseDraft3001OB`:
  * Đọc thêm cột STO (`cSTO`) và Creator (`cCreator`).
  * Chấm điểm ứng viên: Ưu tiên tuyệt đối xóa draft do chính user tạo ra (`currentUser`), tránh xóa nhầm draft hợp lệ của người khác.
- Trong `OB_ReadVL06O` và `SAPX_FindAndReleaseDraft3001OB`:
  * Mở rộng `IT_WADAT-HIGH` từ `Date` thành `DateAdd("d", 60, Date)` để bao quát cả các OB có Planned GI trong tương lai.
- Trong `EDCPlanOne`:
  * Ưu tiên chọn STO có Due đáp ứng trọn vẹn nguyên batch trước khi chọn STO chủ đạo.
  * Nếu mã đã có biên nhận tạo một phần nhưng `missing > 0`: Tiếp tục lập kế hoạch tạo phần còn thiếu khi STO có thêm Due, không `Exit Sub` bỏ dở.
- Trong `PrepareStock`:
  * Đồng bộ: Chỉ chặn loại kho khi loại kho đó không nằm trong `POST_ALLOWED_STOCK_TYPES`.

---

## 3. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: PASS 100% (`control.Enabled` chuyển từ `True` sang `False`, không cảnh báo hay lỗi cú pháp).
- **Shape Macro Bindings**:
  * `INPUT!OBV1__0` -> `'261010_OB_TOOL_V1_RC56.xlsm'!OB_Check`
  * `INPUT!OBV1__1` -> `'261010_OB_TOOL_V1_RC56.xlsm'!OB_Execute`
- **Khử bỏ `forceFullCalc`**: Bảo đảm `forceFullCalc="1"` bị xóa sạch khỏi `xl/workbook.xml` (`calcPr calcId="181029"/>`).
- **Mã băm SHA-256**: Tạo bảng kê SHA-256 chính xác cho workbook và handoff zip package.
