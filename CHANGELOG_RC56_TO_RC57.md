# CHANGELOG - OB TOOL V1 RC57

## 1. Tổng quan Bản phát hành RC57 (Rà Soát Toàn Diện & Khép Kín Vòng Lặp - Bảo Vệ Đa Tầng Chống Văng Lỗi Cơ Bản)
- **Mục tiêu**: Thực thi triệt để chỉ đạo P0 của Người dùng ngày 14/10/2026: *"vậy cho hệ thống rà soát và khép kín lại vòng tránh văng lỗi ở các bước và khâu cơ bản"* theo hồ sơ `BRIEF_RC57_CLOSE_THE_LOOP_SAFEGUARDS.md` trên baseline `261010_OB_TOOL_V1_RC56.xlsm`:
  1. **Cơ chế Checkpoint State Machine**: Giữ nguyên kết quả các mã đã tạo thành công (`VERIFIED_FULL`), khóa ghi chú cột D sheet `INPUT`, bỏ qua tạo VL10D. Chỉ tính toán số lượng bù cho các mã thiếu (`Deficit > 0`). Khi đổi Input (`Fingerprint(False) != saved`), tự động Clean Reset 100% (xóa checkpoint cũ, xóa D6:D205, F6:N505, sinh New Request).
  2. **Item-Level Delete trên VL02N**: Sửa đổi logic hủy trong `VL02N` (`ModuleSAP.bas`). Với chứng từ OB đa mã (>= 2 items), chỉ xóa dòng vật tư bị thiếu qua lệnh Menu `Edit -> Delete Item` (`Shift+F2`), xác nhận popup và bấm Save `btn[11]`. Bảo toàn 100% các dòng mã đã đủ trên chứng từ. Chỉ xóa toàn bộ chứng từ qua `btn[14]` khi OB chỉ có đúng 1 item.
  3. **Gom Batch Xóa Draft 3001 & Quét Nhả Tồn Chủ Động Khi Due = 0**: Gom toàn bộ danh sách mã thiếu vào 1 lần mở `VL06O` duy nhất (`SAPX_FindAndReleaseDraft3001OB_Batch`), nhập Multiple Selection của `MATNR`, quét ALV 1 lượt, sang `VL02N` xóa tuần tự và đợi 2s đồng bộ ATP kho 1 lần duy nhất. Kích hoạt quét nhả tồn chủ động ngay trong `EDCPlanOne` khi phát hiện STO Due = 0 hoặc Due < Need.
  4. **Tối ưu VL10D - Loại bỏ mở lặp 2 lần & Tận dụng F8 Refresh tại chỗ**: Tái sử dụng màn hình Due List từ Lần 1 khi đã mở danh sách với các STO (`canReuseScreen = True`), không thoát ra gõ lại `/nVL10D`. Bổ sung hàm `EDC4007_RefreshDueList` nhấn F8 Refresh (`btn[8]`) trực tiếp trên thanh công cụ của bảng Due List khi cần cập nhật số lượng Due.
  5. **Phòng vệ đa tầng chống văng lỗi Runtime & Bảo toàn Pending Key**: Trong `EDCVerifyGroup`, cơ chế chờ mềm (retry đọc lại VL06O sau 2s) và tuyệt đối không gọi `Err.Raise` làm sập luồng sang khối `Failed:` khi chưa thấy receipt, cho phép nhánh xóa 3001 tiếp tục chạy khép kín vòng lặp. Bảo toàn `VL10D_PENDING_KEY` trong `EDCCommitGroup` khi còn kế hoạch `AUTO_RETRY_VL10D`.

---

## 2. Chi tiết 5 Khâu Cải tiến Cốt lõi

### 2.1. Khâu 1: Checkpoint State Machine (`ModuleMain.bas`)
- Trong `EnsureRequestIdentity`:
  * Khi `changed = True` (`Fingerprint(False) != saved`):
    - Tự động sinh `NEW_REQUEST`.
    - Gọi `ClearConfigPrefix` xóa sạch các khóa `VL10D_DIRECT_DONE_*`, `VL10D_BASELINE_IDS_*`, `VL10D_INTENT_*`, `VL10D_PENDING_KEY`, `VL10D_GROUP_PENDING`, `VL10D_GROUP_OB_RECEIPTS`.
    - Xóa trắng vùng kết quả `INPUT!D6:D205` và `INPUT!F6:N505`.
    - Ghi log: `NEW_REQUEST: INPUT thay doi (Clean Reset 100%)`.
- Hàm mới `EDCReconcileInputNotes`:
  * Đối chiếu từng vật tư trên sheet `INPUT`:
    - Nếu `capacity >= need`: Khóa trạng thái `VERIFIED_FULL`, bảo toàn nguyên vẹn ghi chú cột D.
    - Nếu `capacity > 0` và `< need`: Ghi nhận `DA CO OB ... pcs; THIEU ... pcs`.
    - Nếu `capacity = 0`: Xóa trắng ghi chú cũ để sẵn sàng cho lượt chạy mới.
- Trong `BuildLivePlan` và `BuildPlan`:
  * Thay thế lệnh xóa trắng mù quáng `ws.Range("D6:D205").ClearContents` bằng lời gọi `EDCReconcileInputNotes`.
- Trong `EDCPlanOne`:
  * Khi `capacity >= need`: Ghi log `EDC_EXISTING_OB_READY: VERIFIED_FULL, bo qua VL10D` và `Exit Sub`, không đưa vào danh sách tạo mới VL10D.
  * Khi `capacity < need`: Chỉ tính toán và lập kế hoạch tạo đúng phần thiếu `missing = need - capacity`.

### 2.2. Khâu 2: Item-Level Delete trên VL02N (`ModuleSAP.bas`)
- Nâng cấp `SAPX_DeleteDraftOB(ses, ob, Optional targetMat, Optional targetItem)`:
  * Sau khi mở chứng từ trong `VL02N`, chụp snapshot lưới qua `OverviewSnapshot(ses, False)`.
  * **Trường hợp Multi-item OB (`snap.count >= 2`)**:
    - Định vị chính xác dòng tuyệt đối `targetAbsRow` chứa `targetMat` hoặc `targetItem`.
    - Cuộn bảng và chọn dòng: `tbl.getAbsoluteRow(targetAbsRow).selected = True`.
    - Chọn lệnh Menu `Edit -> Delete Item` (`wnd[0]/mbar/mp[1]/mi[4]` hoặc phím tắt `Shift+F2`).
    - Xác nhận cửa sổ Pop-up xóa Item (`wnd[1]/usr/btnSPOP-OPTION1`).
    - Nhấn nút Lưu (Save - `btn[11]`).
    - Xác nhận chứng từ lưu thành công per `IsSaveReceipt`, ghi log `DELETE_ITEM_OK` kèm số lượng item còn lại được bảo toàn nguyên vẹn.
    - Trả về `"DELETE_ITEM_OK"`.
  * **Trường hợp Single-item OB (`snap.count <= 1`)**:
    - Gọi lệnh Toolbar Delete `btn[14]` per `deleteob.vbs`.
    - Xác nhận popup, hậu kiểm trong `VL02N` xác nhận chứng từ đã xóa hoàn toàn.
    - Trả về `"DELETE_VERIFIED"`.

### 2.3. Khâu 3: Gom Batch Xóa Draft 3001 & Quét Nhả Tồn Chủ Động (`ModuleSAP.bas`, `ModuleMain.bas`)
- Hàm mới `SAPX_FindAndReleaseDraft3001OB_Batch(deficits, targetSTO)`:
  * Mở `VL06O` 1 lần duy nhất: Ship-to = `3001`, SLoc = `3093`, WBSTK = `A`, Planned GI window (-60 đến +60 ngày).
  * Nhập toàn bộ danh sách mã thiếu vào Multiple Selection của `MATNR` qua ô `SEARCH_LIST` và nút `btn[24]` (Paste from clipboard).
  * Nhấn F8 (`btn[8]`) thực thi.
  * Quét ALV trong 1 lượt duy nhất, tính điểm ưu tiên và gom danh sách draft OB tối ưu cho từng mã thiếu.
  * Chuyển sang `VL02N` xóa tuần tự theo danh sách đã gom (tự động phân nhánh item-level delete nếu là multi-item OB).
  * Chờ 2s đồng bộ ATP kho 1 lần duy nhất cho toàn bộ batch.
- Hàm `SAPX_FindAndReleaseDraft3001OB` đơn lẻ tái sử dụng trực tiếp qua hàm Batch.
- Trong `EDCReleaseDrafts3001`:
  * Thay thế vòng lặp đơn lẻ bằng 1 lời gọi `SAPX_FindAndReleaseDraft3001OB_Batch`.
- Kích hoạt chủ động trong `EDCPlanOne`:
  * Khi phát hiện STO Due = 0 hoặc tổng Due < Need do nhánh 3001 chiếm giữ: Tự động gọi `SAPX_FindAndReleaseDraft3001OB_Batch` để giải phóng tồn ngay tại bước này. Sau đó làm mới Due list để lấy dung lượng tạo VL10D mà không bị ngắt luồng.

### 2.4. Khâu 4: Tối ưu VL10D - Tái Sử Dụng Màn Hình & F8 Refresh Tại Chỗ (`ModuleSAP.bas`)
- Trong `EDC4007_PrepareGroup`:
  * Mở rộng điều kiện tái sử dụng màn hình: Khi `mEDCSelectedSTO` chứa `sto` hoặc là chuỗi đa STO chứa `sto` (`InStr(";" & mEDCSelectedSTO & ";", ";" & sto & ";") > 0`), hệ thống nhận diện toàn bộ các dòng của STO mục tiêu đã hiển thị sẵn trên màn hình từ Lần 1.
  * Tái sử dụng trực tiếp `mEDCLastDue` và màn hình đang mở, ghi log `VL10D_SCREEN_REUSED`, loại bỏ hoàn toàn việc thoát ra gõ lại `/nVL10D`.
- Hàm mới `EDC4007_RefreshDueList`:
  * Gửi lệnh nhấn F8 Refresh (`wnd[0]/tbar[1]/btn[8]` hoặc `wnd[0]/tbar[0]/btn[8]`) trực tiếp trên thanh công cụ của màn hình Due List VL10D.
  * Chờ xử lý và đọc lại kết quả lưới ALV tại chỗ, không thoát ra màn hình Selection Screen.

### 2.5. Khâu 5: Phòng Vệ Đa Tầng Runtime & Bảo Toàn Pending Key (`ModuleMain.bas`)
- Trong `EDCVerifyGroup`:
  * Thêm cơ chế chờ mềm: Khi chưa thấy dòng receipt ngay (`records = ""`), tự động chờ 2s và đọc lại `VL06O` tối đa 2 lần trước khi kết luận.
  * Khi sau các lượt chờ vẫn chưa có receipt: Tuyệt đối không gọi `Err.Raise` làm sập luồng sang khối `Failed:`. Thay vào đó, ghi nhận toàn bộ mã vào `defDict`, ghi log `VL10D_GROUP_NO_RECEIPTS` và thoát hàm an toàn (`Exit Sub`). Luồng tiếp tục chuyển sang bước xóa draft 3001 và retry bình thường.
- Trong `EDCCreateGroup`:
  * Xác định cờ `willRetry = (releasedDraft And deficits.count > 0)`.
  * Truyền `keepPendingKey = willRetry` vào `EDCCommitGroup`.
  * Đảm bảo `VL10D_PENDING_KEY` được bảo toàn nguyên vẹn cho lượt `AUTO_RETRY_VL10D`.

---

## 3. Bằng chứng Kiểm định & Xác nhận Bàn giao
- **Native VBE Debug Compile**: PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
- **Gắn Macro Nút Bấm**:
  * `INPUT!OBV1__0` -> `'261010_OB_TOOL_V1_RC57.xlsm'!OB_Check`
  * `INPUT!OBV1__1` -> `'261010_OB_TOOL_V1_RC57.xlsm'!OB_Execute`
- **Khử bỏ `forceFullCalc`**: Loại bỏ hoàn toàn thuộc tính `forceFullCalc="1"` trong `xl/workbook.xml`.
- **Đóng gói Handoff**: `RC57_HANDOFF.zip` chứa đầy đủ Workbook, diff, changelog, mã nguồn VBA và SHA-256 manifest.
