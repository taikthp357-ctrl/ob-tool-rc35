# CHANGELOG - OB TOOL V1 RC58

## 1. Tổng quan Bản phát hành RC58 (Khắc Phục 3 Lỗi Ngầm & Tối Ưu Hóa Toàn Diện Vận Hành)
- **Mục tiêu**: Thực thi triệt để chỉ đạo của Người dùng ngày 14/10/2026: *"vậy xử lí nhé"* nhằm khắc phục dứt điểm 3 lỗi ngầm trong Item Delete & Batch 3001 và triển khai 3 điểm tối ưu hóa vận hành, xây dựng trên baseline `261010_OB_TOOL_V1_RC57.xlsm`:
  1. **Khắc phục Lỗi ngầm 1 (Batch Release nhận diện thiếu `DELETE_ITEM_OK`)**:
     - Trong `SAPX_FindAndReleaseDraft3001OB_Batch` (`ModuleSAP.bas:5753`), cập nhật kiểm tra `If delStatus = "DELETE_VERIFIED" Or delStatus = "DELETE_ITEM_OK" Then`.
     - Giúp hệ thống công nhận giải phóng tồn thành công khi xóa 1 item trên OB đa mã (trước đây chỉ so khớp `DELETE_VERIFIED`, dẫn tới hiểu nhầm `DELETE_ITEM_OK` là thất bại và không tăng biến đếm `countSuccess`).
  2. **Khắc phục Lỗi ngầm 2 (Hàm đơn lẻ xóa cả chứng từ đa mã qua `btn[14]`)**:
     - Trong `SAPX_FindAndReleaseDraft3001OB` (`ModuleSAP.bas:5896`), truyền tham số `targetMat` vào hàm `SAPX_DeleteDraftOB(ses, bestOB, targetMat)`.
     - Cập nhật kiểm tra `If delStatus = "DELETE_VERIFIED" Or delStatus = "DELETE_ITEM_OK" Then`.
     - Ngăn chặn hoàn toàn rủi ro bấm nhầm `btn[14]` xóa sạch cả chứng từ khi gặp draft OB có nhiều mã hàng.
  3. **Khắc phục Lỗi ngầm 3 (Dòng Item dở dang bị kẹt vĩnh viễn trên OB đa mã)**:
     - Trong `PARTIAL_STO_CLEANUP` (`ModuleMain.bas:4221`), khi `obHasCompletedMat = True`, trước đây tool chỉ ghi log bỏ qua, khiến dòng thiếu bị kẹt trên OB và khóa dung lượng Due của STO.
     - Nay kích hoạt `SAPX_DeleteDraftOB(pSes, pTargetOB, CStr(pCleanMat))` để xóa riêng dòng item bị thiếu, trả Due về STO và bảo toàn nguyên vẹn 100% các mã đã tạo đủ.
  4. **Tối ưu 1 (Tự động xác nhận Pop-up sau khi Save trên VL02N)**:
     - Sau lệnh Save `btn[11]` trong xóa Item, bổ sung bẫy đóng pop-up cảnh báo/thông tin `wnd[1]` bằng phím Enter (VKey 0), tránh treo luồng SAP.
  5. **Tối ưu 2 (Short-circuit thoát sớm trong `EDCReleaseDrafts3001`)**:
     - Kiểm tra ngay từ đầu: nếu tổng số lượng thiếu $\le 0$, hàm kết thúc ngay lập tức, tiết kiệm tài nguyên và không mở VL06O dư thừa.
  6. **Tối ưu 3 (Chuẩn hóa thông điệp trạng thái cột D sheet `INPUT`)**:
     - Phân định rõ ràng giữa:
       * Xóa dòng Item trên OB đa mã: `"DA XOA DONG ITEM THIEU TRONG OB (TAO ... PCS): TRA LAI STO DUE; CAC MA KHAC GIU NGUYEN"`.
       * Xóa toàn bộ chứng từ OB: `"DA XOA TOAN BO OB THIEU (TAO ... PCS): TRA LAI STO DUE"`.

---

## 2. Chi tiết Kỹ thuật Các Thay Đổi (Code Diff Analysis)

### 2.1. `ModuleSAP.bas`
- **Hàm `SAPX_DeleteDraftOB`**:
  * Bổ sung cơ chế tự động đóng pop-up cảnh báo sau khi bấm Lưu `btn[11]`:
    ```vba
    If Not SAPX_FindById(ses, "wnd[1]") Is Nothing Then
     ses.findById("wnd[1]").sendVKey 0
     SAPX_WaitIdle ses, 10
    End If
    ```
  * Chuẩn hóa giá trị trả về khi xóa thành công cấp item: `SAPX_DeleteDraftOB = "DELETE_ITEM_OK"`.
- **Hàm `SAPX_FindAndReleaseDraft3001OB_Batch`**:
  * Mở rộng điều kiện kiểm tra kết quả xóa:
    ```vba
    delStatus = SAPX_DeleteDraftOB(ses, bestOB, mat)
    If delStatus = "DELETE_VERIFIED" Or delStatus = "DELETE_ITEM_OK" Then
     countSuccess = countSuccess + 1
     ModuleLog.WriteLog "DRAFT_3001_RELEASE_OK", ...
    ```
- **Hàm `SAPX_FindAndReleaseDraft3001OB`**:
  * Truyền `targetMat` để kích hoạt nhánh xóa Item:
    ```vba
    delStatus = SAPX_DeleteDraftOB(ses, bestOB, targetMat)
    If delStatus = "DELETE_VERIFIED" Or delStatus = "DELETE_ITEM_OK" Then
     SAPX_FindAndReleaseDraft3001OB = bestOB & "|" & CStr(bestQty)
    ```

### 2.2. `ModuleMain.bas`
- **Khối `PARTIAL_STO_CLEANUP`**:
  * Khi `obHasCompletedMat = True`, kích hoạt xóa riêng dòng item bị thiếu:
    ```vba
    delItemRes = ModuleSAP.SAPX_DeleteDraftOB(pSes, pTargetOB, CStr(pCleanMat))
    If delItemRes = "DELETE_VERIFIED" Or delItemRes = "DELETE_ITEM_OK" Then
     ModuleLog.WriteLog "VL10D_PARTIAL_ITEM_CLEANUP", ...
     If verifiedPlans.exists(CStr(pCleanMat)) Then verifiedPlans.Remove CStr(pCleanMat)
     mEDCWaitByMat(CStr(pCleanMat)) = "DA XOA DONG ITEM THIEU TRONG OB (TAO " & pActQty & "/" & pReqQty & " PCS): TRA LAI STO DUE; CAC MA KHAC GIU NGUYEN"
     EDCSetInputNote CStr(pCleanMat), CStr(mEDCWaitByMat(pCleanMat))
     PutConfig "VL10D_DIRECT_DONE_" & CStr(pCleanMat), ""
     ThisWorkbook.Save
    ```
  * Cập nhật nhánh xóa toàn bộ OB (`obHasCompletedMat = False`) với thông điệp: `"DA XOA TOAN BO OB THIEU (TAO " & pActQty & "/" & pReqQty & " PCS): TRA LAI STO DUE"`.
- **Hàm `EDCReleaseDrafts3001`**:
  * Bổ sung short-circuit kiểm tra tổng số lượng thiếu:
    ```vba
    Dim totalDeficit As Double, kDef As Variant
    For Each kDef In deficits.keys
     totalDeficit = totalDeficit + Val(deficits(kDef))
    Next kDef
    If totalDeficit <= 0 Then Exit Function
    ```

---

## 3. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: PASS 100% (`before=True`, `after=False`).
- **Shape Macro Bindings**:
  * `INPUT!OBV1__0` -> `'261010_OB_TOOL_V1_RC58.xlsm'!OB_Check`
  * `INPUT!OBV1__1` -> `'261010_OB_TOOL_V1_RC58.xlsm'!OB_Execute`
- **Version Labels**:
  * `INPUT!A1` = `"OB TOOL V1 RC58"`
  * `CONFIG!B2` = `"RC58"`
  * `INPUT!R22` = `"RC58 LATENT BUGS RESOLVED & OPTIMIZED"`
  * `INPUT!R25` = `"58.0"`
- **CalcPr Optimization**: Thuộc tính `forceFullCalc="1"` được khử bỏ hoàn toàn.
- **4 Logistics Invariants**:
  * Invariant 1 (Branch 3001 vs 2001 EWM): Tuân thủ nghiêm ngặt.
  * Invariant 2 (FIFO): Bảo toàn thứ tự STO và batch.
  * Invariant 3 (WholeBatchGate): Chỉ xuất nguyên batch khi đủ điều kiện.
  * Invariant 4 (Idempotency): Tái chạy an toàn, bảo vệ chứng từ đã tạo đủ.
