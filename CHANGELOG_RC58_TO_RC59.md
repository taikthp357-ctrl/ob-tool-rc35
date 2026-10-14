# CHANGELOG - OB TOOL V1 RC59

## 1. Tổng quan Bản phát hành RC59 (Phân Định Thông Điệp Ghi Chú & Xóa Toàn Bộ OB Nháp 3001 Chưa Có SLoc)
- **Mục tiêu**: Thực thi chỉ đạo của Người dùng:
  * Phân định rõ thông điệp ghi chú Cột D Sheet `INPUT`:
    - Xóa cả chứng từ đơn mã: `"DA XOA OB THIEU (TAO .../... PCS) DE TRA STO DUE"`.
    - Xóa dòng trên chứng từ đa mã: `"DA XOA DONG ITEM THIEU (... PCS) TREN OB ...; GIU NGUYEN CAC DONG DA DU"`.
  * Nhánh xóa OB để trả tồn bên 3001: Nếu draft OB chứa nhiều mã bên 3001 chưa có SLoc (`sloc = ""`) và chưa ai xử lý (unpicked / `batch = ""`), tự động xóa toàn bộ chứng từ OB (`btn[14]`) để giải phóng toàn bộ tồn kho ATP về STO.

---

## 2. Chi tiết Kỹ thuật Các Thay Đổi

### 2.1. Chuẩn hóa Thông điệp Cột D Sheet `INPUT` (`ModuleMain.bas:4229-4255`)
- Khi xóa cấp dòng item dở dang trên OB đa mã (4007):
  ```vba
  mEDCWaitByMat(CStr(pCleanMat)) = "DA XOA DONG ITEM THIEU (" & pActQty & " PCS) TREN OB " & pTargetOB & "; GIU NGUYEN CAC DONG DA DU"
  ```
- Khi xóa toàn bộ chứng từ OB đơn mã hoặc tất cả các mã đều thiếu:
  ```vba
  delRes = ModuleSAP.SAPX_DeleteDraftOB(pSes, pTargetOB)
  mEDCWaitByMat(CStr(pCleanMat)) = "DA XOA OB THIEU (TAO " & pActQty & "/" & pReqQty & " PCS) DE TRA STO DUE"
  ```

### 2.2. Nhánh Xóa Toàn Bộ Draft OB 3001 Chưa Có SLoc & Chưa Xử Lý (`ModuleSAP.bas`)
- Trong `SAPX_DeleteDraftOB(ses, ob, targetMat, targetItem, isDraft3001Release)`:
  * Duyệt qua toàn bộ item của `OverviewSnapshot`: kiểm tra nếu tất cả các dòng đều chưa có SLoc (`LIPS-LGORT = ""`) và chưa ai xử lý (`LIPS-CHARG = ""` / chưa gán batch).
  * Khi gọi từ nhánh giải phóng tồn 3001 (`isDraft3001Release = True`): nếu thỏa mãn điều kiện trên, hệ thống bỏ qua nhánh xóa từng dòng và chuyển thẳng sang xóa toàn bộ chứng từ qua `btn[14]`.
  * Ghi log: `DELETE_3001_DRAFT_WHOLE_OB: Draft 3001 OB=... (... items) has NO sloc and NO batch processing; deleting whole OB via btn[14] to restore stock`.
- Trong `SAPX_FindAndReleaseDraft3001OB_Batch` và `SAPX_FindAndReleaseDraft3001OB`:
  * Truyền tham số `isDraft3001Release = True` để kích hoạt cơ chế xóa toàn bộ OB nháp khi hợp lệ.

---

## 3. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: PASS 100% (`before=True`, `after=False`).
- **Shape Macro Bindings**:
  * `INPUT!OBV1__0` -> `'261010_OB_TOOL_V1_RC59.xlsm'!OB_Check`
  * `INPUT!OBV1__1` -> `'261010_OB_TOOL_V1_RC59.xlsm'!OB_Execute`
- **Version Labels**:
  * `INPUT!A1` = `"OB TOOL V1 RC59"`
  * `CONFIG!B2` = `"RC59"`
  * `INPUT!R22` = `"RC59 DRAFT 3001 WHOLE-OB RELEASE & EXPLICIT NOTES"`
  * `INPUT!R25` = `"59.0"`
- **CalcPr Optimization**: Thuộc tính `forceFullCalc="1"` được khử bỏ hoàn toàn.
