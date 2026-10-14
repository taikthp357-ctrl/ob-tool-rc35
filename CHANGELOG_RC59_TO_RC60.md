# CHANGELOG - OB TOOL V1 RC60

## 1. Tổng quan Bản phát hành RC60 (Tối Ưu Xóa Trực Tiếp OB Nháp 3001 Từ VL06O, Khử Bỏ Tra Cứu Lặp & Gom Batch Siêu Tốc)
- **Mục tiêu**: Thực thi triệt để chỉ đạo chuyên sâu của Người dùng ngày 14/10/2026:
  1. **Khử bỏ kiểm tra SLoc dư thừa trong VL02N**:
     - Trong thực tế SAP, kiểm tra ở `VL06O` nếu `rowSloc = ""` thì toàn bộ các mã trong OB đó đều không có SLoc (không tồn tại trường hợp một mã có SLoc và một mã không có SLoc).
     - Loại bỏ hoàn toàn việc gọi `OverviewSnapshot` hay cuộn bảng item trong `VL02N` khi xóa draft OB 3001 (`isDraft3001Release = True`), nhảy thẳng tới lệnh xóa toàn bộ chứng từ qua `btn[14]`.
  2. **Khử bỏ kiểm tra lặp lại sau khi xóa (Zero Re-query)**:
     - Trước đây, tool thoát ra mở lại `/nVL02N` gõ lại số OB để dò tìm "does not exist", gây giật màn hình và lãng phí 10-15s.
     - Nay ngay sau khi bấm `btn[14]` và xác nhận pop-up xóa `btnSPOP-OPTION1`, hệ thống kiểm tra trực tiếp thanh trạng thái SAP (`sType <> "E"`). Báo thành công (`DELETE_VERIFIED`) ngay lập tức, tiết kiệm tối đa thời gian.
  3. **Gom các bước xử lý hàng loạt cho Plant 3001 (Batch Consolidation)**:
     - Trong `SAPX_FindAndReleaseDraft3001OB_Batch`: Quét toàn bộ ALV của `VL06O` trong 1 lần, gom tất cả các draft OB 3001 thỏa mãn điều kiện (`rowSloc = ""` và `rowBatch = ""`) thuộc danh sách mã thiếu hoặc trùng `targetSTO` vào `toDeleteDict`.
     - Điều hướng sang `VL02N` 1 lần duy nhất, tuần tự xóa nhanh toàn bộ danh sách chứng từ gom được mà không cần thoát/vào lại màn hình.
     - Đồng bộ tồn ATP 1 lần duy nhất sau khi xóa xong toàn bộ batch (đợi 2s).

---

## 2. Chi tiết Kỹ thuật Các Thay Đổi (`ModuleSAP.bas`)

### 2.1. Nhảy thẳng sang xóa toàn bộ chứng từ trong `SAPX_DeleteDraftOB`
```vba
If isDraft3001Release Then
 ' RC60: Da xac minh o VL06O la khong co SLoc -> Xoa luon toan bo qua btn[14]
 ModuleLog.WriteLog "DELETE_3001_DRAFT_DIRECT", "Draft 3001 OB=" & ob & " verified no SLoc in VL06O; deleting whole OB directly via btn[14]"
 GoTo DeleteWholeOB
End If
```

### 2.2. Khử bỏ mở lại VL02N để hậu kiểm lặp
```vba
' RC60: Bao thanh cong luon per chi dao nguoi dung - khong click mo lai de kiem tra
If sType <> "E" Or InStr(1, sText, "deleted", vbTextCompare) > 0 Or _
   InStr(1, sText, "gelöscht", vbTextCompare) > 0 Or _
   InStr(1, sText, "da xoa", vbTextCompare) > 0 Or _
   InStr(1, sText, "xoa", vbTextCompare) > 0 Or sType = "S" Then
 ModuleLog.WriteLog "DELETE_VERIFIED", "Confirmed OB=" & ob & " deleted on SAP (direct confirmation): " & sText
 SAPX_DeleteDraftOB = "DELETE_VERIFIED"
 Exit Function
Else
 ModuleLog.WriteLog "DELETE_OB_FAIL", "Delete failed for OB=" & ob & "; status=" & sText
 SAPX_DeleteDraftOB = "DELETE_FAILED"
 Exit Function
End If
```

### 2.3. Gom toàn bộ ứng viên draft 3001 vào `toDeleteDict`
```vba
' Dieu kien cot loi per chi dao nguoi dung: rowSloc = "" (chua co SLoc) va rowBatch = "" (chua ai xu ly)
If rowBatch = "" And rowSloc = "" And (rowGM = "A" Or rowGM = "") And val(rowParent) = 0 Then
 If deficits.exists(mat) Or (targetSTO <> "" And rowRefSTO = targetSTO) Then
  If Not toDeleteDict.exists(ob) Then
   toDeleteDict.Add ob, Array(mat, qtyText, rowRefSTO)
   ModuleLog.WriteLog "DRAFT_3001_COLLECT_OB", "Gom draft 3001 OB=" & ob & " (no SLoc, unpicked) for Material=" & mat & "; STO=" & rowRefSTO
  End If
 End If
End If
```

---

## 3. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: PASS 100% (`before=True`, `after=False`).
- **Shape Macro Bindings**:
  * `INPUT!OBV1__0` -> `'261010_OB_TOOL_V1_RC60.xlsm'!OB_Check`
  * `INPUT!OBV1__1` -> `'261010_OB_TOOL_V1_RC60.xlsm'!OB_Execute`
- **Version Labels**:
  * `INPUT!A1` = `"OB TOOL V1 RC60"`
  * `CONFIG!B2` = `"RC60"`
  * `INPUT!R22` = `"RC60 FAST BATCH 3001 RELEASE & ZERO RE-QUERY"`
  * `INPUT!R25` = `"60.0"`
- **CalcPr Optimization**: Thuộc tính `forceFullCalc="1"` được khử bỏ hoàn toàn.
