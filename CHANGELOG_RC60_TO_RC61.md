# CHANGELOG - OB TOOL V1 RC61

## 1. Tổng quan Bản phát hành RC61 (Vá Triệt Để Lỗi GROUP_RECEIPT_MATERIAL_MISMATCH Khi Auto-Retry & Đồng Bộ Hợp Nhất Chứng Từ OB)
- **Vấn đề thực tế phát hiện trong vận hành RC60**:
  Khi chạy đơn hàng gồm nhiều mã, nếu Lượt 1 (Attempt 1) tạo được một số mã thành công nhưng có một số mã bị thiếu (deficit) do kho 3001 giữ tồn:
  1. Tool đã giải phóng draft 3001 rất tốt và kích hoạt `AUTO_RETRY_VL10D` cho các mã thiếu.
  2. Tuy nhiên, trước khi gọi Create retry, cấu hình `VL10D_GROUP_OB_RECEIPTS` chưa được xóa trắng. Do đó, hàm xác minh đọc lại danh sách chứng từ của Lượt 1 (vốn chứa cả các mã đã đủ, không nằm trong danh sách retry).
  3. Khi hàm `EDCVerifyGroupReceipts` quét qua các dòng chứng từ cũ của Lượt 1, câu lệnh `If Not plans.exists(CStr(fields(2))) Then Err.Raise ... GROUP_RECEIPT_MATERIAL_MISMATCH` đã ném lỗi runtime `-2147217235`, làm ngắt toàn bộ tiến trình retry và ghi đè lỗi lên Cột D sheet `INPUT`.
  4. Ngoài ra, trong Lượt Retry, nếu các mã thiếu được tạo bù trên chứng từ mới, số lượng đã bị gán đè thay vì cộng dồn (`vPlan(3) = vPlan(3) + rPlan(3)`), và danh sách chứng từ giữa 2 lượt chưa được gộp lại đầy đủ.

---

## 2. Chi tiết Kỹ thuật Các Thay Đổi Cốt Lõi (`ModuleMain.bas`)

### 2.1. Xóa Trắng `VL10D_GROUP_OB_RECEIPTS` Trước Khi Retry & Hợp Nhất Sau Retry
- Trước khi submit retry qua VL10D, bảo toàn danh sách chứng từ Lượt 1 vào biến `priorVerifiedReceipts` và gọi `PutConfig "VL10D_GROUP_OB_RECEIPTS", ""` để Lượt Retry tự động kích hoạt `EDCDiscoverReceiptsFromVL06O` quét tìm chứng từ mới sinh.
- Sau khi retry hoàn tất, hợp nhất danh sách chứng từ của cả 2 lượt:
```vba
Dim retryReceipts As String
retryReceipts = Trim$(ModuleHelper.ToolConfig("VL10D_GROUP_OB_RECEIPTS", ""))
If priorVerifiedReceipts <> "" Then
 If retryReceipts <> "" Then
  PutConfig "VL10D_GROUP_OB_RECEIPTS", priorVerifiedReceipts & vbLf & retryReceipts
 Else
  PutConfig "VL10D_GROUP_OB_RECEIPTS", priorVerifiedReceipts
 End If
End If
```

### 2.2. Khử Bẫy Lỗi `GROUP_RECEIPT_MATERIAL_MISMATCH` (Defensive Receipt Filtering)
- Trong `EDCVerifyGroupReceipts`, nếu gặp dòng chứng từ thuộc mã không nằm trong kế hoạch xác minh hiện tại (`Not plans.exists`), hệ thống sẽ nhẹ nhàng bỏ qua (skip) thay vì ném ngoại lệ làm sập luồng:
```vba
' RC61: Bo qua dong receipt cua ma khong thuoc ke hoach xac nhan hien tai (tranh loi MATERIAL_MISMATCH)
If plans.exists(CStr(fields(2))) Then
 ...
End If
```

### 2.3. Cách Ly Candidate Discovery Khỏi OB Cũ Của Lượt 1
- Trích xuất toàn bộ số OB đã tạo ở Lượt 1 (`priorOBs`) và nạp vào `rPlanned(8)` (baseline IDs) cho các mã tham gia retry.
- Đảm bảo `EDCDiscoverReceiptsFromVL06O` nhận diện các OB của Lượt 1 là baseline (`isBaseline = True`) và chỉ khớp chính xác với OB mới sinh từ lệnh retry.

### 2.4. Cộng Dồn Số Lượng Xác Nhận & Làm Sạch Ghi Chú Cột D Sheet `INPUT`
- Khi gộp `retryVerified` vào `verifiedPlans`, cộng dồn số lượng tạo giữa Lượt 1 và Lượt Retry (`vPlan(3) = CDbl(vPlan(3)) + CDbl(rPlan(3))`), đồng thời cập nhật `mEDCCreatePlan`.
- Khi tổng số lượng đạt 100% nhu cầu, tự động xóa ghi chú thiếu (`mEDCWaitByMat.Remove`, `EDCSetInputNote "", ""`), đưa giao diện người dùng về trạng thái sạch đẹp và sẵn sàng cho Nút 2 (Execute).

### 2.5. Phòng Vệ Khóa Trùng Lặp Trong `EDCEnsureGroupInSnapshot`
- Bổ sung kiểm tra `If Not expected.exists(key)` và chỉ theo dõi các key thuộc `plans` để loại bỏ nguy cơ lỗi runtime 457 (key already exists).

---

## 3. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: PASS 100% (`before=True`, `after=False`).
- **Shape Macro Bindings**:
  * `INPUT!OBV1__0` -> `'261010_OB_TOOL_V1_RC61.xlsm'!OB_Check`
  * `INPUT!OBV1__1` -> `'261010_OB_TOOL_V1_RC61.xlsm'!OB_Execute`
- **Version Labels**:
  * `INPUT!A1` = `"OB TOOL V1 RC61"`
  * `CONFIG!B2` = `"RC61"`
  * `INPUT!R22` = `"RC61 AUTO-RETRY LIFECYCLE & RECEIPT MERGE"`
  * `INPUT!R25` = `"61.0"`
- **CalcPr Optimization**: Thuộc tính `forceFullCalc="1"` không tồn tại trong XML.
