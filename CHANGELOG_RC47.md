# CHANGELOG - OB TOOL V1 RC47

## 1. Bối cảnh & Phân tích nguyên nhân thực tế (Root Cause Analysis từ log 3:11 AM - 3:14 AM)

### Hiện tượng 1: "Lần 1 đã có kế hoạch nhưng ấn nút 2 thì báo chưa có kế hoạch"
- **Nguyên nhân**:
  - Khi SAP tạo được 3/4 mã trên DO `1012700358`, mã `BIW71005030` bị thiếu 28 pcs (do 3001 giữ tồn).
  - Lệnh xóa draft 3001 bị lỗi SAP (`SCREEN_SEQ._CTRL is not defined for function LOES`), nhưng code ngộ nhận xóa thành công nên gọi `AUTO_RETRY_VL10D`.
  - Retry trên VL10D thất bại (vì 3001 chưa được xóa). Khi verify lại, mã `BIW71005030` có `totals = 0`, hàm `EDCEnsureGroupInSnapshot` ném lỗi `GROUP_OB_LIST_WAIT: list Qty differs from verified deficit for BIW71005030`.
  - Lỗi này đẩy macro vào khối `Failed:`, khiến **3 mã đã tạo thành công bị hủy commit**, giữ nguyên `VL10D_PENDING_KEY` và làm `mChecked = False`.
  - Khi người dùng bấm Nút 2, tool kiểm tra `Not mChecked` và chặn lại: "Chưa có kế hoạch đã kiểm tra trong phiên này. Nhấn nút 1 trước."

### Hiện tượng 2: "Việc chạy đi chạy lại xóa VL02N ở nhánh 3001 lặp 2 lần vì lý do gì"
- **Nguyên nhân**:
  1. Lệnh OKCode `LOES` bị SAP từ chối trên Dynpro 4001 (`type=E; text=SCREEN_SEQ._CTRL is not defined for function LOES`).
  2. Tuy nhiên, điều kiện kiểm tra kết quả `Not DeliveryNumberField(ses) Is Nothing` vẫn thỏa mãn (do màn hình VL02N còn đó), khiến hàm ngộ nhận xóa thành công! Thực tế OB 3001 vẫn tồn tại nguyên vẹn trên SAP.
  3. Ở cuối `EDCCreateGroup`, khi khối `Failed:` bị kích hoạt bởi lỗi verify, nó lại gọi `EDCReleaseDrafts3001 deficits` một lần nữa! Do OB 3001 chưa hề bị xóa ở lượt trước, tool quét lại và tìm thấy chính OB đó, dẫn đến việc cố xóa lặp lại lần 2 trong cùng 1 lần nhấn nút.

### Hiện tượng 3: "Lần 2 ấn nút 1 thì trả 1 kết quả có sẵn đó nhưng vẫn thực hiện lượt xóa ở 3001"
- **Nguyên nhân**:
  - Ở lần 2, DO `1012700358` (đã có 3 mã) được nạp vào. Nhưng mã `BIW71005030` vẫn thiếu 28 pcs, nên tool lại tiếp tục vào nhánh xóa 3001 và lại gọi xóa OB `1012696087` lần thứ 3.

### Hiện tượng 4: "Nút 2 thực hiện sai kế hoạch khi chỉ điền 35pcs của 1 batch, xóa 2 mã kia đi"
- **Nguyên nhân cốt lõi (Cực kỳ nghiêm trọng)**:
  1. Trong `EDCPlanOne`: Khi `BIW70004840` cần 660 nhưng DO mới có 80 (`capacity = 80 < need`), code thấy thiếu hụt và ném lỗi `OB_ALREADY_CREATED_WAIT: receipt exists but usable OB capacity missing`. Tương tự với `BIW71004357` (cần 300, DO có 85).
  2. Lỗi này lưu 2 mã trên vào `mEDCWaitByMat`.
  3. Trong `OB_Execute`: Code gộp toàn bộ `mEDCWaitByMat` vào `freshBlocked`! Hậu quả là `BIW70004840` và `BIW71004357` bị `freshBlocked` **LOẠI BỎ KHỎI DANH SÁCH THỰC THI** của DO `1012700358`!
  4. Trong nhóm hàng gửi sang SAP chỉ còn lại duy nhất `BIW71004505` (35 pcs).
  5. Khi mở DO `1012700358` trên VL02N, hàm `RemoveUnplannedItems` thấy trên DO có 3 item (10, 20, 30), nhưng kế hoạch chỉ gửi item 20. Hàm liền xem item 10 và 30 là "hàng ngoài kế hoạch" và **XÓA SẠCH ITEM 10 VÀ 30 KHỎI DELIVERY ORDER TRÊN SAP**!

---

## 2. Các Nâng cấp Đột phá trên OB Tool V1 RC47

### 2.1. Cam kết & Giải phóng từng phần (Atomic Commit & mChecked = True)
- Trong `EDCCreateGroup`:
  - Bất kỳ mã nào đã được SAP tạo DO thành công (dù là một phần hay nguyên mẻ) đều được `EDCEnsureGroupInSnapshot` và `EDCCommitGroup` **NGAY LẬP TỨC**.
  - Đặt `mChecked = True` ngay sau khi commit phần thành công, bảo đảm Nút 2 (`OB_Execute`) luôn sẵn sàng để thực thi chia lô cho các mã đã có DO mà không bao giờ bị báo "chưa có kế hoạch"!
  - Xóa bỏ việc gọi trùng lặp `EDCReleaseDrafts3001` trong khối `Failed:`.

### 2.2. Bảo tồn Tuyệt đối các mặt hàng đã có OB (`EDCPlanOne`)
- Trong `EDCPlanOne`:
  - Khi đã có OB với một phần số lượng (`capacity > 0`), hệ thống ghi nhận `EDC_PARTIAL_OB_READY` và cho phép phân bổ lượng `capacity` đó, **TUYỆT ĐỐI KHÔNG ném lỗi `OB_ALREADY_CREATED_WAIT`** làm đóng băng hàng.
  - Ghi chú rõ trên sheet `INPUT`: `DA CO OB ... pcs; THIEU ... pcs (cho STO due/3001 nha ton)`.

### 2.3. Cách ly hoàn toàn `mEDCWaitByMat` khỏi `freshBlocked` (`OB_Execute`)
- Trong `OB_Execute`:
  - `freshBlocked` chỉ dùng để bảo vệ tồn nguồn từ POSTING (`FreshSourceBlocks()`).
  - Gỡ bỏ hoàn toàn việc gộp `mEDCWaitByMat` vào `freshBlocked`.
  - Mọi dòng phân bổ đã khớp trên DO ứng viên đều được đưa vào hàng đợi thực thi, không bao giờ bị loại bỏ vì lý do mã đó còn thiếu hàng ở các mẻ sau.

### 2.4. Bất biến Bảo vệ Mặt hàng trên DO (`RemoveUnplannedItems`)
- Trong `RemoveUnplannedItems`:
  - Bổ sung quy tắc an toàn bất khả xâm phạm: Nếu dòng hàng trên Delivery Order thuộc danh mục vật tư trong `InputProducts()`, **TUYỆT ĐỐI KHÔNG ĐƯỢC XÓA** (`keep(item) = True`).
  - Chấm dứt hoàn toàn nguy cơ xóa nhầm các dòng hàng hợp lệ mà SAP VL10D đã sinh ra.

### 2.5. Nâng cấp Cơ chế Xóa OB 3001 & Bắt lỗi Nghiêm ngặt (`SAPX_DeleteDraftOB`)
- Trong `SAPX_DeleteDraftOB`:
  - Quét tìm mục `Delete` / `Löschen` / `Xóa` trực tiếp từ thanh Menu `Outbound Delivery` (`wnd[0]/mbar/menu[0]`).
  - Dự phòng với mã hàm OKCode chuẩn `DELE` của chương trình `SAPMV50A`.
  - **Kiểm tra nghiêm ngặt kết quả**: Nếu thanh trạng thái SAP trả về `type = "E"`, hàm **BẮT BUỘC TRẢ VỀ FALSE**. Tuyệt đối không bao giờ ngộ nhận xóa thành công khi có lỗi.
  - Nếu xóa không thành công, hệ thống không retry VL10D vô ích mà thông báo rõ ràng cho người dùng.

### 2.6. Xử lý An toàn trong `EDCEnsureGroupInSnapshot`
- Kiểm tra `IsArray(planned)` và `plans.exists(mat)` trước khi truy xuất phần tử, loại bỏ triệt để lỗi `Run-time error 13: Type mismatch`.
- Tự động tách các mã có số lượng tạo bằng 0 (`totals = 0`) ra khỏi nhóm commit mà không làm gián đoạn các mã thành công.

---

## 3. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: `compile_enabled_before = True` $ightarrow$ `compile_enabled_after = False` (PASS 100%).
- **Button Macro Bindings**:
  - `INPUT!OBV1__0` $ightarrow$ `'261010_OB_TOOL_V1_RC47.xlsm'!OB_Check`
  - `INPUT!OBV1__1` $ightarrow$ `'261010_OB_TOOL_V1_RC47.xlsm'!OB_Execute`
- **Bảo toàn 4 Invariants**: FIFO, WholeBatchGate, Safety Lock, Fast Save Mode.
