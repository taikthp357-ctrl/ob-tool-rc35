# CHANGELOG - OB TOOL V1 RC46

## Tổng quan nâng cấp RC46
Bản nâng cấp **OB Tool V1 RC46** giải quyết triệt để vấn đề:
1. **"Lượt nào ra lượt đó"**: Không bị ngộ nhận, block chùm hay xóa nhầm ứng viên OB có sẵn khi có một hoặc vài mã còn thiếu due hoặc chưa tạo được trên VL10D.
2. **"Tự can và tự tạo lại" (Auto-Cancel 3001 Draft & Immediate Auto-Retry VL10D)**: Tự động xóa draft OB của 3001 chiếm giữ tồn tại 3093 và **lập tức gọi tạo lại ngay trên VL10D** trong cùng một phiên xử lý, giúp sinh đủ DO và phân bổ trọn vẹn.
3. **Khắc phục lỗi xóa Draft OB VL02N**: Thay thế lệnh `Shift+F12` (VKey 24 không khả dụng trên Dynpro 4001 gây lỗi `The virtual key is not enabled`) bằng SAP function code chuẩn `LOES` trong OKCode field, toolbar `btn[14]` và Menu `Outbound Delivery -> Delete`.

---

## Chi tiết các thay đổi kỹ thuật

### 1. `ModuleMain.bas` -> `EDCCandidateAllowed` & `OB_ClassifyCandidates`
- **Trước RC46**: Khi kiểm tra `EDCCandidateAllowed(mat, sto)`, nếu `mEDCWaitByMat.exists(mat)` thì trả về `False`. Điều này khiến các Delivery Order đã tạo từ trước (ví dụ DO `1012699408` có 7 mặt hàng) bị `OB_ClassifyCandidates` đánh dấu `BLOCK_OB: EDC STO/MATERIAL SCOPE BLOCK` trên cả DO, làm tê liệt toàn bộ 7 mặt hàng sẵn sàng giao!
- **RC46**: Loại bỏ điều kiện chặn `mEDCWaitByMat` trong `EDCCandidateAllowed`. Phân tách phạm vi đánh dấu: chỉ block riêng item vi phạm phạm vi STO (`BLOCK_ITEM: EDC STO/MATERIAL SCOPE BLOCK`), tuyệt đối không lây nhiễm làm hỏng cả DO (`BLOCK_OB`).

### 2. `ModuleMain.bas` -> `EDCRefreshGroupOB`
- **Trước RC46**: Khi gọi `OB_ReadVL06O`, hàm chỉ truyền danh sách các mặt hàng vừa tạo (`plans`), dẫn tới VL06O chỉ lọc riêng mã đó và vô tình làm mất dữ liệu các mã ứng viên khác đang lưu trên sheet `VL06O_DATA`.
- **RC46**: Truyền toàn bộ `InputProducts()` để nạp đầy đủ tất cả các mặt hàng thuộc phạm vi yêu cầu, đồng thời gọi `OB_ClassifyCandidates ws` ngay sau khi gộp (`MergeUntouched`) nhằm duy trì trạng thái `CANDIDATE` hợp lệ của các mặt hàng sẵn có.

### 3. `ModuleSAP.bas` -> `SAPX_DeleteDraftOB`
- **Trước RC46**: Sử dụng `ses.findById("wnd[0]").sendVKey 24` (`Shift+F12`). Tuy nhiên trên giao diện Dynpro 4001 màn hình tổng quan VL02N, phím ảo VKey 24 không được bật, phát sinh lỗi: `DELETE_OB_FAIL: OB=...; The virtual key is not enabled`.
- **RC46**: 
  - Điều hướng OKCode field (`wnd[0]/tbar[0]/okcd`) với mã hàm chuẩn SAP `LOES` + Enter.
  - Dự phòng với nút thùng rác toolbar `wnd[0]/tbar[0]/btn[14]`.
  - Dự phòng menu `Outbound Delivery -> Delete` (`wnd[0]/mbar/menu[0]/menu[6]`).
  - Xử lý xác nhận popup xóa tự động và kiểm tra chuỗi trạng thái thành công.

### 4. `ModuleMain.bas` -> `EDCCreateGroup` ("Tự can và tự tạo lại")
- **Trước RC46**: Khi phát hiện thiếu hụt (deficits), công cụ xóa draft 3001 xong thì dừng lại và yêu cầu người dùng phải bấm nút 1 thủ công lần nữa.
- **RC46**: 
  - Khi `EDCReleaseDrafts3001` giải phóng thành công draft OB của 3001, hệ thống chuyển sang giai đoạn `AUTO_RETRY_VL10D`.
  - Tự động chuẩn bị lại danh sách các mã thiếu hụt (`retryPlans`), tự động gọi tạo lại ngay trên `VL10D` (`EDC4007_PrepareGroup` & `EDC4007_CreateGroup`).
  - Nạp lại kết quả vào `VL06O_DATA` và gộp vào kế hoạch commit trong cùng phiên làm việc.
  - Trường hợp lượt 1 đã có DO sẵn có và lượt 2 cần xử lý tiếp: toàn bộ dữ liệu ứng viên của DO sẵn có được bảo toàn nguyên vẹn, đảm bảo "lượt nào ra lượt đó" - người dùng bấm Nút 2 (`OB_Execute`) là hoàn tất lượt 1 trơn tru.

---

## Bằng chứng kiểm thử & Tiêu chuẩn phát hành
- **Native VBE Debug Compile**: `compile_enabled_before = True` -> `compile_enabled_after = False` (PASS 100%).
- **Bảo toàn 4 Invariants**: Toàn vẹn WholeBatchGate, FIFO, Fast Save, và cấu trúc 16 cột (A:P).
- **Trọng lượng & Tốc độ**: LOG được thu gọn về 5 dòng header, LEDGER thu gọn về 1 dòng header, workbook mở và lưu cực nhanh.
