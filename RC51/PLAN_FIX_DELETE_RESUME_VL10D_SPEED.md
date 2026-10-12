# Kế hoạch sửa RC51 — Delete draft 3001, OB đã bị xóa và tốc độ VL10D
Ngày 13/10/2026. Phạm vi: nhánh 4007 và recovery draft 3001. Tư liệu: workbook người dùng RC51 (giữ LOG), deleteob.vbs. Chưa sửa workbook trong lần bàn giao này.

## 1. Bằng chứng và yêu cầu đã chốt

- Lượt 20261013_050454 tạo OB 1012719513: sáu mã đủ; BIW32211830AA cần200, tạo102, thiếu98.
- Recovery tìm OB 3001 1012709673 Qty98, nhưng DELETE_OB_RESULT type/text trống; báo RELEASE_FAIL rồi NOT_FOUND sai nghĩa.
- Lượt 20261013_050930 bị OB_ALREADY_CREATED_WAIT ở bảy mã dù Due hợp lệ; kết thúc 0 allocations.
- Người dùng xác nhận đã xóa OB sinh chưa đủ trước lượt kiểm tra lại. Cần đối chiếu đúng ID đã lưu để chuyển receipt thành OB_DELETED_CONFIRMED và lập lại kế hoạch phần chưa hoàn tất, không chặn mã vĩnh viễn.
- Lỗi blocked trước đó do có người mở STO; phải tách WAIT_STO_LOCK khỏi thiếu ATP, không cancel 3001 để chữa khóa.
- EA là đơn vị mục tiêu theo cấu hình; PCS tương đương 1:1 theo nghiệp vụ người dùng. Đơn vị khác lấy lượng EA quy đổi thực; không ép lượng nguồn thành EA khi chưa biết tỷ lệ.
- Lỗi424 chọn mảng/Is Nothing đã qua trong lượt RC51; không lấy lỗi cũ làm nguyên nhân lượt mới.

## 2. P0: đường Delete theo VBS

Bản ghi bắt đầu khi đã ở màn hình thích hợp, gồm:
1. wnd[0].sendVKey 0
2. wnd[0]/tbar[1]/btn[14].press
3. wnd[1]/usr/btnSPOP-OPTION1.press

Tích hợp vào SAPX_DeleteDraftOB sau khi mở chính xác OB cần xóa:
- Xác nhận tcode/màn hình, số OB trên header và các item hiện tại. Không hiểu VBS là tự mở OB; bản ghi không chứa bước mở.
- Đối chiếu draft: ship-to3001, mã cần giải phóng, Batch/SLoc trống, trạng thái phù hợp. Nếu xóa cả delivery, kiểm tra mọi item; OB hỗn hợp có item đã phân không được xóa cả.
- Enter theo bản ghi, chờ SAP sẵn sàng, tìm btn14 trên màn hình đã xác minh.
- Nhận diện popup xác nhận Delete, kiểm tra nội dung trước bấm SPOP-OPTION1; không bấm OPTION1 của popup không rõ.
- Ghi status type/text và popup; xử lý thêm popup chỉ khi đã nhận diện đúng. Bản ghi không có Save riêng, không tự thêm Save/nhấn Delete lần nữa để đoán.
- Hậu kiểm bằng mở/tra đúng OB. Không còn ở VL06O chưa đủ chứng minh xóa nếu có filter/layout/sync lỗi; dùng trạng thái không tồn tại/đã xóa của chính ID hoặc nguồn chứng minh tương đương.
- Kết quả: DELETE_VERIFIED, DELETE_REJECTED, DELETE_UNKNOWN, WAIT_OB_LOCK, NO_CANDIDATE. Thay NOT_FOUND giả bằng DELETE_FAILED khi đã chọn ứng viên mà xóa lỗi.
- Nếu timeout sau xác nhận, chỉ read-back trước retry. Đọc lại nguồn khả dụng/ATP liên quan sau xóa; không coi Delete thành công tự chứng minh tồn đã đồng bộ.
- Giữ batch plan của mã khác; chỉ phục hồi nhóm thiếu.

## 3. P0: receipt phải cho chạy tiếp sau OB bị xóa

Vị trí kiểm tra: EDCPlanOne, EDCExistingCapacity, ResolveSaveReceipts, EDCResolvePendingGroup, EDCCommitGroup và VL10D_DIRECT_DONE_<mat>.

- Receipt thành công Create lưu ID OB/item, Request, STO/item, Material, Qty thực đã tạo và attempt. DIRECT_DONE là bằng chứng đã tạo lúc trước, không phải lượng hiện tại hoặc toàn bộ yêu cầu đã hoàn tất.
- Khi receipt còn nhưng usable capacity=0, tự đối chiếu đúng OB/item đã lưu. Không dừng ngay OB_ALREADY_CREATED_WAIT.
- Phân loại:
  * OB_PRESENT_UNALLOCATED: dùng lại cho bước điền, không tạo thêm.
  * OB_PRESENT_ALLOCATED_MATCH: đối chiếu batch/qty với ledger; ghi đúng lượng hoàn tất, không sửa lại.
  * OB_PRESENT_CHANGED: lập lại kế hoạch riêng mã/nhóm bị ảnh hưởng, không chiếm phần đã phân cho người khác.
  * OB_DELETED_CONFIRMED: vô hiệu hóa receipt đang hoạt động của OB bị xóa, giữ lịch sử; trừ phần hoàn tất còn hiệu lực, lập lại phần thiếu theo nguồn live.
  * UNKNOWN_READBACK: chờ xác minh, không tự coi là xóa hay Create tiếp.
- Theo xác nhận của người dùng, OB sinh chưa đủ đã bị xóa: sau khi xác minh, không cộng102 của OB cũ vào hiện tại. Nếu chưa có xuất/hoàn tất nào khác và cần200 thì phần phải tạo lại có thể là200, không mặc định chỉ98.
- Nếu chỉ một item bị xóa hoặc còn OB khác thuộc request, giữ lượng còn hiệu lực của chúng.
- Chặn phát sinh trùng khi người khác vừa tạo OB trong thời gian đối chiếu; cập nhật baseline ngay trước Create.
- Không đổi Request ID hoặc xóa toàn bộ receipt để lách lỗi; cập nhật trạng thái theo từng OB/item.
- Hỗ trợ dữ liệu pending cũ và mới: LOG có Invalid pending item identity sau intent sáu trường có timestamp. Kiểm tra serializer/parser ở phiên bản đích; version hóa schema, đọc định dạng đã biết, giữ UNKNOWN đối với bản ghi không đọc được; không xóa pending chưa rõ để cho Create lại.
- Giữ chính sách nguyên batch: partial102/200 không tự coi là đủ batch200; phần thực đã tạo, phần được phép apply và lượng còn thiếu phải là số riêng.

## 4. P1: tối ưu VL10D trước Background

Hiện tượng người dùng: điền và tick từng mã, sau đó delay trước Background. Cần đo, không quy tất cả chậm cho SAP.

### Đo theo nhóm
VL10D_TIMING: open_selection, multi_filter, execute_due, scan_due, build_plan, locate_rows, write_qty, select_rows, preflight_readback, persist_intent, background_submit, sap_wait, receipt_readback. Ghi số mã, số trang, số lần Execute/scroll và session. Tách thao tác Excel/Save khỏi thời gian SAP.

### Thay đổi
- Giữ một lượt đọc nhiều mã × tối đa4 STO cho planning; lookup bằng STO/item/Material/Plant trong RAM.
- Chỉ reload trước Create khi màn hình không còn đúng scope, nguồn đã thay hoặc cần refresh; cache không thay thế preflight live.
- Map các dòng mục tiêu một lần trên màn hình hiện tại. Với Classic, nhóm thao tác theo trang: tới mỗi trang một lần, ghi Qty/tick các dòng trên trang đó rồi đọc lại nhóm, không cuộn đi/về từng mã.
- Với ALV dùng selectedRows danh sách nhiều dòng một lần sau ghi Qty, nếu control hỗ trợ. Classic checkbox riêng thì vẫn tick từng control nhưng không thêm idle wait cố định sau mỗi tick.
- Gom read-back mục tiêu; không chạy lại toàn bộ scanner cho từng mã. Bảo đảm toàn bộ selection không chứa dòng ngoài kế hoạch, kể cả checkbox còn từ lượt trước và trang không nhìn thấy.
- Loại sleep cố định khi chỉ gán ô/checkbox mà SAP chưa gửi request. Chờ Busy/state khi Enter/Execute/scroll/Background thực sự cần; timeout có giới hạn.
- Nếu hai vòng staging hiện có lặp write/scroll/read, hợp nhất phần trùng; trước Background vẫn kiểm tra đúng STO/item/mã/qty/UoM và selection.
- Lưu intent/pending một lần ngay trước Background cho cả nhóm; gom ghi CONFIG/LOG theo vùng, không Save từng mã. Không bỏ checkpoint trước tác vụ tạo.
- Khi scope/Qty thay đổi, cập nhật đủ nguyên batch và lượng đã có; không giảm kiểm chứng nghiệp vụ để lấy tốc độ.
- Background một lần mỗi nhóm STO đã chọn. Có thể ưu tiên STO chủ đạo nhưng chỉ trên các ứng viên đáp ứng qty/whole-batch; không cộng Due các STO khác thành một.
- Nếu STO khóa: đóng popup có nhận diện, đọc kết quả Create từng item, giữ OB đã tạo; deferred STO xuống cuối, retry có giới hạn sau kiểm tra lượng mới. Không gọi recovery3001 chỉ vì blocked.
- Không thay nhánh3001 chính/2001 trong patch tối ưu4007.

## 5. Ca nghiệm thu

| Ca | Kỳ vọng |
|---|---|
| Draft1012709673 Qty98, hợp lệ | Mở đúng OB, btn14 → SPOP-OPTION1, hậu kiểm xóa, đọc nguồn sau |
| Popup không phải xác nhận Delete | Không bấm OPTION1 mù; log nội dung |
| OB bị khóa hoặc đã có item phân hàng | Không xóa, trả trạng thái riêng |
| Timeout sau Delete | Hậu kiểm trước thao tác lại |
| OB1012719513 đã bị người dùng xóa | Xác minh đúng ID, loại receipt hiệu lực, lập lại lượng thiếu thực |
| Không có trên VL06O do filter/sync | UNKNOWN, không nhận xóa giả |
| Một item xóa, item khác còn | Chỉ lập lại phần bị mất |
| OB đã điền đúng batch | Giữ kết quả, không tạo thêm |
| Pending sáu trường timestamp | Parser đọc đúng hoặc báo schema cụ thể; không clear pending mù |
| Partial102/batch200 | Không apply102 như batch nguyên; không retry98 khi OB102 đã bị xóa |
| 7–20 mã nhiều trang | Một lượt planning; staging theo trang; một Background/nhóm |
| Mã trùng nhiều batch | Gom qty theo mã cho Create; giữ phân bổ batch khi Apply |
| STO đang có người mở | WAIT_STO_LOCK, deferred; không delete3001 |
| SAP tạo một phần trước lock | Giữ receipt đúng OB mới và chỉ retry phần thiếu đã xác minh |
| EA/PCS | Theo quy ước1:1; đọc lại Qty EA sau Enter |
| Refresh đổi vị trí dòng | Tìm lại STO/item/mã, không dùng row index cũ để ghi |
| Nhánh3001 và2001 | Hành vi không đổi do patch4007 |

## 6. Tiêu chí giao
- Compile Excel/VBE thực, mở không repair, nút macro bind đúng.
- LOG phải chứng minh trạng thái Delete, phục hồi receipt và lượng thiếu.
- So sánh timing trước/sau trên số mã/trang tương đương; không báo cải thiện tốc độ nếu chưa đo.
- ZIP/hash phải khớp đúng workbook đã kiểm tra; ghi rõ các ca live chưa chạy.
- Bản workbook ở thư mục này là RC51 người dùng gửi để chẩn đoán, chưa phải bản đã sửa/QA lại.
