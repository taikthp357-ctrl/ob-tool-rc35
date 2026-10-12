# Kế hoạch sửa nhánh 4007 — receipt, giải phóng OB 3001 và UoM VL10D

Ngày: 13/10/2026. Cơ sở: đọc VBA workbook release v1.0.0-rc49 và lịch sử chạy người dùng gửi từ 01:41 tới 04:14 ngày 13/10/2026. Đây là kế hoạch sửa, chưa phải mã đã sửa hay bản phát hành đã nghiệm thu. Nếu triển khai trên phiên bản mới hơn, đối chiếu mã hiện tại trước khi áp dụng.

## 1. Mục tiêu và phạm vi

Giữ luồng hai nút: nút 1 kiểm tra Posting, VL06O và chuẩn bị/tạo OB thiếu cho 4007; nút 2 điền Qty/SLoc/Batch cho OB đã xác minh. Giữ đọc gom nhiều mã và tối đa 4 STO. Chỉ sửa đường 4007 và bộ giải phóng draft 3001 được 4007 gọi; không thay thuật toán nhánh 3001 chính hoặc nhánh EWM 2001.

“OB trống” ở đây là OB chưa phân batch/SLoc, còn lượng và chưa GI/pick; không phải OB Qty=0. Hủy OB giải phóng lượng giữ trên hệ thống; không mặc định đó là thao tác chuyển vật lý về bin. Cần đọc lại nguồn liên quan để chứng minh tồn có thể dùng.

## 2. Bằng chứng và mức chắc chắn

| Lượt | Quan sát chắc chắn | Ý nghĩa / việc cần xác minh |
|---|---|---|
| 01:42, Request REQ_20261013_014158_1 | 2 mã, 3 STO; VL10D eligible rows=0; WAIT_STO_DUE_4007; không có Create | Đường hiện tại thoát trước vị trí gọi release 3001. Không có Due không tự chứng minh 3001 giữ tồn |
| 01:50 | Receipt absent, sau đó xác minh 2 mã và commit Materials=2 | Là đối chứng luồng đọc kết quả tạo OB |
| 01:54 và 04:06 | Receipt present=True; COMMITTED Materials=0 rồi GROUP_DEFERRED; không có RELEASE_3001_START | Kết quả báo cáo mâu thuẫn; phải ghi lỗi verify. Receipt cũ/sai scope là giả thuyết có cơ sở từ mã, chưa được LOG chứng minh nguyên nhân cuối cùng |
| 04:06, BIW32022638AB | Kế hoạch tạo 20; ảnh sau đó có OB 4007 1012719113 lượng 3; các dòng 3001 trống Batch/SLoc tổng 80 | Cần đối chiếu đúng OB/item và lượng mới sinh; nếu 3 là lượng thuộc lượt này thì thiếu 17. Chưa chứng minh ATP sẽ nhả khi xóa |
| 04:13 và 04:14 | Uom coordinate=70; đọc một số candidate rồi CHECK UoM không khớp | Chưa ghi actual/expected và dòng lỗi; không đủ kết luận hàng khác EA hoặc chắc chắn layout lệch |

## 3. Thứ tự sửa

### P0 — sửa xác minh và dữ liệu thiếu trước khi cho phép cancel

Vị trí chính: ModuleMain.EDCCreateGroup, EDCVerifyGroup, EDCCommitGroup và EDCResolvePendingGroup.

1. Lưu intent tạo OB bất biến: Request, STO/item, Material, Qty định tạo, UoM, baseline OB/item và thời điểm. Không dùng dictionary plans bị sửa làm nguồn duy nhất cho tổng cần tạo.
2. Receipt gắn đúng Request + STO + lượt Create. Receipt trước phải lưu lịch sử riêng; không tái sử dụng nếu sai scope. Khi scope cũ còn chưa rõ kết quả, đối chiếu read-only trước; không xóa pending chỉ để cho phép Create lại.
3. Bỏ On Error Resume Next bao quanh verify/commit/retry có thể che lỗi nghiệp vụ. Trả kết quả theo từng mã: VERIFIED_FULL, VERIFIED_PARTIAL, VERIFIED_ZERO, UNKNOWN_READBACK, RECEIPT_SCOPE_MISMATCH, SOURCE_READ_ERROR.
4. Không lấy Qty kế hoạch thay Qty thực đọc khi OB không đọc được hoặc Qty<=0. Không nhận actual>=planned là đủ mà bỏ qua phần dư: actual>planned phải kiểm tra trùng và giữ trạng thái sai lệch.
5. Tính thiếu theo từng mã và STO/item trên lượng thực đã xác minh, loại OB baseline và OB/item trùng. Không đếm cả parent batch-split và các child để tránh double.
6. Commit riêng phần đã xác minh; bảo toàn lượng cần tạo gốc, lượng đã sinh và phần thiếu. Không báo COMMITTED Materials=0 như thành công.
7. UNKNOWN_READBACK không được biến thành thiếu tồn/ATP và không kích hoạt hủy 3001 hoặc tự Create lại.
8. Khi có nhiều STO, lượng mỗi STO/item được giữ độc lập; không trừ hoặc sinh trùng cùng nhu cầu giữa các nhóm.

### P1 — đưa recovery 3001 vào đúng các đường quyết định

Vị trí chính: EDCPlanOne, EDCCreateGroup, EDCReleaseDrafts3001, SAPX_FindAndReleaseDraft3001OB và SAPX_DeleteDraftOB.

| Trạng thái 4007 | Hành động |
|---|---|
| Due đọc hợp lệ, có STO/item đủ điều kiện; Create thiếu hoặc không sinh và đã xác minh kết quả | Kiểm tra nguyên nhân ATP/giữ lượng; tìm draft 3001 liên quan theo đúng mã |
| Due=0, đọc nguồn đầy đủ nhưng chưa rõ lý do | Chẩn đoán STO/item còn mở, điều kiện ngày/plant và bằng chứng ATP; có thể đọc draft 3001, chưa tự hủy |
| STO thiếu lượng, đóng/không đến hạn hoặc mã không có trong STO | Báo WAIT_STO; không dùng cancel 3001 làm cách chữa |
| Lỗi layout/UoM/receipt hoặc kết quả chưa rõ | Chờ sửa/đối chiếu; không cancel |
| Có OB 4007 đủ nhưng Batch/SLoc đã phân | Đối chiếu với request/history để tránh tạo thêm; không coi mọi dòng BLOCK là thiếu lượng mới |

Quy tắc ứng viên draft 3001:
- Đúng ship-to 3001, shipping point/plant nguồn theo cấu hình và đúng Material.
- Batch và SLoc trống, chưa GI; kiểm tra picking/EWM distribution/WT và lượng thực trước hành động.
- Nếu chỉ xóa item, định danh đúng OB/item. Nếu dùng xóa cả delivery, phải đọc toàn bộ item và chứng minh toàn bộ OB phù hợp; không xóa cả OB hỗn hợp chỉ vì một dòng trống.
- Loại OB có người khác đã phân hàng hoặc thay đổi trạng thái giữa đọc và thao tác.
- Lập danh sách ứng viên và lượng có thể nhả, chọn lượng phù hợp phần thiếu; không chỉ một OB gần nhất nếu chưa đủ, không xóa thêm khi đã đủ.
- Log lý do chọn/bỏ từng ứng viên; kiểm tra lại live ngay trước xóa.
- Đọc lại xác nhận hủy và lượng khả dụng/ATP phù hợp sau hủy. Không coi nút Delete/Save hoặc không còn dòng là bằng chứng đủ về tồn.
- Chờ đồng bộ có giới hạn; chỉ đọc lại các mã liên quan. Khi chưa nhả báo WAIT_RELEASE_SYNC.
- Sau giải phóng, trừ mọi OB 4007 mới sinh trong thời gian chờ rồi lập kế hoạch phần thiếu còn lại theo nguyên batch.
- Chỉ retry Create có intent/receipt mới được kiểm soát, giới hạn một lượt phục hồi cho nhóm; lỗi không rõ kết quả thì đối chiếu trước retry.

### P2 — sửa đọc Classic UoM và cô lập lỗi từng mã

Vị trí chính: ModuleSAP.EDC_ClassicColumn, EDC_ClassicValue, EDC_ClassicDue và đường ALV UoM.

1. Bắt buộc log raw value và normalized value, expected từ CONFIG, Material, STO/item, page, row, tọa độ x và control ID tại điểm đọc.
2. Chọn cột UoM theo header/field và xác minh ô dữ liệu thực tại dòng hiện tại. Sau đổi trang/scroll kiểm tra layout; không dựa riêng tọa độ tiêu đề để mặc định mọi dòng.
3. Phân biệt UOM_CELL_MISSING, UOM_EMPTY, UOM_COLUMN_AMBIGUOUS và UOM_MISMATCH. Ô thiếu/trống không được gọi là đơn vị khác.
4. Chuẩn hóa chữ hoa, khoảng trắng đầu/cuối và ký tự khoảng trắng ẩn được xác định. Không thay đơn vị UN/EA hay đơn vị khác bằng nhau khi chưa có bảng chuyển đổi được chốt.
5. Không ép expected=EA để chữa lỗi đọc; không bỏ gate đơn vị.
6. Lỗi đọc ở một mã được tổng hợp theo mã; các mã khác chỉ tiếp tục nếu dữ liệu của chúng đủ và chính sách tạo nhóm cho phép. Nếu nguồn chung không đọc đầy đủ hoặc trạng thái dòng không xác định, dừng nhóm thay vì giả định đầy đủ.
7. Thu thập đầy đủ lỗi read-only trước Create, không vừa đọc vừa tạo khi một phần scope còn chưa rõ.
8. Ghi mapping header/control theo page để kiểm chứng giả thuyết tọa độ 70 bị sai. Không sửa cứng sang tọa độ mới nếu chưa có bằng chứng.

## 4. Log bắt buộc

- VL10D_VERIFY_ERROR: Err.Number, stage, Request/STO, receipt scope và mã liên quan.
- VL10D_MATERIAL_RESULT: planned, actual_verified, remaining, trạng thái kết quả và nguồn chứng minh.
- RELEASE_3001_DECISION: có/không gọi, lý do, deficit đã xác minh và AUTO_DELETE_DRAFT_3001 thực tế.
- DRAFT_3001_CANDIDATE / REJECTED / RECHECK / RELEASE_VERIFIED.
- RELEASE_SOURCE_READBACK: lượng trước/sau, nguồn đọc, timestamp và trạng thái đồng bộ.
- VL10D_UOM_READ / READ_ERROR: raw, normalized, expected, Material, STO/item, page/row/x/control.
- GROUP_SUMMARY: số mã thành công, thiếu, chưa rõ, bị bỏ do lỗi; không chỉ Materials=0.
- PARTIAL_BATCH_WAIT: phần thiếu chưa đủ để tạo nguyên batch.

## 5. Ca nghiệm thu

| Ca | Kết quả bắt buộc |
|---|---|
| 2 mã × 3–4 STO | Đọc nguồn gom; giữ đúng từng STO/item, không vòng lặp mã × STO |
| Receipt cũ từ request trước | Không commit giả; log scope mismatch; đối chiếu lượt hiện tại |
| Create một mã 20 nhưng OB mới đúng chỉ 3 | Ghi actual=3, deficit=17; giữ OB đã sinh; xét recovery 3001 đúng nguyên nhân |
| Create trả summary nhưng không đọc được OB | UNKNOWN_READBACK, không tự cancel hoặc Create thêm |
| Không có Due do STO không còn mở | WAIT_STO, không hủy 3001 |
| Không có Due và có draft 3001 | Đọc chẩn đoán trước; chỉ hủy khi xác minh điều kiện và liên hệ lượng giữ |
| Draft 3001 có item khác đã phân batch | Không xóa cả OB; chỉ item phù hợp nếu đường xóa item được xác minh |
| Draft bị người khác sửa sau khi chọn | Bỏ ứng viên, không tác động theo snapshot cũ |
| Hai draft 3001 cần phối hợp lượng | Chọn có giới hạn; không xóa dư và đối chiếu nguồn sau mỗi hành động |
| Xóa thành công nhưng ATP chưa đồng bộ | WAIT_RELEASE_SYNC; không retry Create ngay |
| Lượng OB 4007 xuất hiện trong lúc chờ | Trừ lượng đó khỏi nhu cầu trước tạo phần còn thiếu |
| UoM EA thật | Đọc EA, pass khi expected=EA |
| UoM có khoảng trắng ẩn | Log raw; chuẩn hóa có kiểm soát; pass nếu cùng EA |
| UoM cell trống hoặc đọc nhầm | Báo READ_ERROR kèm dòng; không mặc định EA |
| Một mã UoM khác thật | Chặn đúng mã/nhóm liên quan, không đổi đơn vị |
| Parent Qty0 + child batch split | Không cộng parent/child trùng; phân biệt lượng đã xử lý |
| Timeout ngay sau Delete/Create | Đọc lại trạng thái trước thao tác lại; không duplicate |
| Phần thiếu không đủ batch | Giữ phần đã có; chờ sức chứa/nguồn theo quy tắc nguyên batch |
| 3001 chính và 2001 | Không đổi dispatcher/nghiệp vụ do bản sửa 4007 |

## 6. Điều kiện giao bản sửa

- Đọc mã phiên bản đích hiện tại, kiểm tra cấu hình và receipt trước sửa; không khẳng định LOG là của RC49 nếu chưa có VERSION của lượt chạy.
- Compile bằng Excel/VBE thực, kiểm tra bindings nút và mở workbook không repair. Hash/ZIP/XML pass không thay thế compile.
- Chạy PREVIEW với dữ liệu lỗi thực; xác nhận các decision log và UoM raw.
- Nghiệm thu live recovery trên đúng OB/item được chọn, kiểm tra nguồn trước/sau và OB 4007 mới.
- Không ghi PASS cho các ca chưa chạy; ghi rõ giới hạn kiểm chứng.
- Không phát hành dựa trên COMMITTED Materials=0 hoặc summary Create. Release phải chứa workbook đã QA đúng hash.
