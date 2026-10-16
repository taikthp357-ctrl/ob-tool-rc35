# Phân tích nhánh 2001 XMĐ trên nền RC47

Ngày: 12/10/2026. Phạm vi: đọc bốn VBS và cấu trúc/cấu hình workbook RC47; chưa sửa VBA, chưa chạy SAP live. Nhánh 3001 và 4007 được giữ theo trạng thái đã ổn do người dùng xác nhận. Không khẳng định đã kiểm tra toàn bộ mã VBA RC47.

## 1. Kết luận

MON là nơi đọc và điều phối WT; màn hình xử lý ODO và tạo WT của EWM cũng cần tham gia. Bốn VBS đã cung cấp đường thao tác thực tế quan trọng, nhưng chưa đủ để khẳng định mọi hậu quả nghiệp vụ của từng nút. Đặc biệt, xác nhận WT về 0, hủy WT, sửa số lượng ODO và hủy delivery phải là những hành động khác nhau trong thiết kế.

Quy tắc nguyên batch phải kiểm tra trên toàn bộ OB trong phạm vi yêu cầu. Không hủy một WT chỉ vì OB riêng lẻ thiếu batch. Không cộng lẫn WT đã hủy, WT đã xác nhận và WT đang mở; không lấy số lượng tồn khả dụng sau giữ hàng làm số lượng nguyên batch ban đầu.

## 2. Bóc tách từng VBS

| File | Bằng chứng trong script | Kết luận có thể dùng | Chưa chứng minh |
|---|---|---|---|
| mon.vbs | Dòng 16 mở favorite F00077; 17–31 mở node N000000010, lọc S_ERP_H qua multiple selection; 32 và 65–67 gọi N000000096, có selectAll bảng trên | Có đường mở danh sách theo số ERP OB và tải bảng chi tiết từ các dòng chứng từ được chọn | Favorite không phải mã Tcode; script không nhập Warehouse, không xác minh ship-to 2001; clipboard được dùng nhưng nội dung không nằm trong VBS |
| mon.vbs | 33–56 có MATNR, CHARG, VLPLA, VSOLM, NISTM, DOCNO; 57–60 có DOCNO_H | Có các cột vật tư, batch, bin nguồn và số lượng ở bảng WT; có cột chứng từ trên/dưới | Không có TANUM, trạng thái, item, UoM trong thao tác ghi lại; DOCNO và DOCNO_H chưa được đối chiếu với số ERP |
| mon.vbs | 72–84 mở N000000139; S_MATNR=BIW70007707, S_OWNER=3093, S_BATCH=260923-103 | Có đường tra tồn theo mã/Owner/batch trên MON | Tên node Physical/Available Stock chưa có trong script |
| mon.vbs | 87–91 xóa Owner và batch; 95–115 có LGTYP, INC_QUAN, OUT_QUAN, ABSQUAN, QUAN, CHARG | Có nhiều trường số lượng khác nhau; cần chụp tiêu đề và dữ liệu để chọn đúng trường | Không chứng minh QUAN hay ABSQUAN là tồn khả dụng. Script không cho công thức quan hệ giữa các trường |
| cfveo.vbs | 16–19 chọn WT dòng 1, METHODS → @M00006; 21 ODP1_CONF_FGR; 22–24 DIF_ALL=true, NISTA=0, EXCCODE=OVP2; 27 save; 29 refresh | Đường xác nhận về 0 có ghi mã exception thực tế OVP2 | Không có kết quả status bar, trạng thái WT/ODO hay tồn sau save; chưa chứng minh STO được nhả |
| can.vbs | 16–21 mở danh sách, gọi chi tiết rồi METHODS → @M00008 | Có một action riêng, theo tên file được dùng cho cancel | Không có selectedRows, popup, trạng thái sau thao tác; không đủ chứng minh phạm vi hủy và kết quả thành công |
| candienmoi.vbs | 16 OIP_CHANGE; 21 sửa QTY_UI dòng 0 thành 180; 24 sendVKey 25; 25 BUTTON_1; 28 batch=261011-540; 31 OK_OI_DLV_CREATE_AND_SAVE | Sửa số lượng ODO trong giao diện EWM rồi đi vào màn hình tạo WT theo batch chỉ định | Không có lệnh cancel WT trong file; không nhập bin hay qty riêng trên màn hình tạo; không chứng minh WT sinh đúng 180 hay nguyên batch |

Các giá trị OB, batch, số 180 và row 0/1 chỉ là dữ liệu mẫu. Không đưa chúng thành giá trị mặc định nghiệp vụ.

## 3. Những chỗ cần sửa trong phân tích Antigravity

1. “Chuẩn xác 100%” chưa có cơ sở: script ghi thao tác, không ghi kết quả nghiệp vụ. Cần phân biệt điều quan sát được và điều phải kiểm chứng live.
2. Đúng hướng khi dùng MON đọc tồn và WT, nhưng không mặc định node N000000139 là Physical Stock hoặc Available Stock. Chọn sai ý nghĩa cột tồn sẽ làm sai gate nguyên batch.
3. Confirm 0 có đường cụ thể OVP2; cancel có đường riêng @M00008. Không dùng hai hành động thay thế tự động cho nhau. Sau mỗi action phải đọc lại WT, ODO và tồn liên quan.
4. Không có bằng chứng “hủy WT → STO tự nhả hạn ngạch”. WT đóng/hủy không tự chứng minh delivery qty đã giảm hoặc ERP/STO đã đồng bộ. Trước VL10D phải đọc lại lượng còn mở và lượng delivery hiện hữu.
5. “VL02N chỉ khi thiếu SLoc” là định hướng nghiệp vụ người dùng cung cấp, chưa được bốn VBS chứng minh. candienmoi cho thấy EWM có đường sửa ODO/tạo WT; nhánh 2001 phải có bộ xử lý riêng.
6. “FIFO vét cạn” và “nguyên batch” phải được định nghĩa cùng nhau: batch cũ chỉ còn 10 sau một lần xuất trước thì có lấy hết 10 được không? Đề xuất lập kế hoạch theo toàn bộ phần còn đủ điều kiện xuất tại thời điểm bắt đầu yêu cầu, tách rõ khỏi số lượng kiện ban đầu. Không tự đưa quy ước này vào code khi chưa chốt.

## 4. Luồng đề xuất

### Nút 1: đọc nguồn và lập kế hoạch

1. Nhận mã, qty, batch tùy chọn; giữ Request ID cho lần chạy tiếp. Xác định ship-to 2001, Owner và kho riêng; không suy ra cả ba đều là 3093 từ ô Owner.
2. Lấy danh sách OB thuộc đúng phạm vi XMĐ và STO/item liên quan. Nếu dùng VL06O làm nguồn OB thì kiểm tra đúng ship-to và trạng thái trước đưa sang MON. Bốn VBS chưa có bước tự tìm OB XMĐ.
3. Nhập toàn bộ OB vào S_ERP_H; chọn các dòng chứng từ tương ứng rồi gọi chi tiết WT một lần. Loại OB trùng, đọc WT theo khóa duy nhất, không cộng lại vì refresh.
4. Tra tồn MON theo các mã có trong kế hoạch; giữ Owner=3093 nếu cấu hình yêu cầu. Không sao chép thao tác xóa Owner/batch trong lượt khám phá của mon.vbs.
5. Chụp snapshot nguồn tồn, WT và ODO; phân loại giữ nguyên, confirm 0, cancel, thay batch/tạo WT, chờ nguồn hoặc chờ sức chứa. Chưa ghi SAP.

### Nút 2: thực hiện theo kế hoạch

1. Đọc lại đúng WT/ODO sắp thay đổi; nếu người khác sửa thì lập lại kế hoạch riêng mã và nhóm batch bị ảnh hưởng.
2. Mã ngoài danh sách: chỉ tác động WT trong OB thuộc phạm vi lượt chạy. Dùng confirm 0 nếu đó là quy tắc nghiệp vụ đã xác nhận; không hủy toàn bộ OB hỗn hợp.
3. Batch sai/thiếu: xác minh đủ tất cả OB trong nhóm trước khi cancel; xác minh tồn được giải phóng trước tạo WT thay thế.
4. Nếu cần sửa lượng ODO, dùng đường EWM đã ghi trong candienmoi; sau sửa kiểm tra lại item, lượng và trạng thái. Chỉ dùng VL02N cho bước phụ được phép.
5. Tạo WT theo batch kế hoạch. Đọc lại bin, batch, qty, trạng thái và item thực tế được sinh; chọn batch không bảo đảm hệ thống tự chọn đúng bin mong muốn.
6. Chỉ khi lượng delivery còn thiếu thực sự và STO còn đủ lượng mới dùng VL10D tạo phần thiếu. Nếu chưa đồng bộ thì chờ và kiểm tra lại nhóm đó, không tạo thêm ngay.
7. Đọc lại cả nhóm nhiều OB; ghi kết quả theo từng hành động và trạng thái đối chiếu. Bốn VBS chưa có đường confirm bình thường cho WT giữ lại hay Post GI, nên chưa gộp chúng vào DONE.

## 5. Kiểm soát nguyên batch và FIFO

Khóa đối chiếu tối thiểu: kho + Owner + Material + Batch + loại tồn + UoM; nếu quy tắc là nguyên rack thì thêm bin/HU tương ứng. Không gộp hai mã có cùng tên batch, hai Owner, hai loại tồn hoặc hai đơn vị khác nhau.

Với từng nhóm, giữ riêng: tồn vật lý, tồn khả dụng, lượng giữ bởi WT thuộc yêu cầu, lượng giữ bởi yêu cầu khác và lượng đã xác nhận trong chính Request ID. Chưa chọn công thức từ QUAN/ABSQUAN/OUT_QUAN khi chưa biết ý nghĩa thực tế. Có thể phải tra thêm WT ngoài danh sách để phân biệt tồn bị người khác giữ.

| Tình huống | Xử lý đề xuất |
|---|---|
| Batch cần 50, A=25 và B=25, cùng phạm vi và trạng thái hợp lệ | Giữ cả hai; đánh giá tổng 50 |
| Batch cần 50, toàn bộ OB chỉ có 40 | Đánh dấu thiếu 10; xác minh nguồn/OB đầy đủ rồi mới xử lý theo quy tắc hủy đã chốt |
| A=25 mở, B=25 đã hủy | Không cộng thành 50 |
| A=25 đã confirm trong lượt này, B=25 còn mở | Đối chiếu phần đã hoàn tất và phần còn phải làm; không confirm A lần nữa |
| Batch cũ còn 10 hợp lệ, batch mới có task 40 | Lập kế hoạch vét 10 cũ; kiểm tra sức chứa ODO và quy tắc phần batch mới trước thay task |
| Tồn khả dụng=0 vì chính WT của nhóm đã giữ 50 | Không kết luận hết tồn hoặc thiếu batch chỉ từ available=0 |
| Batch cũ bị request khác giữ | Không lấy lượng đó làm nguồn có thể thay thế ngay |
| Tổng WT lớn hơn lượng batch mục tiêu | Kiểm tra trùng WT, trùng OB/STO item và lượng người khác giữ; không confirm vượt |

Batch người dùng nhập cần có ý nghĩa rõ: batch bắt buộc hay batch gợi ý. Nếu bắt buộc, không tự đổi sang batch FIFO khác. Thứ tự FIFO cũng cần trường ngày phù hợp hoặc quy ước batch đã xác nhận; tên batch mẫu chưa đủ để chứng minh ngày FIFO.

## 6. Tích hợp với RC47

Workbook có INPUT, POSTING, VL06O_DATA, LOG, CONFIG, BATCH_PLAN, POSTING_RAW, SEARCH_LIST và LEDGER. INPUT hiện lưu ship-to 4007; CONFIG có EXPECTED_UOM=EA, Owner=3093, allowed stock FU và blocked QT/QI/QU. CONFIG E9 ghi rõ “2001 chưa triển khai”. Đây là bằng chứng về cấu hình lưu trong file, không phải kết luận audit toàn bộ VBA.

Đề xuất thêm bộ điều phối riêng cho 2001, dùng lại đầu vào, Request ID, kế hoạch và ledger nếu hợp đồng dữ liệu tương thích. Nguồn MON không ghi đè POSTING/VL06O_DATA bằng dữ liệu khác nghĩa. Có thể thêm MON_STOCK, MON_WT và MON_PLAN khi cần giữ snapshot có thể kiểm tra. Mọi sửa đổi sau này phải giới hạn dispatcher 2001 và không thay thuật toán đang ổn của 3001/4007.

Ledger cần khóa action gồm Request ID + chứng từ/item + TANUM + loại hành động; timeout sau save đặt trạng thái chưa rõ kết quả rồi đọc SAP trước retry. Với tạo WT, cần đối chiếu WT mới theo chứng từ/item/batch và khoảng thời gian thao tác, không chỉ theo TANUM cũ. DONE phải có mức rõ: lập kế hoạch, WT đã tạo, confirm 0 thành công, WT xác nhận bình thường, GI hoàn tất.

## 7. Dữ liệu còn thiếu trước triển khai

| Cần bổ sung | Lý do |
|---|---|
| Tên node N000000010/N000000139 và ảnh tiêu đề grid | Chốt đúng nhánh và ý nghĩa trường tồn |
| Warehouse/monitor, ship-to, ERP OB ↔ DOCNO_H/DOCNO và item | Ngăn xử lý nhầm shop hoặc chứng từ |
| TANUM, trạng thái WT, UoM, stock type, Owner, bin/HU, ngày FIFO | Đọc đúng trạng thái, nguyên batch và FIFO |
| can.vbs ghi đầy đủ chọn một WT → cancel → refresh | Xác minh @M00008, phạm vi và hậu kiểm |
| cfveo ghi status/WT/ODO/tồn trước và sau | Xác minh hiệu lực OVP2, không suy ra hậu quả STO |
| candienmoi từ MON mở đúng ODO/item đến WT mới xuất hiện | Chốt bước điều hướng, qty, batch, bin và popup |
| Trường hợp đủ batch trên nhiều OB, WT đã confirm/hủy và ODO thiếu SLoc | Kiểm chứng các nhánh quyết định thực tế |
| Đường confirm bình thường và GI nếu thuộc phạm vi nút 2 | Bốn file hiện chưa chứng minh luồng hoàn tất xuất |

Không cần ghi lại tất cả thao tác: ưu tiên bổ sung kết quả trước/sau action và các cột định danh còn thiếu. Các click tiêu đề, caretPosition, cửa sổ mở rồi đóng và thao tác khám phá không đưa nguyên xi vào VBA.

## 8. Bộ tình huống nghiệm thu

1. A25+B25 đủ batch50: không cancel nhầm.
2. Chỉ A40/batch50: báo thiếu đúng 10; không tác động trước khi đủ snapshot.
3. OB chứa mã kế hoạch và mã ngoài kế hoạch: chỉ xử lý đúng WT ngoài kế hoạch.
4. Confirm0 OVP2: WT/ODO/tồn được đọc lại; không tự tuyên bố STO đã nhả.
5. Cancel WT rồi tạo lại: không còn task cũ mở và không tạo trùng task mới.
6. Batch cũ còn10, batch mới40: kế hoạch tuân thủ quy tắc đã chốt, không vượt lượng ODO.
7. Timeout ngay sau save/create: chạy tiếp đối chiếu SAP trước retry.
8. Người khác thay WT: chỉ lập lại mã/nhóm bị ảnh hưởng.
9. Tồn QU/QI/QT và FU cùng batch: không cộng tồn bị loại vào nguồn xuất.
10. UoM khác EA hoặc số lượng khác chuẩn: báo riêng, không ép số.
11. ODO thiếu SLoc hoặc chưa có WT: chuyển đúng bước phụ, không xem là hết tồn.
12. Delivery chưa nhả nhưng VL10D còn hiển thị dòng: không sinh OB trùng.

Kết luận triển khai: đã đủ để dựng đặc tả và bộ đọc/preview MON; bộ ghi SAP cần chốt các trường định danh và hậu quả action bằng một số lượt ghi bổ sung. Chưa có cơ sở tuyên bố nhánh 2001 chạy hoàn chỉnh chỉ từ bốn VBS này.
