# MA CONTROL — Python & Webapp Local, FG First
Ngày lập: 13/10/2026. Trạng thái: kế hoạch triển khai; chưa phải ứng dụng chạy SAP đã nghiệm thu.

## 1. Mục tiêu và phạm vi
Tiếp nối MA Logistic Control cũ, xây một hệ thống local trên Windows. Triển khai nhánh FG trước dựa trên gói 261008_BBBG_FG_V2_2_14.zip; sau khi FG ổn định mới mở rộng MA, tải tồn tự động, điều phối line và Teams listener.
Đợt này lập dự án và lưu baseline. Chưa chuyển toàn bộ VBA sang Python, chưa chạy thao tác SAP live.

## 2. Baseline
ZIP gồm BBBG V2.2.14, Stock FG/MA/BODY V3.6, MIGO ATT V3.4.3.
SHA256 ZIP: 2b8a3a9d86947c53699d46e524c54e94af8a654df3549012cba47d26f25c9edb.
Đã kiểm tra CRC ZIP và container của 3 XLSM. Đây chỉ là kiểm tra tính toàn vẹn gói; chưa chứng minh macro đúng hoặc SAP chạy thành công.
Trước khi port cần trích VBA, bảng cấu hình, template, selector và log từ từng workbook; lập bảng hàm → nghiệp vụ → bằng chứng → ca kiểm thử. Không mặc định mọi bản nâng cấp đều là golden source.

## 3. Kiến trúc đề xuất
- Python backend + FastAPI; UI HTML/CSS/JavaScript phục vụ trực tiếp tại localhost. Chốt phiên bản dependency ở bước dựng môi trường.
- SQLite lưu dữ liệu nghiệp vụ và trạng thái job, thay cho việc lấy workbook đang mở làm cơ sở dữ liệu.
- Một Windows worker xử lý SAP GUI Scripting qua COM/pywin32 trong một phiên tương tác đã đăng nhập. SAP và Outlook cần desktop session phù hợp; webserver chạy 24/7 không đảm bảo SAP worker chạy được khi Windows logout.
- Mỗi SAP session chỉ một job đang thao tác; mọi yêu cầu từ UI đi vào hàng đợi. COM được khởi tạo và sử dụng trên đúng worker thread.
- Adapter SAP, Outlook, Excel export, printer và UNC tách khỏi domain. Excel có thể giữ làm template/export trong giai đoạn chuyển tiếp.
- Job chạy nền, UI hiển thị tiến độ, lỗi, bước đang chạy, dừng và tiếp tục. Không chạy COM trong request HTTP.
- Bind localhost mặc định. Nếu mở LAN để máy khác điều khiển phải bổ sung đăng nhập, phân quyền và cấu hình mạng trước.

## 4. Luồng FG đầu tiên
Nhập/paste Doc, số phiếu, OB và dữ liệu hàng → chuẩn hóa/kiểm tra → đọc MB51/VL06I theo nhu cầu → đối soát → tạo BBBG → in/gửi mail/lưu tổng và shop → attachment BBBG/VAT → tổng kết.
Mỗi thao tác có nút riêng và Run All dùng cùng pipeline; tách Check và tạo BBBG. Cho phép cấu hình thứ tự Run All theo nghiệp vụ đã xác minh.
Không đẩy nghiệp vụ mới hoặc luồng queue chưa được duyệt vào MVP.

## 5. Quy tắc cần bảo toàn và xác minh
- Không chèn dữ liệu vào giữa phiếu cũ. SQLite giữ bản ghi chuẩn; export tổng theo khóa phiếu/document đã xác định, tìm vị trí nối thêm thật, kiểm tra trùng trước ghi.
- Không dùng số phiếu đơn độc làm khóa duy nhất khi chưa xác minh phạm vi đánh số. Khóa kỹ thuật UUID; khóa chống trùng gồm loại chứng từ, năm, plant/shop, số phiếu/doc phù hợp từng nguồn.
- OB đã cancel về 0 và không có lịch sử MB51: phân biệt không có dữ liệu với lỗi đọc SAP. Chỉ cho tiếp tục theo điều kiện nghiệp vụ của V2.2.14 đã xác minh, ghi rõ lý do; không báo thành công giả.
- Model lấy bảng mapping cấu hình, không suy diễn đè lên mapping người dùng. Kiểm tra BIW71001001 và BIW70002022 theo cấu hình đã cung cấp.
- Người nhận mail theo shop; trống/sai dùng cấu hình theo quy tắc đã xác minh. Người gửi, người bàn giao/ký lấy đúng nguồn; người nhận BBBG được để trống.
- Hàng ngoài hệ thống: mã/batch/số lượng bắt buộc, OB nếu có; bỏ check SAP cho nhóm đó.
- Template BBBG Auto/Set, giữ font/cỡ chữ người dùng, trang in đủ chân trang; chia trang khi quá số dòng.
- GOS: có Display Attachment mới kiểm tra danh sách; chưa có thì attach. Trùng tên VAT bỏ qua theo logic nguồn. Chỉ đánh dấu đủ BBBG+VAT khi có bằng chứng; thông báo SAP thành công có thể đủ theo quy tắc hiện hành.
- Tồn tách theo Plant + SLoc + Material + Batch + Stock type + đơn vị; không cộng trùng các vùng 3093/3000/3001. Cảnh báo kho, line MA và Body dùng đúng ngữ nghĩa từng nguồn.

## 6. Dữ liệu
Bảng dự kiến: materials, model_mapping, shops, recipients, settings, document_headers, document_lines, stock_snapshots, stock_rows, jobs, job_steps, attachments, exports, audit_events.
Lưu số lượng Decimal, mã/batch/doc dạng text giữ số 0 đầu. Mỗi snapshot có thời điểm và nguồn; dữ liệu cũ phải hiển thị tuổi dữ liệu.
Settings có version; job lưu snapshot cấu hình để không bị đổi giữa chừng.
UNC là đích export, không đặt SQLite chính trên share mạng.

## 7. Chống lặp và phục hồi
Một khóa job nghiệp vụ chống bấm Run All trùng. Mỗi bước lưu trạng thái pending/running/succeeded/failed/uncertain/skipped cùng bằng chứng và thời lượng.
Crash sau gửi mail/in/ghi SAP nhưng trước ghi trạng thái phải chuyển uncertain: đối chiếu trước chạy lại, không tự gửi/in/attach lần hai.
Lưu file local bằng temp + atomic replace; export UNC có khóa phối hợp, kiểm tra version và retry giới hạn. Khi UNC lỗi, giữ bản chờ đồng bộ và báo rõ chưa lưu mạng.
Dừng sau bước an toàn; không ngắt COM giữa thao tác lưu. Resume chỉ chạy phần còn thiếu, không chạy lại toàn bộ các mã đã hoàn tất.
Log gồm job ID, khóa nghiệp vụ, bước, timeout, SAP status và bằng chứng; không chứa mật khẩu.

## 8. Màn hình
1. Tổng quan FG: kết nối SAP/Outlook, job đang chạy, hàng chờ, lỗi cần xử lý.
2. BBBG: paste bảng, Doc/OB/số phiếu, shop/model, Check, tạo BBBG, in, mail, lưu, Run All.
3. Tồn FG/MA/BODY: snapshot, Plant/SLoc, lọc model, ngưỡng, cảnh báo.
4. Attachment: BBBG/VAT theo document, trạng thái và lỗi.
5. Cấu hình: mapping, shop/người nhận, template Auto/Set, font, thư mục, máy in, timeout.
6. Lịch sử job: xem từng bước, bằng chứng, tiếp tục job lỗi.

## 9. Lộ trình và đầu ra
P0 — Audit baseline: trích mã/cấu hình/template, sơ đồ luồng, xác định tính năng thật của V2.2.14 và các test nguồn.
P1 — Khung local: backend, SQLite migration, UI nhập/paste, cấu hình, queue, log; adapter giả lập được ghi rõ.
P2 — FG offline: nhập nguồn đã xuất, mapping, đối soát, BBBG, export; so với workbook bằng các bộ dữ liệu chuẩn.
P3 — SAP read-only: kết nối/session, MB51/VL06I/MB52, selector và timeout; đo trên máy Windows thật.
P4 — Tác vụ ngoài: in, Outlook, lưu UNC, GOS; kiểm chứng chống lặp và phục hồi riêng từng adapter.
P5 — Run All FG: nối pipeline, chạy pilot song song Excel, chỉ chuyển vận hành khi đạt nghiệm thu.
P6 — MA mở rộng: tải tồn đầu ca/chu kỳ, DO/STO, cảnh báo, cuối ca MB51; Teams DOM allowlist theo ràng buộc công ty, không dùng Graph/API/webhook khi bị cấm.
Không đặt lịch tự động ở đợt lập kế hoạch này.

## 10. Ca nghiệm thu bắt buộc
- Hai lượt lưu cùng phiếu, hai người ghi tổng cùng lúc, phiếu có khoảng trống, file tổng đang mở: không chèn/ghi đè phiếu khác.
- OB cancel 0 và MB51 rỗng; SAP đọc lỗi; VL06I GR=C; OB chưa có IB: kết quả phân biệt rõ và đúng rule.
- Mapping LIMO/VF7; batch giống nhau trên nhiều OB; hai cột GM nguồn; hàng ngoài hệ thống.
- BBBG ít/nhiều dòng, Auto/Set, font 14 và font tự đặt; in nhiều trang không mất chân.
- VAT trùng, chỉ BBBG, chỉ VAT, đủ hai file, không có Display Attachment.
- Outlook gửi xong rồi worker crash; UNC ngắt mạng; SAP popup lạ; timeout; dừng/resume; bấm Run All hai lần.
- Tổng tồn theo Plant/SLoc bằng nguồn đã kiểm chứng, không cộng trùng.
- PASS offline không được gọi là PASS SAP live. Nghiệm thu live lưu log và artifact đúng hash.

## 11. Hiệu năng và phát hành
Đo thời gian từng bước, p50/p95 theo số document và số dòng; mục tiêu Run All khoảng 30 giây là mục tiêu cần đo, không cam kết trước.
Gom truy vấn SAP theo danh sách khi được hỗ trợ; dùng snapshot trong một job; không quét/mở lại SAP từng dòng vô ích.
Release có ZIP, hash, dependency lock, hướng dẫn cài Windows, backup/migration và bằng chứng test. Không phát hành READY khi dữ liệu/đầu ra rỗng hoặc chỉ chạy mock.
