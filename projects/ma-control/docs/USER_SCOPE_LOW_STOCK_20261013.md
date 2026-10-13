# Phạm vi nghiệp vụ người dùng chốt — tồn thấp và OB (13/10/2026)
Chỉ dẫn này thay thế phần suy luận trước về việc tự tạo pick và tự đưa OB vào chờ CF.

1. CHỜ CF DO NGƯỜI DÙNG QUẢN LÝ
- Cho dán trực tiếp danh sách từ Excel vào Chờ CF, không bắt buộc phải đi qua đơn pick do ứng dụng tự sinh. Hỗ trợ OB/mã hàng/plant/batch/số lượng và thông tin packing box/rack khi có; giữ số 0 đầu.
- Packing box, rack, ghép/tách đơn và quyết định ra đơn do người xử lý. Ứng dụng không tự tính/chọn cách đóng gói hoặc tự điều phối việc ra đơn.
- Chờ CF là danh sách người dùng đưa vào theo quy trình của họ; không tự nhận OB/SAP đã CF và không tự suy diễn trạng thái từ số OB. Nhập/lưu/lọc/chỉnh sửa và chống dán trùng phải hoạt động thật.

2. CHỨC NĂNG CƠ BẢN ƯU TIÊN
- Tải snapshot tồn, phân biệt mã hàng + Plant + SLoc + Batch + stock type, không cộng lịch sử hoặc trộn shop/zone. Quản lý cảnh báo tồn thấp theo ngưỡng cấu hình và nguồn đầy đủ/còn mới.
- Điền OB thông thường giữ luồng hiện có, không tự bật chiến lược tồn thấp.

3. CHẾ ĐỘ TÙY CHỌN ĐIỀN OB THEO TỒN THẤP
- Người dùng phải chủ động chọn chế độ này; chọn mã hàng, plant và phạm vi kho/shop cần kiểm tra. Hiển thị tồn hiện tại, ngưỡng, số lượng yêu cầu và dữ liệu nguồn.
- Kiểm tra theo mã hàng + plant, lấy chi tiết từng batch/SLoc thay vì gộp mù. Đề xuất/chọn 2–3 batch có tồn khả dụng để cho đi trước theo nhu cầu; cho người dùng xem và điều chỉnh số lượng mỗi batch trước khi thực hiện.
- Không lấy toàn bộ tồn của batch chỉ vì được ưu tiên. Tổng lượng không vượt nhu cầu/tồn khả dụng; loại hoặc ghi rõ batch khóa/chưa xác minh. Chống dùng trùng lượng đã được phân bổ trong cùng lượt.
- Tránh ảnh hưởng cấp hàng shop khác: dùng đúng phạm vi nguồn, mức dự trữ và lượng đã phân bổ có bằng chứng. Nếu chưa có thông tin nhu cầu/dự trữ shop khác thì ghi rõ chưa đánh giá được, không tự cam kết tránh dừng chuyền hoặc âm thầm lấy nguồn khác.
- Phân biệt đề xuất/lập kế hoạch với tạo OB thật trên SAP; chỉ nhận thành công khi có kết quả thật, còn chế độ nhập nguồn/manual phải ghi đúng là ghi nhận/đề xuất.

4. ASTRA TỰ CHIA VIỆC VÀ KIỂM TRA
- Kai triển khai nhập Excel Chờ CF + quản lý tồn thấp + chế độ OB tùy chọn; không tự động hóa packing/rack.
- Luna đối chiếu phạm vi mới và baseline; Nova kiểm thử dán 0 đầu/trùng, tắt-bật chế độ, nhiều plant/shop, 2–3 batch, thiếu tồn, nguồn cũ, giới hạn lượng và bấm lặp.
- Giữ các bảo vệ BBBG đã có, không mở rộng tự động hóa ngoài phạm vi người dùng vừa chốt. Báo cáo tối đa5 dòng; không làm lại phần đã qua kiểm thử nếu không đổi mã.
