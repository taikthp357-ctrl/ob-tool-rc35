# Bổ sung RC49: ánh xạ OB ERP với Document EWM

Nguồn: `ob1.vbs`, bản ghi SAP GUI người dùng cung cấp ngày 13/10/2026. Giữ nguyên byte và mã hóa của file gốc.

Luồng ghi lại: mở node `N000000010`, nạp danh sách OB qua multiple selection `S_ERP_H`, chọn tất cả chứng từ ở bảng trên rồi gọi toolbar `N000000011`. Bảng dưới có các cột `DOCNO_H`, `ERP_NUMBER`, `PRODUCTNO`, `PO_NUMBER` và `UOM`.

Dùng bảng này để đọc quan hệ OB ERP (`ERP_NUMBER`) với Document EWM (`DOCNO_H`), giữ thông tin material/reference ở từng dòng và kiểm chứng trên dữ liệu live trước triển khai. Đây là đường bổ sung khác với nút `N000000096` dùng gọi WT trong các VBS trước.

Không gán document dòng đầu tiên cho mọi OB. Đọc ánh xạ cả danh sách khi đầy đủ; chỉ tra riêng OB chưa rõ. Bản ghi không chứa dữ liệu clipboard, không đọc giá trị kết quả và chưa cung cấp TANUM/trạng thái WT. Bước nối Document/Item với WT vẫn cần định danh và đối chiếu riêng trước confirm/cancel.

File này là tư liệu bổ sung cho RC49; chưa tích hợp vào workbook hoặc thay đổi tag/release hiện có.
