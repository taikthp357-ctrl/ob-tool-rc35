# Test pilot ma-control tối 13/10/2026

Bản này dành cho pilot với dữ liệu xuất thật từ SAP. Đã đối chiếu baseline RC52 và BBBG V2.2.14; chưa nghiệm thu các thao tác ghi SAP, GR, GOS, Outlook, in máy in hoặc UNC tự động.

## Mở ứng dụng

Trên máy hiện tại: `http://127.0.0.1:8012`. Database pilot riêng: `data/ma_control_pilot_evening.db`. Database cũ và lịch sử không bị xóa. Backup SQLite tại `backups/pilot-20261013-132534`.

Nếu mang sang máy Windows khác, giải nén gói pilot, cài Python 3.12 nếu máy chưa có, rồi chạy PowerShell trong thư mục:

```powershell
powershell -ExecutionPolicy Bypass -File .\run-pilot.ps1 -Install
```

Các lần sau bỏ `-Install`. Cửa sổ chạy server cần giữ mở. Cổng 8000 của Video Factory không bị dùng lại. `127.0.0.1` phải mở trên chính máy đang chạy ứng dụng.

## Trình tự thử bằng một đơn nhỏ

1. Xuất MB52 từ SAP theo Plant/SLoc cần dùng, tải file tại “Snapshot tồn MB52”. Snapshot mới chỉ thay thế nguồn đang dùng của đúng scope; lịch sử được giữ nhưng không cộng vào tồn mới. Body mặc định kiểm tra nguồn 3001/1003; EDC mặc định 4007/4004.
2. Dán các dòng mã hàng/model/số lượng/batch/shop/STO/OB. Giữ mã/batch/OB ở dạng text để không mất số 0 đầu. Số OB phải là số đã kiểm tra trên SAP; ứng dụng không tự tạo/xóa OB. Bấm “Kiểm tra” để dùng tồn server, rồi ghi nhận OB và tạo pick. Không dùng nút để thay cho thao tác tạo OB thật trên SAP.
3. Kiểm tra hàng thực tế. Mở “Kiểm tra CF”, nhập người kiểm tra và số lượng thực tế cho từng dòng. Không có điền sẵn. Số lượng 0/thiếu/sai sẽ bị chặn; không được đổi số yêu cầu chỉ để ép khớp.
4. Xuất MB51 và VL06I đầy đủ cho chính tất cả OB của đơn pick. Chọn đúng đơn ở phần nguồn, nhập thời điểm xuất thật và xác nhận nguồn đủ. Tải cả hai file rồi bấm “Đối soát hai nguồn”. Báo cáo quá 24 giờ hoặc thiếu nguồn không được dùng để xác nhận phiếu.
5. Nhập người giao, người nhận và vị trí thật rồi xuất BBBG. Thiếu đối soát chỉ có bản nháp. Bản đã đối soát giữ quantity/unit/IB/Doc từ nguồn; tên file, hash nguồn và hash mẫu được lưu để truy vết. Nếu đã có bản nháp, xuất lại sau đối soát sẽ nâng cùng phiếu, không tạo trùng.
6. Tải file, đối chiếu với phiếu Excel cũ cho cùng bộ dữ liệu: mã hàng/batch/OB/IB/Doc, đơn vị, lượng ròng, shop, tổng, người ký, layout và trang in. Chỉ bàn giao sau khi người phụ trách chấp nhận kết quả đối chiếu. GOS/GR/in/email vẫn xử lý theo quy trình đã kiểm chứng bên ngoài ứng dụng.

## Cột nguồn chấp nhận

File CSV/TSV/TXT hoặc XLSX/XLSM xuất dữ liệu riêng, không kèm dòng tổng và không chạy macro. Có thể dùng tên cột kỹ thuật hoặc tên tiếng Anh tương đương:

| Nguồn | Cột bắt buộc |
|---|---|
| MB52 | MATNR, CHARG, Quantity/Unrestricted, MEINS, WERKS, LGORT |
| MB51 | XBLNR, MATNR, CHARG, MENGE, MEINS, BWART, MBLNR |
| VL06I | LIFEX, MATNR, CHARG, LFIMG, VRKME, VBELN, WERKS, LGORT, WBSTK/WBSTA |

MB51 chỉ dùng movement xuất 641/643/645/647/601/351 và các mã đảo tương ứng 642/644/646/648/602/352. 101/102 không được trộn vào luồng xuất này. Đối soát theo OB–mã hàng–batch, đơn vị phải giống nhau; lệch số lượng, khóa thiếu, trạng thái GR không hợp lệ hoặc nhiều shop sẽ bị chặn. Trạng thái GR C chỉ là thông tin nguồn, không khiến ứng dụng tự GR lại.

Nếu CSV có số dạng `1,000` mơ hồ, xuất lại dạng số trong XLSX hoặc dùng số thập phân rõ ràng; không tự suy đoán dấu phân nhóm. Nếu không nhận được header, kiểm tra tên cột trong bảng trên và dùng file riêng cho từng báo cáo.

## Dừng khi có lỗi

Giữ nguyên nguồn và file kết quả, ghi số pick/phiếu và thông báo lỗi. Không xóa database, không ép CF, không đổi file nguồn của phiếu đã xuất. Nguồn của phiếu chính thức được khóa; nếu dữ liệu cần sửa, đối chiếu với người phụ trách trước khi lập quy trình điều chỉnh riêng. Không bấm lặp thao tác SAP để bù cho lỗi mạng của ứng dụng.

## Bằng chứng và giới hạn

Bộ kiểm thử: 98 ca qua trên SQLite tạm; gồm CF âm/0/sai, snapshots cũ, tab/batch isolation, gửi lại đồng thời, đối soát nguồn và draft→verified cùng ID. UI được mở kiểm tra và không có lỗi console. File BBBG đã render kiểm tra bố cục theo mẫu, giữ logo và chữ ký.

Đây là bằng chứng kỹ thuật/offline; đối chiếu với dữ liệu SAP thật tối nay mới xác nhận mức phù hợp nghiệp vụ. Ba việc hệ thống theo dõi là AIM-1172 (Kai), AIM-1173 (Luna), AIM-1174 (Nova); nguồn mã/baseline được khóa hash ở `work/artifacts/ma_pilot_ready_for_review.json`.
