# RC63 — OB TOOL ALL IN ONE

Bản phát hành chính: [OB TOOL V1 RC63](../261010_OB_TOOL_V1_RC63.xlsm) và [RC63_HANDOFF.zip](../RC63_HANDOFF.zip). Chi tiết thay đổi: [CHANGELOG_RC62_TO_RC63.md](../CHANGELOG_RC62_TO_RC63.md).

## Tài liệu tham chiếu VL10D

- [vl10d1(1).vbs](./references/vl10d1(1).vbs) — bản ghi SAP GUI Scripting do người dùng cung cấp cho luồng VL10D, Ship-to/Receiving Plant 4007.
- Mục đích: đối chiếu thao tác lựa chọn STO/Material/Plant, tạo delivery và các bước trên bảng kết quả VL10D khi rà soát hoặc nâng cấp RC63.
- Giữ nguyên logic lệnh và thứ tự thao tác từ bản ghi; chuyển mã hóa từ UTF-16 LE (Windows recording) sang UTF-8 để dễ xem/diff trên GitHub.
- **Reference only**: script sử dụng các ô tọa độ SAP GUI cố định (`chk[1,5]`, `txt[63,5]`) và giá trị mẫu `100`, chưa có điều kiện xác minh, giới hạn quyền thao tác hay phục hồi. Không tích hợp/chạy nguyên mẫu như code production.
- Các nguyên tắc nghiệp vụ RC63 vẫn theo CHANGELOG: với Ship-to 4007, kiểm tra STO/Due, chống tạo OB trùng, tính lại giải phóng 3001 dựa trên xóa đã xác nhận, giữ nguyên Batch và kiểm tra kết quả VL02N.

Không thay đổi file XLSM, ZIP phát hành hay SHA-256 đã công bố của RC63 khi bổ sung tài liệu tham chiếu này.
