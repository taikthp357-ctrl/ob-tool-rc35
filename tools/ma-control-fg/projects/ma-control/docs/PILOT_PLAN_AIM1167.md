### 🏛️ Báo Cáo Tiếp Nhận Chỉ Đạo & Kế Hoạch Giám Sát Pilot Tối 13/10 (Astra)

Astra xác nhận tiếp nhận chỉ đạo từ Ban điều hành (comment `46dc43a5-8aa3-4377-a3fc-9889183ac500`) và chính thức mở lại việc giám sát dự án **MA Control — Webapp Điều hành FG**:

---

#### 1. Ghi Nhận Đầy Đủ Các Lỗ Hổng Từ Hồ Sơ Kiểm Tra Độc Lập (`docs/REVIEW_20261013.md`)
1. **Lỗ hổng CF thực tế = 0 vẫn vượt qua**: Cần kiểm soát nghiêm ngặt `actual_quantities`, không cho phép CF=0 hoặc sai/thiếu được tính là xác nhận thành công, chặn cứng việc sinh BBBG từ các dòng chưa đủ điều kiện.
2. **Minh bạch trạng thái kết nối SAP**: Phân định rõ ràng giữa MOCK/Chưa kết nối và phiên SAP thật; không gán nhãn tĩnh "SAP đã kết nối" khi chưa có phiên live.
3. **Xử lý OB thực chất**: Không suy diễn kết quả xóa draft 3001 hay sinh số OB ảo từ SQLite ID; phải tích hợp kết quả SAP thật hoặc đánh dấu rõ trạng thái mô phỏng.
4. **Bằng chứng xuất BBBG & đính kèm GOS**: Không tự động gán cờ `unc_export_status='SAVED'` hay `gos_bbbg_attached=1` khi chưa có file thật và kết quả thao tác thực tế.
5. **Tách cổng dịch vụ**: Không dùng cổng `8000` (đang phục vụ AI Video Factory V4.5); tách sang cổng độc lập (ví dụ `8001` hoặc `8050`).

---

#### 2. Kế Hoạch Theo Dõi Chuỗi Nhiệm Vụ Pilot
Tuân thủ nghiêm ngặt chỉ đạo **không sửa mã nguồn trực tiếp** (lượt sửa mã do phiên người dùng hiện tại thực hiện), Astra kích hoạt theo dõi tiến độ của 3 nhiệm vụ chuyên trách:
- 💻 **AIM-1172 (Kai)**: Đối chiếu baseline BBBG V2.2.14 / RC52 và kế hoạch pilot, kiểm tra độc lập nguồn MB51/VL06I và quy trình manual CF; xuất artifact `work/artifacts/ma_pilot_kai_baseline.json`.
- 🔍 **AIM-1173 (Luna)**: Giám sát chờ artifact `work/artifacts/ma_pilot_ready_for_review.json` xuất hiện, sau đó rà soát độc lập các ca CF0, snapshot tồn trùng, branch isolation, OB giả và GOS false success; xuất artifact `work/artifacts/ma_pilot_luna_review.json`.
- 🧪 **AIM-1174 (Nova)**: Giám sát chờ readiness artifact để kiểm thử độc lập trên CSDL tạm (`MA_CONTROL_DB_PATH`); kiểm thử các ca bắt buộc (CF0, thiếu/sai, MB51/VL06I, file BBBG thật); xuất artifact `work/artifacts/ma_pilot_nova_qa.json` và phân định rạch ròi giữa `PASS_OFFLINE` và `LIVE_NOT_VERIFIED`.

---

#### 3. Điều Phối Trạng Thái Issue Cha (AIM-1167)
- Chuyển `AIM-1167` sang trạng thái **`blocked`** phụ thuộc vào 3 subtask: `[AIM-1172, AIM-1173, AIM-1174]`.
- Tuyệt đối không dùng trạng thái `done` cũ hoặc kết quả test mock làm bằng chứng nghiệm thu SAP live. Astra sẽ tổng hợp và đánh giá lại khi các artifact pilot được bàn giao đầy đủ.
