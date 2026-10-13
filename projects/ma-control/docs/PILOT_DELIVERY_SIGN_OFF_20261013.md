# BÁO CÁO NGHIỆM THU ĐỢT SỬA LỖI & SẴN SÀNG PILOT TỐI 13/10 (AIM-1167)

- **Mã công việc cha**: `AIM-1167`
- **Chỉ huy điều phối & Tổng hợp**: Astra — Giám đốc điều hành (`e5dfdb49-dd75-4a07-9522-4c4ab786e84b`)
- **Thời điểm hoàn thành**: 13/10/2026
- **Phạm vi nghiệm thu được phê duyệt**: **`READY_IMPORTED_SOURCE_PILOT`**
- **Phân định ranh giới kỹ thuật**: **`LIVE_SAP_NOT_VERIFIED`** (Chưa kiểm chứng trên SAP live; pilot tối 13/10 vận hành ở chế độ nhập tệp/đối soát nguồn xuất thật)
- **Cổng dịch vụ Webapp**: `http://127.0.0.1:8012` (Đã tách độc lập, không xung đột với AI Video Factory tại cổng 8000)

---

## 1. Tổng Kết Chuỗi Nhiệm Vụ Pilot Chuyên Trách

Toàn bộ 3 nhiệm vụ chuyên trách độc lập, không chồng lấn được Ban điều hành giao đã hoàn tất 100%:

| Mã Task | Chuyên trách | Vai trò | Trạng thái | Báo cáo Artifact Đầu Ra |
|---|---|---|---|---|
| **AIM-1172** | **Kai** (`7878f735-db2f-4842-a65b-e4c9a2fb2bc4`) | Trưởng kỹ thuật | `done` | `work/artifacts/ma_pilot_kai_baseline.json`<br>- Đối chiếu chi tiết MB51 (Plant 3093, layout `/khoma`), VL06I (Item View btn[18]), khóa 3 trường `OB|Material|Batch`, mẫu BBBG `261008_BBBG_FG_V2_2_14.xlsm`, quy trình manual CF và bảng so sánh Offline vs Live SAP. |
| **AIM-1173** | **Luna** (`e229aed5-811d-4327-acae-9ac496f7ca87`) | Trưởng kiểm toán chất lượng | `done` | `work/artifacts/ma_pilot_luna_review.json`<br>- Phán quyết: **`APPROVED`** cho phạm vi `READY_IMPORTED_SOURCE_PILOT`.<br>- 26/26 mã băm SHA-256 mã nguồn và baseline template khớp 100%.<br>- 98/98 bài kiểm thử PASS trên CSDL tạm.<br>- Xác nhận 6/6 trọng điểm: sửa triệt để CF0, snapshot tồn chống cộng dồn, cô lập tab Body/EDC, loại bỏ OB giả, chặn GOS false success và chống lặp/idempotent. |
| **AIM-1174** | **Nova** (`3d3b8059-60e8-44dd-a664-bbaf294b0bc0`) | Chuyên trách kiểm thử độc lập | `done` | `work/artifacts/ma_pilot_nova_qa.json`<br>- Phán quyết: **`PASS_OFFLINE`** (và ghi nhận minh bạch `LIVE_NOT_VERIFIED`).<br>- 98/98 tests PASS trong 16.90s trên CSDL tạm (`MA_CONTROL_DB_PATH`).<br>- CSDL vận hành hoàn toàn nguyên vẹn (SHA-256 trước và sau test không đổi).<br>- Vượt qua toàn diện 8 ca thử nghiệm bắt buộc (TC-01 đến TC-08). |

---

## 2. Kết Quả Khắc Phục Các Lỗ Hổng Nêu Tại REVIEW_20261013.md

1. **Khắc phục triệt để lỗ hổng CF = 0**:
   - Hàm `verify_cf()` bắt buộc `actual_quantities` phải là dict đầy đủ ID từng dòng pick, kiểm tra hữu hạn và `> 0`. Từ chối lập tức khi CF = 0 hoặc thiếu dòng (HTTP 400).
   - Tuyệt đối không tự động gán `cf_actual_qty = req_qty` khi dữ liệu không hợp lệ; giữ nguyên trạng thái `Chờ người dùng CF`.
   - Chặn cứng tạo BBBG trước khi đơn pick được người dùng xác nhận CF hợp lệ.
2. **Minh bạch hóa trạng thái kết nối SAP**:
   - Loại bỏ nhãn tĩnh "SAP đã kết nối" trên HTML.
   - `/api/health` phản ánh trung thực trạng thái `DISCONNECTED` / `MANUAL` kèm `live_write_verified: false` khi không có phiên SAP live đang hoạt động.
3. **Loại bỏ hoàn toàn OB giả lập & demo seed**:
   - Ngoài chế độ `TEST`, cấm tự sinh OB từ SQLite ID; yêu cầu nhập số OB thực tế đã đối chiếu SAP.
   - Xóa draft 3001 ngoài chế độ test trả về `verified: false` và `LIVE_WRITE_NOT_IMPLEMENTED`, không báo thành công ảo.
4. **Chặn cờ thành công giả lập cho BBBG & GOS**:
   - Khi tạo BBBG, hệ thống luôn đặt cờ khởi tạo `unc_export_status='PENDING'`, `gos_bbbg_attached=0`, `gos_vat_attached=0`.
   - Xuất file `.xlsx` thật từ template `BBBG_NEW` của baseline `261008_BBBG_FG_V2_2_14.xlsm`, lưu mã băm SHA-256 vào `job_events`, hỗ trợ tải qua `/api/bbbg/{id}/download`.
   - Tự động dịch chuyển dòng chân trang và vùng in khi danh sách hàng dài hơn 18 dòng.
5. **Tách cổng dịch vụ**:
   - Webapp MA Control cấu hình chạy tại cổng **`8012`**, giải quyết triệt để xung đột với AI Video Factory tại cổng `8000`.

---

## 3. Quy Chế & Hướng Dẫn Vận Hành Pilot Tối 13/10

### Chế độ Được Hỗ Trợ & Đã Kiểm Chứng (Supported & Verified):
1. Nạp snapshot tồn MB52 (phân tách độc lập 3093, 3000, 3001; không cộng dồn lịch sử).
2. Ghi nhận số OB thực tế đã đối soát SAP (cô lập rành mạch nhánh Body 3001 và EDC 4007).
3. Tạo đơn pick và kiểm soát cổng CF thủ công bắt buộc từng dòng (`actual_quantities > 0`).
4. Nạp tệp xuất thực tế MB51 + VL06I để đối soát 3 trường (`OB|Material|Batch`), đối xứng movement và đơn vị tính.
5. Xuất file BBBG `.xlsx` cục bộ chuẩn template V2.2.14 có banner phân định rõ trạng thái bản nháp vs đã đối chiếu.

### Chế độ Chưa Kiểm Chứng (Chưa Thực Thi Trên SAP Live):
- Không tự động tạo/xóa OB trực tiếp trên SAP qua GUI Scripting.
- Không tự động thực hiện Post Goods Receipt (GR).
- Không tự động đính kèm GOS, không tự động in ra máy in vật lý và không tự động gửi email Outlook.

### Lệnh Khởi Chạy Webapp Pilot:
```powershell
# Chạy Webapp Pilot trên cổng 8012
python -m uvicorn app.main:app --host 127.0.0.1 --port 8012 --app-dir G:\AI-Manager\projects\ma-control
```
Truy cập: `http://127.0.0.1:8012`

---

## 4. Kết Luận Nghiệm Thu (Final Disposition)

- **Astra phê chuẩn đóng công việc cha AIM-1167 với trạng thái `done`**.
- Toàn bộ các phát hiện kỹ thuật đã được khắc phục, kiểm toán độc lập (`AIM-1173: APPROVED`) và kiểm thử tự động 98/98 test pass (`AIM-1174: PASS_OFFLINE`).
- Hệ thống sẵn sàng 100% cho đợt pilot thực tế tối 13/10 trên dữ liệu nguồn xuất thật theo đúng phạm vi `READY_IMPORTED_SOURCE_PILOT`.
