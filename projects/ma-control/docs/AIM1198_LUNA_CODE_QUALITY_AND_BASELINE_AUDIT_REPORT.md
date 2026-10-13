# BÁO CÁO KIỂM TOÁN CHẤT LƯỢNG MÃ NGUỒN & ĐỐI CHIẾU BASELINE MA CONTROL FG (AIM-1198)

- **Mã công việc**: `AIM-1198`
- **Tên công việc**: `[Review] MA Control - Tách trạng thái OB, chuẩn hóa Chờ CF & hoàn thiện điều hướng API thực`
- **Người thực hiện**: Luna — Trưởng kiểm toán chất lượng (`e229aed5-811d-4327-acae-9ac496f7ca87`)
- **Đơn vị phụ trách**: Paperclip Company Control Plane
- **Thời điểm hoàn thành**: 13/10/2026
- **Kết luận kiểm toán**: **CHẤP THUẬN NGHIỆM THU (APPROVED / PASS 100%)**

---

## 1. Tổng Quan & Căn Cứ Kiểm Toán

Căn cứ chỉ đạo từ Hội đồng Quản trị và Giám đốc điều hành Astra:
1. Rà soát chuẩn tắc 8 trạng thái OB và gate kiểm soát tại hàm `create_pick_order`.
2. Kiểm toán quy tắc nghiệp vụ Chờ CF (tách biệt trạng thái, luồng CF đủ tự rời danh sách, CF thiếu giữ số lượng và lý do, định tuyến nhóm "Cần xử lý").
3. Kiểm toán tính trung thực của nhãn hiển thị khi chưa có SAP live, ngăn chặn suy đoán hủy đơn từ MB51 rỗng và không đoán đuôi OB.
4. Kiểm toán trang Chờ CF, 5 mục điều hướng webapp (chạy 100% API thật, không nút giả/alert ảo), và tính bất biến khi chuyển tab/refresh/restart (idempotency).
5. Rà soát an toàn repository GitHub `MA-Logistic-Control` (nhánh `codex/ma-control-fg-portable`, PR #1) và gói đóng gói Portable Windows x64 ZIP.

---

## 2. Kết Quả Kiểm Toán Chi Tiết Theo 5 Mục Tiêu Trọng Điểm

### 2.1. Rà soát 8 trạng thái OB & Gate kiểm soát

- **Quy chuẩn 8 trạng thái OB trong CSDL SQLite và API (`app/models/schemas.py`, `app/api/routes_workflow.py`):**
  Hệ thống chuẩn hóa thống nhất toàn bộ trạng thái OB qua hàm `normalize_ob_status()`, ánh xạ thành 8 mã trạng thái chuẩn tắc (canonical status codes) thuộc 5 nhóm nghiệp vụ rõ ràng:
  1. `UNVERIFIED` ("Chưa kiểm tra") — Nhóm **Chưa kiểm tra** (`is_ready_for_pick: False`)
  2. `INVALID_CANCELLED` ("Không hợp lệ / Đã hủy") — Nhóm **Cần xử lý** (`is_action_needed: True`)
  3. `LOCKED_SHORTAGE` ("Đang khóa / Thiếu tồn") — Nhóm **Cần xử lý** (`is_action_needed: True`)
  4. `READY_FOR_PICK` ("Đủ điều kiện tạo pick" / "Sẵn sàng điền") — Nhóm **Đủ điều kiện** (`is_ready_for_pick: True`)
  5. `WAITING_CF` ("Chờ CF" / "Chờ pick / CF") — Nhóm **Chờ CF** (`is_waiting_cf: True`)
  6. `PARTIAL_CF` ("CF một phần") — Nhóm **Chờ CF** (`is_waiting_cf: True`)
  7. `CF_COMPLETED` ("Đã CF đầy đủ" / "Đã CF") — Nhóm **Đã hoàn thành** (`is_cf_completed: True`)
  8. `READY_FOR_BBBG` ("Đủ điều kiện BBBG" / "Đã lập BBBG") — Nhóm **Đã hoàn thành** (`is_cf_completed: True`)

- **Kiểm toán Gate kiểm soát trong `PickCFService.create_pick_order()` (`app/services/pick_cf_service.py`):**
  - **Chặn nhận OB chưa kiểm chứng:** Khi người dùng truyền dòng nhập tay chưa kiểm chứng (`unverified: True` hoặc trạng thái thuộc `UNVERIFIED`, `LOCKED_SHORTAGE`, `INVALID_CANCELLED`), hàm chặn đứng với ngoại lệ:
    `"Chặn nhận OB nhập tay chưa có bằng chứng kiểm chứng... chỉ cho phép tạo pick cho dòng OB đã kiểm tra tồn và có trạng thái hợp lệ/READY"`.
  - **Chặn linked row chưa READY:** Với các dòng đã liên kết trong CSDL (`ob_items`), hệ thống đối chiếu nghiêm ngặt: nếu `status_code` không thuộc `READY_FOR_PICK` / `WAITING_CF` và `rc52_code` không phải `RC52_READY` / `RC52_WAIT_PICK`, hàm ném `ValueError` và giữ nguyên trạng thái cũ, tuyệt đối không cho phép tạo pick hoặc nhảy cóc trạng thái.
  - **Chống trùng lặp đơn pick:** Nếu dòng OB đã có `pick_order_id` hoặc đã được xác nhận hoàn tất, hệ thống từ chối tạo lại đơn pick.

### 2.2. Kiểm toán quy tắc Chờ CF (Waiting CF Rules)

- **Tách biệt độc lập giữa trạng thái OB và Đơn Pick:**
  - Bảng `pick_orders` duy trì trường `status` riêng biệt (`Chờ người dùng CF`, `PARTIAL_CF`, `Đã CF`).
  - Bảng `pick_order_items` quản lý tiến độ từng dòng bằng `cf_status` (`PENDING_CF`, `PARTIAL_CF`, `CONFIRMED`), `cf_actual_qty`, `missing_qty`, `shortage_reason`.
  - Bảng `ob_items` phản ánh trạng thái tổng thể của dòng nhu cầu nguồn.
- **Quy tắc rời danh sách Chờ CF khi đủ 100%:**
  - Khi thực hiện kiểm đếm thực tế qua API `POST /api/waiting-cf/save-line`, chỉ khi tất cả các dòng trong đơn pick đạt `cf_actual_qty == req_qty` (`all_done = True`):
    - Đơn pick cập nhật `status = 'Đã CF'`.
    - Cờ `bbbg_unlocked = True`, `left_waiting_cf = True`.
    - API truy vấn `GET /api/waiting-cf` (với mặc định `include_completed=False`) tự động lọc bỏ các đơn đã CF đủ 100%, giúp danh sách Chờ CF tự động thu gọn sạch sẽ.
- **Xử lý thiếu / sai số lượng (Partial CF):**
  - Khi `actual_qty < req_qty`, hệ thống tính toán chính xác `missing_qty = req_qty - actual_qty`.
  - Ràng buộc bắt buộc: Bắt buộc người dùng cung cấp lý do thiếu (`shortage_reason`), nếu để trống hệ thống trả mã lỗi `400 Bad Request`.
  - Đơn pick được đánh dấu `PARTIAL_CF`, chốt chặn BBBG giữ nguyên khóa (`bbbg_unlocked = False`).
- **Định tuyến vào nhóm "Cần xử lý":**
  - Mọi dòng OB có trạng thái `INVALID_CANCELLED` hoặc `LOCKED_SHORTAGE` đều mang thuộc tính `status_group = "Cần xử lý"` và `is_action_needed = True`, giúp giao diện và bộ lọc tập trung cảnh báo cho điều phối viên kho.

### 2.3. Kiểm toán nhãn & hiển thị trung thực

- **Nhãn hiển thị trung thực khi chưa có kết nối SAP live:**
  - Webapp không giả mạo trạng thái SAP connected; badge hiển thị minh bạch `● SAP: Chưa kết nối` (`badge-danger`).
  - Toàn bộ dòng Chờ CF hiển thị nhãn nguồn: `"Xác nhận thực tế thủ công/nguồn nhập (Không báo SAP đã CF)"`.
  - Banner quy định nêu rõ: *"Đây là Xác nhận thực tế thủ công / nguồn nhập (kiểm đếm hàng thực tế). Chức năng CF trực tiếp trên SAP live: Chưa hỗ trợ trong bản pilot (hệ thống tuyệt đối không tự báo SAP đã CF)"*.
- **Chống suy đoán dữ liệu (Zero Speculation):**
  - `OBService.replan_deleted_ob()`: Nếu dữ liệu MB51 rỗng (`mb51_rows == []`) và chưa có xác nhận hủy từ SAP (`not verified_deleted`), hàm chặn đứng với lỗi: `"Không suy đoán xóa/hủy OB chỉ từ MB51 rỗng; cần xác nhận chứng từ đã hủy trên SAP"`.
  - Không suy đoán đuôi số OB: Ở chế độ chạy thực tế, ứng dụng từ chối tự sinh số OB giả mạo và yêu cầu nhập số OB đã đối soát với SAP.

### 2.4. Kiểm toán trang Chờ CF & 5 mục điều hướng

- **Trang Chờ CF:**
  - Hỗ trợ dán trực tiếp từ Excel (`POST /api/waiting-cf/paste-excel`), bảo toàn số 0 ở đầu các chuỗi (Mã hàng, Batch, OB, Plant, Box, Rack).
  - Tích hợp kiểm tra chống dán trùng (Anti-Duplicate Check) trực tiếp trên SQLite.
  - Cho phép người dùng trực tiếp sửa dòng (`PUT /api/waiting-cf/lines/{id}`) và xóa dòng dở dang (`DELETE /api/waiting-cf/lines/{id}`).
- **5 mục điều hướng chạy 100% API thật:**
  1. **Tồn & ưu tiên (`view-stock`):** Gọi `/api/health`, `/api/stock/visual-matrix`, `/api/stock/two-tables`, `/api/stock/low-stock/suggestions`.
  2. **Chờ CF (`view-waiting_cf`):** Gọi `/api/waiting-cf`, `/api/waiting-cf/history`, `/api/waiting-cf/save-line`, `/api/waiting-cf/bulk-save`.
  3. **BBBG (`view-bbbg`):** Gọi `/api/bbbg/records`, `/api/bbbg/{id}/download`.
  4. **Lịch sử (`view-history`):** Gọi `/api/jobs/events?limit=100`.
  5. **Cấu hình (`view-config`):** Gọi `/api/settings`, `/api/model-mappings`.
  - *Kết luận:* Không tồn tại bất kỳ nút giả hoặc popup `alert()` đánh lừa người dùng. Toàn bộ hành động đều tương tác với backend FastAPI và lưu vết CSDL SQLite.
- **Tính bất biến & Idempotency:**
  - Khi đổi tab qua lại, refresh F5 hoặc restart service, toàn bộ tiến độ kiểm đếm từng dòng (số lượng thực tế, số lượng thiếu, lý do, người kiểm, timestamp) đều được bảo toàn nguyên vẹn.
  - Cơ chế Idempotent Save: Khi gửi yêu cầu lưu trùng lặp dữ liệu đã lưu, backend nhận biết và trả về `idempotent: True` mà không ghi đè log rác hay làm hỏng dữ liệu.

### 2.5. Kiểm toán Repository GitHub & Đóng gói Portable

- **Rà soát kho chứa GitHub (`https://github.com/taikthp357-ctrl/MA-Logistic-Control`):**
  - Nhánh phân phối: `codex/ma-control-fg-portable`.
  - Pull Request đã tạo: `PR #1` (`feat(fg): Tich hop MA Control FG module & Portable Windows x64 support`).
  - Cấu trúc thư mục: Toàn bộ mô-đun FG được tổ chức riêng biệt trong thư mục `ma-control-fg/`, bảo toàn toàn vẹn ứng dụng gốc của kho lưu trữ.
- **Rà soát an toàn thông tin & bí mật (Security & Sanitization):**
  - Không chứa mật khẩu, API key, token cá nhân hoặc thông tin người dùng thật.
  - Tệp cấu hình nhạy cảm, file database SQLite và file log đều được loại trừ nghiêm ngặt qua `.gitignore`.
- **Rà soát & Hiệu chỉnh các launcher Windows Batch:**
  - Đã rà soát và chuẩn hóa định dạng dòng CRLF (`\r\n`) cho các file launcher `START_MA_CONTROL.cmd`, `STOP_MA_CONTROL.cmd`, `RUN_TESTS.cmd`, khắc phục triệt để hiện tượng lỗi font hoặc lệnh cmd `le / ho` không nhận diện được trên Windows.
- **Kiểm thử độc lập Portable:**
  - Kiểm thử trong thư mục giải nén độc lập có khoảng trắng `MA Control Extracted Test Folder` đạt **129 / 129 test case PASS (100%)** với Exit Code 0 bằng Python runtime nhúng.

---

## 3. Ma Trận Nghiệm Thu Kỹ Thuật (Audit Verification Matrix)

| Tiêu Chí Kiểm Toán | Trạng Thái | File Mã Nguồn & Bằng Chứng | Đánh Giá Của Kiểm Toán Viên |
|---|:---:|---|---|
| **8 Trạng thái OB chuẩn tắc** | ✅ PASS | `app/models/schemas.py`<br>`app/api/routes_workflow.py` | 8 mã chuẩn tắc ánh xạ vào 5 nhóm; cờ phân loại chính xác. |
| **Gate `create_pick_order`** | ✅ PASS | `app/services/pick_cf_service.py`<br>`tests/test_pilot_runtime.py` | Chặn đứng OB chưa kiểm chứng và row chưa READY. |
| **Tách biệt Pick & OB** | ✅ PASS | `app/services/pick_cf_service.py`<br>`tests/test_cf_validation.py` | Trạng thái đơn pick độc lập hoàn toàn với trạng thái OB. |
| **Luồng CF đủ tự rời danh sách** | ✅ PASS | `pick_cf_service.py:list_waiting_cf_lines`<br>`tests/test_pilot_runtime.py` | Đơn 100% CF tự rời danh sách Chờ CF và mở khóa BBBG. |
| **CF thiếu giữ lượng & lý do** | ✅ PASS | `pick_cf_service.py:save_line_cf`<br>`tests/test_pilot_runtime.py` | Bắt buộc nhập lý do nếu thiếu; giữ trạng thái PARTIAL_CF. |
| **Nhãn trung thực khi chưa có SAP** | ✅ PASS | `app/templates/index.html`<br>`app/static/js/app.js` | Hiển thị rõ 'Xác nhận thủ công / Nguồn nhập'; không giả mạo SAP live. |
| **Không suy đoán xóa/hủy OB** | ✅ PASS | `app/services/ob_service.py:replan_deleted_ob` | Ném lỗi nếu MB51 rỗng mà chưa có xác nhận hủy từ SAP. |
| **Không tự đoán đuôi số OB** | ✅ PASS | `app/services/ob_service.py:fill_ob_items` | Bắt buộc số OB thật trong chế độ thông thường. |
| **5 Mục điều hướng chạy API thật** | ✅ PASS | `app/static/js/app.js`<br>`app/templates/index.html` | Tồn & ưu tiên, Chờ CF, BBBG, Lịch sử, Cấu hình chạy 100% API thật. |
| **Tính bất biến & Idempotency** | ✅ PASS | `app/services/pick_cf_service.py`<br>`tests/test_flow_invariants.py` | Bảo toàn dữ liệu qua tab switch, refresh, restart; chống trùng lặp. |
| **Bảo mật Repo & Portable Package** | ✅ PASS | `MA_CONTROL_PORTABLE_WINDOWS_X64_v1.0.0.zip`<br>PR #1 GitHub | Sạch bí mật, loại trừ DB thật, 129/129 tests PASS offline. |

---

## 4. Kết Luận & Bàn Giao Nghiệm Thu

Hệ thống **MA Control FG** đã vượt qua toàn diện 5 mục tiêu kiểm toán chất lượng mã nguồn và đối chiếu baseline theo đúng chỉ đạo của Hội đồng Quản trị:
- **Bộ mã nguồn dự án**: Sạch sẽ, không chứa dữ liệu nhạy cảm, tích hợp đúng nhánh và tạo Pull Request minh bạch trên GitHub.
- **Gói đóng gói Portable**: Chạy hoàn toàn độc lập offline với runtime Python nhúng, đã kiểm thử thành công trên môi trường Windows x64.
- **Tiêu chuẩn chất lượng**: **129 / 129 test case PASS (100%)**.

Trưởng kiểm toán chất lượng Luna chính thức phê duyệt nghiệm thu công việc **AIM-1198**.
