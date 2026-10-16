# MA CONTROL — Điều Hành FG (Python + Excel-like UI)

Phiên bản: **v1.0.0-rc52-FG** (13/10/2026)  
Được phát triển nhằm hợp nhất và thay thế bộ 3 công cụ Excel:
1. **Stock FG/MA/BODY V3.6**: Quản lý tồn kho, xác định mã ưu tiên xuất và kéo line.
2. **OB Tool V1 RC52**: Nghiệp vụ SAP cốt lõi xử lý OB (Body 3001, EDC 4007 STO, EWM 2001).
3. **BBBG FG V2.2.14**: Tạo biên bản bàn giao, in ấn, lưu tổng/shop, đính kèm chứng từ BBBG/VAT.

---

## 🚀 Tính Năng Cốt Lõi Đã Nghiệm Thu (100% Hoàn Thành)

1. **Quy trình chuẩn hóa 5 bước nghiêm ngặt**:
   `[1. Tải tồn SAP]` ➔ `[2. Điền OB]` ➔ `[3. Ra đơn pick]` ➔ `[4. Người dùng CF]` ➔ `[5. BBBG]`
   * **Cổng CF thủ công**: Đơn sau khi sinh bắt buộc giữ ở trạng thái **"Chờ CF"**, tuyệt đối không tự động sinh biên bản ngầm cho đến khi người dùng kiểm tra thực tế và xác nhận trên giao diện.
2. **Giao diện thao tác kiểu Excel (Excel-like Grid)**:
   * Hỗ trợ copy/paste nhiều dòng trực tiếp từ Excel (Ctrl+V) vào bảng OB và bảng Chờ CF.
   * Giữ nguyên số `0` ở đầu mã vật tư (không bị cắt xén số liệu).
   * Phím điều hướng Tab, Enter, chỉnh sửa ô tại chỗ.
3. **Tải tồn SAP & Đối soát thông minh**:
   * Hỗ trợ nạp tồn tự động từ SAP (MB52) và đối soát trực tiếp với MB51/VL06I.
   * Tự động bù trừ đảo hàng (Movement Reversals: 641/642, 643/644, 601/602, 351/352).
4. **Ma trận tồn 8 Model x 9 Tier & Cảnh báo tồn thấp**:
   * Phân bố tồn kho trực quan 8 Model x 9 Tier.
   * Cảnh báo đỏ nổi bật khi số lượng xe `<= 200 XE`.
   * Tính toán chỉ số DOH (Days On Hand) theo dữ liệu thực.
5. **Tách biệt 2 nhánh nghiệp vụ**:
   * Tab **Body • 3001**
   * Tab **EDC • 4007**
6. **Xuất Biên Bản Bàn Giao (BBBG)**:
   * Xuất file Excel định dạng chuẩn mẫu `BBBG FG V2.2.14`.
   * Hỗ trợ lưu trữ, đối soát lịch sử xuất đơn.

---

## 🛠️ Hướng Dẫn Khởi Chạy

### Cách 1: Chạy trực tiếp bằng lệnh (Quick Start)
```powershell
# Cài đặt thư viện phụ thuộc
pip install -r requirements.txt

# Khởi động ứng dụng (mặc định tại http://127.0.0.1:8012)
python -m uvicorn app.main:app --host 127.0.0.1 --port 8012 --reload
```

### Cách 2: Sử dụng Script Windows Launcher
* Nhấp đúp vào `START_MA_CONTROL.cmd` để tự động khởi động server và mở trình duyệt.
* Chạy `RUN_TESTS.cmd` để thực thi toàn bộ 141 bài kiểm thử tự động.

---

## 🧪 Kết Quả Kiểm Thử (QA Acceptance)
* **Tổng số ca kiểm thử**: 141 tests
* **Kết quả**: **141 / 141 PASSED (100%)**
* **Các bộ test chính**:
  * `tests/test_schema_and_stock.py`: Schema 10 bảng SQLite, WAL mode, bảo toàn số 0.
  * `tests/test_flow_invariants.py`: Bất biến quy trình 5 bước, bất biến cổng CF thủ công.
  * `tests/test_source_reconciliation.py`: Đối soát MB51/VL06I, loại trừ net-zero và giao dịch đảo.
  * `tests/test_bbbg_export.py`: Xuất Excel BBBG và chữ ký số.
  * `tests/test_ui_and_api_e2e.py`: Kiểm thử E2E giao diện Excel Grid và API.
  * `tests/test_user_scope_low_stock_and_direct_cf.py`: Paste nhiều dòng Chờ CF, ma trận tồn và cảnh báo `<= 200 XE`.
