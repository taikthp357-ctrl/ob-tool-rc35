# NHẬT KÝ THAY ĐỔI VÀ NÂNG CẤP KỸ THUẬT: OB TOOL V1 RC44

**Ngày phát hành:** 11/10/2026  
**Phiên bản:** OB TOOL V1 RC44  
**Tập tin Workbook:** `261010_OB_TOOL_V1_RC44.xlsm`  
**Gói bàn giao:** `RC44_HANDOFF.zip`  

---

## 1. NGUYÊN NHÂN GỐC RỄ & GIẢI PHÁP TỐI ƯU HÓA TRÊN RC44

Dựa trên dữ liệu LOG và phản ánh thực tế từ phiên chạy 11:54 PM - 11:56 PM:

### 1.1. [P0 - Nhánh 4007] Mở khóa lọc danh sách nhiều vật tư trên VL10D (`EDCSetGroupMaterialFilter`)
* **Hiện tượng:**
  - Khi sheet INPUT có nhiều hơn 1 vật tư (ví dụ danh sách 6 mã), tool mở cửa sổ popup Multiple Selection trên VL10D (`wnd[1]`), paste danh sách 6 mã từ clipboard nhưng dừng lại và không bấm F8 để đóng popup.
  - Xuất hiện chuỗi lỗi lặp lại liên tiếp: `-2147217243; Cannot verify pasted Material selection table.` cho tất cả các mã vật tư trong danh sách.
* **Nguyên nhân gốc rễ:**
  - Sau khi dán danh sách (`btn[24]`), hàm cố gắng kiểm tra bảng dữ liệu qua `SAPX_FindByTokens(ses, "tblSAPLALDBSINGLE", False)`.
  - Hàm `SAPX_FindByTokens` chỉ quét trong phạm vi cửa sổ chính `wnd[0]`, không quét trong cửa sổ popup `wnd[1]`, do đó luôn trả về `Nothing` và văng lỗi 4261 ngay trước khi dòng lệnh bấm F8 (`btn[8]`) được chạy. Popup không được đóng lại và màn hình chính không được điền Plant `4007`.
* **Giải pháp trên RC44:**
  - Chuẩn hóa theo mô hình lọc ổn định của `POST_ApplyBatchFilter` và `POSTING`: Bấm nút Paste (`btn[24]`), sau đó bấm ngay nút OK / F8 (`btn[8]`) để áp dụng và đóng popup `wnd[1]`.
  - Loại bỏ hoàn toàn khối lệnh quét table control thừa thãi, đảm bảo tool đóng popup trơn tru và tiếp tục điền Receiving Plant `4007`, Supplying Plant `3093` và chạy F8 trên VL10D.

---

### 1.2. [P0 - Nhánh 4007] Bảo toàn số lượng Due đọc được trên Classic View VL10D (`EDC_ClassicDue`)
* **Hiện tượng:**
  - Trên màn hình VL10D của SAP, STO của mã `BIW70004127` thực tế còn Due là **399 pcs**.
  - Tuy nhiên tool lại ghi log: `VL10D_DUE_CANDIDATES: eligible rows=0` và báo lỗi `WAIT_STO_DUE_4007: need=1261; OB capacity=0; missing=1261`.
* **Nguyên nhân gốc rễ:**
  - Khi mã vật tư ngắn hơn (11 ký tự thay vì 14 ký tự), bảng dạng văn bản của Classic View bị xô lệch toạ độ (`Open=48; Uom=68`).
  - Cột `Open` bị dịch sang vị trí 48 nơi không có dữ liệu trên dòng hàng, khiến `openQty = CleanDouble("") = 0`.
  - Câu lệnh `If openQty < qty Then qty = openQty` đã vô tình ép `qty = Min(399, 0) = 0`, làm điều kiện `If qty > 0 Then` bị sai và loại bỏ hoàn toàn dòng Due 399 pcs.
* **Giải pháp trên RC44:**
  - Thêm điều kiện bảo vệ: `If openQty > 0 And openQty < qty Then qty = openQty`. Khi cột Open trống rỗng hoặc bằng 0, không bao giờ triệt tiêu số lượng Due khả dụng.
  - Bổ sung cơ chế quét lân cận $\pm 6$ cột nếu ô nhập số lượng `txt[xQty, rowNo]` bị xê dịch toạ độ do độ dài chuỗi văn bản.

---

### 1.3. [P0 - Thuật toán phân bổ] Tích hợp bộ tính `WholeBatch-fit CreateQty` (WholeBatchGate-aware)
* **Bài toán thực tế:**
  - Nhu cầu mã `BIW70004127` là 1261 pcs với 7 lô FIFO:
    `261001-210` (161 pcs), `261001-150` (200 pcs), `261009-210` (100 pcs), `261009-209` (200 pcs)...
  - STO Due trên VL10D hiện chỉ còn **399 pcs**.
  - Nếu tạo OB với số lượng thô 399 pcs: Tool sẽ gán Lô 1 (161) + Lô 2 (200) = 361 pcs. Còn lại 38 pcs không thể ghép vào Lô 3 (100 pcs) vì vi phạm nguyên tắc WholeBatchGate (không xé lẻ lô), dẫn tới việc OB 399 pcs bị hụt 38 pcs không có batch trên VL02N.
* **Giải pháp trên RC44:**
  - Bổ sung hàm tính toán [`EDCWholeBatchFitQty`](file:///G:/AI-Manager/projects/ob-tool/rc44-work/vba/ModuleMain.bas): Quét các lô FIFO khả dụng trong kho và tính tổng số lượng của các lô nguyên lớn nhất sao cho $\sum \text{BatchQty} \le \text{STO Due}$.
  - Cụ thể: Lô 1 (161) + Lô 2 (200) = **361 pcs** $\le 399$ pcs (Lô 3 cần 100 pcs sẽ vượt quá 399 nên dừng).
  - Đặt số lượng tạo OB: $\text{CreateQty} = \mathbf{361}\text{ pcs}$.
  - **Kết quả:** Nút 1 tạo OB đúng 361 pcs. Sang Nút 2, VL02N chia vừa khít 2 lô nguyên (161 + 200 = 361), không còn số lượng lẻ 38 pcs, bảo toàn 100% WholeBatchGate và lưu SAP thành công tuyệt đối!

---

## 2. KẾT QUẢ NGHIỆM THU KỸ THUẬT & BIÊN DỊCH VBE

- **Native VBE Debug Compile:** **PASS 100%** (`compile_enabled_before = True` $\rightarrow$ `compile_enabled_after = False`).
- **Sheet LOG & LEDGER:** Đã dọn sạch dữ liệu cũ (LOG row = 5, LEDGER row = 1), giữ chuẩn định dạng hiển thị ngày giờ và độ rộng cột.
- **Button Bindings:** Đã liên kết chuẩn xác với file `261010_OB_TOOL_V1_RC44.xlsm`.
- **Bảo toàn 4 Bất biến hệ thống:** Non-target item preservation, WholeBatchGate, FIFO Allocation, Fast Save Mode.
