# NHẬT KÝ THAY ĐỔI VÀ NÂNG CẤP KỸ THUẬT: OB TOOL V1 RC42

**Ngày phát hành:** 10/10/2026  
**Phiên bản:** OB TOOL V1 RC42  
**Tập tin Workbook:** `261010_OB_TOOL_V1_RC42.xlsm`  
**Gói bàn giao:** `RC42_HANDOFF.zip`  

---

## 1. NGUYÊN NHÂN GỐC RỄ & GIẢI PHÁP TỐI ƯU HÓA TRÊN RC42

Dựa trên dữ liệu LOG và hình ảnh thực tế từ phiên chạy 10:54 PM của RC41:

### 1.1. [P0 - Nhánh 4007] Mở khóa bước tạo Delivery trên VL10D ở chế độ APPLY
* **Hiện tượng:**
  - Chế độ ghi SAP đã được thiết lập là `APPLY`.
  - Tool đã tìm thấy STO `3601040480|610` với số lượng Due khả dụng là 280 pcs và lập kế hoạch sẵn sàng:
    `VL10D_PLAN_READY: BIW70004090AA; STO=3601040480/610; CREATE=280; need=680; OB capacity=0; due=280`
  - Tuy nhiên, ngay trước khi bấm nút tạo Delivery, tool dừng lại với lỗi:
    `VL10D_GROUP_WAIT: STO=3601040480; GROUP_PREPARE; -2147217243; Group due missing Material/STO/Item; no Create.; no automatic Create retry`
* **Nguyên nhân gốc rễ:**
  - Trong `ModuleSAP.EDC4007_PrepareGroup` (dòng 4974), câu lệnh kiểm tra:
    `If CStr(a(0)) = sto And val(CStr(a(1))) = val(CStr(planned(2))) And CStr(a(3)) = "4007" And CStr(a(10)) = ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093") Then`
  - Trên màn hình danh sách Due của VL10D dạng Classic View (như hiển thị trong ảnh chụp SAP của người dùng), không có cột hiển thị Supplying Plant (`RESWK`). Do đó, phần tử `a(10)` mang giá trị rỗng `""`.
  - Phép so sánh `"" = "3093"` trả về `False`, khiến biến `targets.count` bằng 0 thay vì 1. Khi đó, điều kiện `targets.count <> plans.count` đã văng lỗi 4261 trước khi gửi lệnh tạo Delivery sang SAP.
* **Giải pháp trên RC42:**
  - Trong `EDC_ReadClassicDue`: Gán mặc định `a(10) = ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093")` vì STO đã được ME2N xác thực Supplying Plant = 3093 từ bước kiểm tra trước đó.
  - Trong `EDC4007_PrepareGroup`: Nới lỏng điều kiện kiểm tra cột nguồn thành `(CStr(a(10)) = "" Or CStr(a(10)) = ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093"))`.
  - Trong `EDCPressCreate`: Bổ sung kiểm tra tooltip hỗ trợ các nhãn tiếng Anh/Đức (`"background"`, `"hintergrund"`, `"deliver"`, `"liefer"`), đảm bảo nút Background Delivery (`btn[19]`) luôn được kích hoạt trơn tru.

---

### 1.2. [P0 - UI & Trình bày] Chuẩn hóa cột ngày tháng và bố cục sheet LOG
* **Khắc phục lỗi ngày tháng hiển thị số thực thô (`46305,95369`):**
  - Trong các bản trước, lệnh `ws.Cells(r, 1).Resize(1, 6).Value2 = Array(Now, ...)` ghi trực tiếp giá trị `Now` kiểu Double vào ô định dạng General, khiến Excel hiển thị dạng số thực `46305,95369`.
  - Trên RC42: Trong `ModuleLog.WriteLog`, định dạng ô `ws.Cells(r, 1).NumberFormat = "m/d/yy h:mm AM/PM"` và gán `.Value = Now`. Toàn bộ cột A sheet LOG được thiết lập mặc định định dạng `[$-409]m/d/yy h:mm AM/PM;@`. Mọi dòng log mới ghi vào luôn hiển thị đúng định dạng ngày giờ (ví dụ: `10/10/26 10:54 PM`).
* **Căn chỉnh độ rộng cột và căn lề (Triệt tiêu hiện tượng tràn chữ / đè cột):**
  - **Cột A (`TIMESTAMP`)**: Độ rộng **20**, căn giữa (Center).
  - **Cột B (`RUN ID`)**: Mở rộng từ 14 lên **20**, căn giữa (Center) — hiển thị trọn vẹn chuỗi `20261010_225359` không bị tràn hay đè chữ sang cột C.
  - **Cột C (`STEP`)**: Mở rộng từ 18 lên **30**, căn trái (Left) — hiển thị trọn vẹn các bước dài như `WAIT_FULL_BATCH_CAPACITY`, `POST_SEARCH_ROUTE`.
  - **Cột D (`OB / REF`)**: Độ rộng **18**, căn giữa (Center) — hiển thị thoáng đãng mã STO `3601040480|610`.
  - **Cột E (`MATERIAL`)**: Độ rộng **20**, căn giữa (Center).
  - **Cột F (`MESSAGE`)**: Độ rộng **90**, căn trái (Left).
  - **Hàng tiêu đề 5**: Chiều cao **24pt**, căn giữa dọc.

---

### 1.3. [P1] Đồng bộ hóa cấu hình sheet CONFIG và LEDGER
* Cập nhật ô `CONFIG!B2` thành `OB_V1_RC42` (khắc phục nhãn cũ `OB_V1_RC34`).
* Tối ưu độ rộng cột sheet `LEDGER`: Cột G (`UPDATED AT`) = 20, Cột I (`MESSAGE`) = 35.

---

## 2. KẾT QUẢ KIỂM THỬ VÀ BIÊN DỊCH VBE

- **VBE Debug Compile:** PASS 100% (Control `Compi&le VBAProject` chuyển từ `Enabled = True` sang `Enabled = False`, không có lỗi cú pháp).
- **Sheet LOG & LEDGER:** Dọn sạch dữ liệu cũ (LOG row = 5, LEDGER row = 1), bố cục định dạng chuyên nghiệp.
- **Button Bindings:** Đã liên kết chuẩn xác với file `261010_OB_TOOL_V1_RC42.xlsm`.
- **Bảo toàn 4 Invariants:**
  1. Bảo toàn item ngoài kế hoạch (Non-target item preservation) $\rightarrow$ Giữ nguyên.
  2. WholeBatchGate (Không tách dở dang 1 batch) $\rightarrow$ Giữ nguyên.
  3. FIFO Allocation $\rightarrow$ Giữ nguyên.
  4. Fast Save Mode & Direct Item Locate $\rightarrow$ Hoạt động ổn định.
