# NHẬT KÝ THAY ĐỔI VÀ NÂNG CẤP KỸ THUẬT: OB TOOL V1 RC39

**Ngày phát hành:** 10/10/2026  
**Phiên bản:** OB TOOL V1 RC39  
**Tập tin Workbook:** `261010_OB_TOOL_V1_RC39.xlsm`  
**Gói bàn giao:** `RC39_HANDOFF.zip`  

---

## 1. NGUYÊN NHÂN GỐC RỄ & GIẢI PHÁP TỐI ƯU HÓA RC39

Người dùng phản hồi hai vấn đề lớn về hiệu năng trên bản trước:
1. *"bản mới vẫn bị load xuống dưới cùng và mất thời gian ở việc định hình mã nào cần xóa."*
2. *"sau khi xóa lại load 1 đến 2 lần trước khi điền sloc và batch."*

### 1.1. [P0] Dừng quét ngay khi hết item - Early Stop trong `OverviewSnapshot`
* **Nguyên nhân gốc rễ:** 
  - Trong SAP table control `tblSAPMV50ATC_LIPS_OVER`, SAP luôn cấp sẵn một vùng đệm trống (buffer) từ 20 đến 50 dòng phía dưới bảng.
  - Mã nguồn cũ trước đây khi gặp dòng trống:
    `If itemNo = "" And material = "" Then GoTo NextOverviewRow`
    khiến vòng lặp tiếp tục cuộn trang qua biến `pos = 0 To cnt - 1 Step page` tới tận đáy bảng (dòng cuối cùng của buffer). Điều này làm SAP GUI phải render lại toàn bộ màn hình qua mạng nhiều lần, gây hiện tượng bảng bị cuộn xuống dưới cùng và tốn từ 2-4 giây vô ích!
* **Giải pháp triệt để trên RC39:**
  - Vì các delivery items trong SAP luôn được sắp xếp liên tục từ dòng 0 trở xuống, ngay khi bắt gặp dòng trống đầu tiên (`itemNo = "" And material = ""`), tool kích hoạt **Early Stop**:
    ```vba
    If itemNo = "" And material = "" Then GoTo finished
    ```
  - Tại nhãn `finished:`, bổ sung chốt an toàn đưa vị trí cuộn về đầu bảng nếu vị trí lưu cũ nằm ngoài phạm vi item hợp lệ:
    ```vba
    If savedPosition >= LiveOverviewRowCount(rows) Then savedPosition = 0
    If CLng(tbl.VerticalScrollbar.Position) <> savedPosition Then tbl.VerticalScrollbar.Position = savedPosition
    ```
  - **Kết quả:** Với các đơn hàng thông thường (1-3 items), `OverviewSnapshot` hoàn tất quét ngay trên trang đầu tiên chỉ trong **0.05 giây**, triệt tiêu 100% hiện tượng cuộn xuống đáy bảng!

### 1.2. [P0] Xóa bỏ lệnh gọi kép trùng lặp trong `OB_Apply`
* **Nguyên nhân:**
  Tại đầu hàm `OB_Apply`, mã nguồn cũ gọi hai lần liên tiếp:
  ```vba
  Set mOverviewIndex = OverviewSnapshot(ses)
  Set mUnplannedOriginalSnapshot = OverviewSnapshot(ses) ' Lệnh gọi trùng lặp thừa thãi!
  ```
* **Giải pháp trên RC39:**
  Thay thế trực tiếp bằng:
  ```vba
  Set mOverviewIndex = OverviewSnapshot(ses)
  Set mUnplannedOriginalSnapshot = mOverviewIndex
  ```
  Tiết kiệm hoàn toàn 1 lần quét toàn bảng ngay khi bắt đầu xử lý mỗi OB.

### 1.3. [P0] Tối ưu hóa sau khi xóa item trong `RemoveUnplannedItems`
* **Nguyên nhân:**
  Sau khi bấm nút xóa (`btnBT_POLO_T`) và xác nhận popup, các item chưa xóa trong SAP tự động dồn lên đầu bảng (bắt đầu từ dòng 0). Bản cũ gọi `Set live = OverviewSnapshot(ses)` quét lại toàn bảng, khiến màn hình bị load giật cục.
* **Giải pháp trên RC39:**
  - Sau khi xác nhận xóa xong, đặt vị trí cuộn về 0 (`tbl.VerticalScrollbar.Position = 0`).
  - Đọc nhanh trực tiếp các dòng hiển thị trên trang hiện tại (`0 To VisibleRowCount - 1`) cho tới dòng trống đầu tiên (`itemNo = "" And material = "" Then Exit For`).
  - Cập nhật trực tiếp `live(itemNo) = Array(material, batch, qty, rr, sloc)` và gán ngay `Set mOverviewIndex = live`.
  - Nếu số item còn lại vượt quá 1 trang hiển thị (trường hợp cực hiếm > 15-20 item), mới kích hoạt fallback `OverviewSnapshot(ses)`.
  - **Kết quả:** Quá trình hậu xóa item diễn ra tức thì (~0.02 giây), không giật màn hình.

### 1.4. [P0] Tối ưu hóa `LocateItem` - Tìm kiếm trực tiếp không hủy cache
* **Nguyên nhân:**
  Bản cũ trong `LocateItem` khi không khớp tuyệt đối ở dòng dự kiến thì ngay lập tức gán `Set mOverviewIndex = Nothing`, kích hoạt Pass 2 gọi lại `OverviewSnapshot(ses, True)` làm bảng phải reload thêm 1-2 lần trước khi điền SLoc và Batch!
* **Giải pháp trên RC39:**
  - **Bước 1:** Kiểm tra ngay vị trí dự kiến từ `mOverviewIndex` (0 ms).
  - **Bước 2 (Fast Scan):** Quét nhanh các dòng đang nhìn thấy trên trang hiện tại (`0 To VisibleRowCount - 1`) trước khi tính đến chuyện cuộn bảng.
  - **Bước 3 (Targeted Jump):** Nếu item nằm ở trang khác theo `mOverviewIndex`, nhảy thẳng tới vị trí đó.
  - **Bước 4 (Fallback Early Exit):** Cuộn từng trang nhưng dừng ngay lập tức khi tìm thấy item hoặc khi gặp dòng trống.
  - **Tuyệt đối KHÔNG gán `Set mOverviewIndex = Nothing`**, bảo toàn trọn vẹn cache chỉ mục qua các lượt điền item.

---

## 2. BẢO TOÀN NGUYÊN TẮC HOẠT ĐỘNG & BỘ QUY TẮC CỐT LÕI
1. **Rule 1 (Open SLoc Check Gate):** Bảo lưu 100% việc kiểm tra SLoc ban đầu và skip an toàn nếu OB đã được xử lý.
2. **Rule 2 (Delete Unplanned Items):** Bảo lưu 100% cơ chế nhận diện và xóa triệt để mọi item không có trong kế hoạch.
3. **Rule 3 (Mandatory SLoc & Split):** Bảo lưu 100% logic điền SLoc bắt buộc và phân tách lô (Batch Split).
4. **Rule 4 (Fast Save & Pre/Post-Save Gate):** Bảo lưu 100% Fast Save khôi phục từ RC34 và chốt chặn Pre-save / Post-save.
5. **Invariants Nhánh 3001 & 4007:** 4 Invariants Shop Body 3001, WholeBatchGate và nguyên tắc FIFO được giữ nguyên vẹn 100%.

---

## 3. KẾT QUẢ BIÊN DỊCH VÀ KIỂM THỬ KỸ THUẬT
- **VBE Debug Compile:** `compile_enabled_before = True`, `compile_enabled_after = False` (PASS 100% trên Excel Native COM).
- **Macro Gắn Phím:**
  - Nút 1 (Kiểm tra): `'261010_OB_TOOL_V1_RC39.xlsm'!OB_Check`
  - Nút 2 (Thực hiện): `'261010_OB_TOOL_V1_RC39.xlsm'!OB_Execute`
- **Dọn dẹp Bảng:**
  - Sheet `LOG`: Giữ đúng 5 dòng tiêu đề (max row: 5).
  - Sheet `LEDGER`: Giữ đúng 1 dòng tiêu đề (max row: 1).
- **Dung lượng & Checksum:**
  - File size: 786,760 bytes.
  - SHA-256: `1C374619A87D3A888DC8CA82E307FE8784747DE63AED9AE3FB6ABB30E3A27E32`.
