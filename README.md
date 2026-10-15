# OB TOOL V1 RC63 (ALL IN ONE)

Canonical XLSM: `261010_OB_TOOL_V1_RC63.xlsm`  
Canonical SHA256: `03A9740F8877FCD14324ED05606698E0BEE01508958209FD73CC6DD3253C4E5B`  
Handoff Archive: `RC63_HANDOFF.zip`  
Archive SHA256: `81FE40984D47EB1A470CB8A4F76B785153FC511B425A5D729335AAABF09821F9`

## 📊 Bản phát hành Công Cụ Tồn Kho: Stock FG MA BODY V3.7 (15/10/2026)

Tệp Excel Macro: `261008_Stock_FG_MA_BODY_V3_7.xlsm`  
Kích thước: 807,813 bytes  
SHA-256: `0A5DA3EB764D41BD85FB3235D974BB03F3F39B13CB748452F626BE2FA1AAC05D`  
Gói nén ZIP: `261008_Stock_FG_MA_BODY_V3_7.zip` (SHA-256: `EED33261C44DB42EA1BE7AF01795FFBEC22DCD9456FB54542460ED672D12D463`)

### 🌟 Điểm Cải Tiến V3.7
- **Tích hợp Dữ liệu Trạm Line MA**: Bổ sung đầy đủ cột `Trạm` liên kết trực tiếp với dữ liệu trạm line thực tế.
- **Tích hợp Kế hoạch Sản xuất XMĐ**: Nhập toàn bộ dữ liệu kế hoạch sản xuất từ file KHSX XMĐ mới nhất.
- **Chuẩn hóa Thứ tự Cột Bảng TỒN TỔNG & CẢNH BÁO**:
  `Material` | `Usage` | `Kho Body` | `Line Body` | `Can build Body` | `Intransit` | `Tồn Kho MA` | `Tồn Line MA` | `Model` | `Trạm`
- **Tự động Đồng bộ Báo cáo Tồn Kho Body – MA**: Cập nhật công thức và logic cảnh báo tồn kho, khả năng build theo thời gian thực.

---

## 🌟 Điểm Cải Tiến Đột Phá Trên RC63 (ALL IN ONE)

**Tài liệu tham chiếu VL10D (Ship-to 4007):** [RC63/references/vl10d1(1).vbs](RC63/references/vl10d1(1).vbs) · [Hướng dẫn RC63](RC63/README.md). Đây là bản ghi thao tác SAP GUI, chưa tích hợp vào XLSM phát hành.

### 1. Nút Bấm Điều Phối Toàn Trình "⚡ CHẠY TẤT CẢ (ALL IN ONE)"
- **Single-Click Execution**: Một nút bấm duy nhất trên sheet `INPUT` tự động hoàn thành từ A đến Z mọi bước:
  * Tự động đặt `MODE = "APPLY"`.
  * Tự động điều hướng theo Ship-to (`INPUT!R5`):
    - **Ship-to 4007**: Kiểm tra STO Due -> Tự động quét giải phóng chứng từ nháp 3001 nếu thiếu -> Tạo OB qua VL10D -> Xác minh kết quả OB trên SAP -> Lập kế hoạch phân bổ -> Ghi SLoc/Batch qua VL02N -> Tự động xử lý phục hồi nếu có lỗi.
    - **Ship-to 2001**: Tự động chuyển qua nhánh EWM Shop XMD (`OB2001_Check` -> `OB2001_Execute`).
    - **Ship-to tiêu chuẩn khác**: Kiểm tra và thực thi liên hoàn.
- Bảo toàn trọn vẹn Nút 1 (`OB_Check`) và Nút 2 (`OB_Execute`) cho người dùng muốn kiểm soát từng bước độc lập.

### 2. Khắc Phục Triệt Để Root Cause A: Vòng Lặp Xóa Cam Kết 2 Pha (`SAPX_FindAndReleaseDraft3001OB_Batch`)
- **Chỉ trừ deficit khi xác nhận xóa thành công (`DELETE_VERIFIED`)**: Không trừ trước số lượng thiếu khi mới chỉ chọn ứng viên.
- **Tự động tìm ứng viên thay thế khi bị khóa**: Khi gặp OB bị khóa (`WAIT_STO_LOCK`, `WAIT_OB_LOCK`) hoặc xóa thất bại, tool không trừ số lượng, đánh dấu đã thử và tự động chọn tổ hợp ứng viên tiếp theo trong danh sách cho đến khi giải phóng đủ 100% số lượng cần thiết hoặc hết ứng viên.
- **Khấu trừ dồn ảnh hưởng đa mã tức thì**: Mỗi khi 1 OB được xóa thành công, toàn bộ các mã vật tư trong OB đó được cập nhật giảm deficit ngay lập tức.

### 3. Khắc Phục Triệt Để Root Cause B: Đối Soát VL03N & Không Crash Khi Lỗi WMS (`VL589`)
- **Tra cứu trực tiếp qua VL03N (`SAPX_CheckDeliveryStatus`)**: Xác định chính xác trạng thái chứng từ (`NOT_FOUND`, `LOCKED`, `EXISTS`, `ERROR`).
- **Phục hồi mềm khi chứng từ biến mất (`OB_ABSENT_CONFIRMED`)**: Khi SAP gặp sự cố WMS `VL589: Outbound delivery could not be distributed` và rollback chứng từ, tool nhận diện chứng từ không tồn tại, đánh dấu `mExcludedOB(ob) = True`, xếp các mã linh kiện vào hàng đợi `QueueRecovery` để lập kế hoạch bù mà **không tung `Err.Raise` làm sập tool**.
- **Bảo toàn 100% thành quả**: Các OB đã lưu thành công trước đó (như `1012754414`) được giữ nguyên trọn vẹn.

### 4. Bảo Toàn Tính Toàn Vẹn Nguyên Batch (Whole-Batch Integrity)
- Không bao giờ chia nhỏ Batch (ví dụ Batch 50 PCS không bao giờ bị cắt vụn gán vào item 17 PCS).
- Phân định rõ ràng `createdObQty`, `usableWholeBatchQty` và `savedWholeBatchQty`.
- Nếu dung lượng OB không đủ nguyên Batch, tool chuyển sang trạng thái `WAIT_FULL_BATCH_CAPACITY` và đưa vào kế hoạch tạo bù an toàn.

### 5. Gia Cố Bảo Vệ Chứng Từ Đa Mã Khi Xóa Cấp Dòng (`SAPX_DeleteDraftOB`)
- Xác thực phản hồi từ SAP khi chọn dòng và xóa dòng item (`cannot be selected`, `cannot be deleted`).
- Nếu không tìm thấy dòng cần xóa (`targetRow < 0`): Báo `DELETE_ITEM_NOT_FOUND` và dừng lại an toàn.
- **Tuyệt đối không bao giờ rơi (fall-through) sang xóa cả chứng từ** khi đang yêu cầu xóa dòng lẻ.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC63 (15/10/2026)

### 📦 Các tệp đính kèm RC63

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC63_HANDOFF.zip`** | 922,732 bytes | `81FE40984D47EB1A470CB8A4F76B785153FC511B425A5D729335AAABF09821F9` | **Gói bàn giao đầy đủ RC63 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC63.xlsm`** | 934,543 bytes | `03A9740F8877FCD14324ED05606698E0BEE01508958209FD73CC6DD3253C4E5B` | Tệp Excel Macro phát hành RC63 (Đã tối ưu & Compile) |
| **`VBA_RC62_TO_RC63.diff`** | 22,599 bytes | `FDC2B06BFDF0818E8C87117F66D5000B363FFCAE5589A165092EB5609425A630` | Chi tiết bản vá mã nguồn VBA từ RC62 lên RC63 |
| **`CHANGELOG_RC62_TO_RC63.md`** | 6,237 bytes | `A7297963925B32ABDB687F1012726D86B3924380491D32C465977E822750AE15` | Nhật ký thay đổi và giải pháp kỹ thuật RC63 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC62 (15/10/2026)
* **`RC62_HANDOFF.zip`**: SHA-256 `DB39C4BD0CD87E48DA41213DF0E305BD2B90010CD30BD3295EDAA82681C1C00D`
* **`261010_OB_TOOL_V1_RC62.xlsm`**: SHA-256 `4947174A06C0D0017BD17924803C14699B6305F6A1253C9ADECD66E0A97C6593`

### OB Tool V1 RC61 (15/10/2026)
* **`RC61_HANDOFF.zip`**: SHA-256 `0DB0D4437D316843DC3DE7CFA517DD2ADE8432631B0E98CD78504A13519180B0`
* **`261010_OB_TOOL_V1_RC61.xlsm`**: SHA-256 `E446C065F22FD3D25CD14318AA54CB66C7599C7E6637650DE2C2A974C1DFDF5D`
