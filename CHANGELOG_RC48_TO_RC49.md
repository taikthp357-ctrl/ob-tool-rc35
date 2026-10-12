# CHANGELOG - OB TOOL V1 RC49

## 1. Tổng quan Bản phát hành RC49 (Branch 2001 Shop XMĐ - EWM Monitor /SCWM/MON)
- **Mục tiêu**: Bổ sung bộ điều phối chuyên biệt cho **Nhánh 2001 (Shop XMĐ)** vận hành trực tiếp trên **`/SCWM/MON` (EWM)** khi ô `INPUT!R5 = "2001"`.
- **Bảo toàn Invariants**: Giữ nguyên vẹn 100% logic, thuật toán và luồng chạy của Nhánh 3001 và Nhánh 4007 khi `R5 <> "2001"`.
- **Dựa trên thực nghiệm**: Tích hợp chính xác các đường điều khiển từ 4 bản ghi SAP GUI Scripting VBS:
  * `mon.vbs`: Điều hướng Tree `/SCWM/MON`, ODO Node `N000000010`, Warehouse Tasks `N000000096`, Stock Overview `N000000139`.
  * `cfveo.vbs`: Thao tác Confirm WT về 0 với mã ngoại lệ `OVP2` (`METHODS` -> `@M00006`, `ODP1_CONF_FGR`, `DIF_ALL = true`, `NISTA = 0`, `EXCCODE = OVP2`).
  * `can.vbs`: Thao tác Hủy Warehouse Task (`METHODS` -> `@M00008`).
  * `candienmoi.vbs`: Sửa số lượng ODO và tạo Warehouse Task theo lô mới (`OIP_CHANGE`, `QTY_UI`, tab `DEFDLV`, `DEFDLV-CHARG`, `OK_OI_DLV_CREATE_AND_SAVE`).

---

## 2. Chi tiết Cải tiến & Nghiệp vụ Mới

### 2.1. Bộ điều phối Chuyên biệt Nhánh 2001 (Dedicated 2001 Dispatcher)
- Trong `ModuleMain.bas`:
  * Tại `OB_Check`: Tách nhánh ngay từ đầu `If Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("R5").Value2)) = "2001" Then OB2001_Check: EndRun: Exit Sub`.
  * Tại `OB_Execute`: Tách nhánh ngay từ đầu `If Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("R5").Value2)) = "2001" Then OB2001_Execute: Exit Sub`.
  * Đảm bảo mọi luồng xử lý của 3001 và 4007 không bị ảnh hưởng.

### 2.2. Nút 1: Đọc & Lập kế hoạch EWM (`OB2001_Check`)
- **Quét VL06O**: Lọc theo Ship-to = `2001`, Shipping Point = `3093` để tìm các Outbound Delivery Order (OB) chứa vật tư ưu tiên (UT). Nếu phát hiện OB chưa có SLoc 3093, ghi nhận vào danh sách cần điền SLoc.
- **Mở EWM Monitor (`/SCWM/MON`)**: Warehouse = `3093`.
- **Đọc ODO & Tasks**:
  * Double click Node ODO `N000000010`, nạp danh sách OB vào ô `S_ERP_H-LOW` (hỗ trợ Multiple Selection khi nhiều OB).
  * Chọn tất cả các dòng ODO trên lưới trên, bấm nút Toolbar `N000000096` để tải toàn bộ Warehouse Tasks xuống lưới dưới.
  * Đọc các trường: `MATNR` (vật tư), `CHARG` (lô), `VLPLA` (bin nguồn), `VSOLM` (số lượng kế hoạch), `NISTM` (số lượng thực), `DOCNO` (mã task), `DOCNO_H` (mã ODO/OB).
- **Đọc Tồn kho Stock Overview**:
  * Double click Node `N000000139`, lọc theo danh sách mã UT và Owner = `3093`.
  * Đọc tồn vật lý `QUAN`, tồn khả dụng `ABSQUAN`, tồn bị giữ `OUT_QUAN`.
- **Ma trận Thẩm định Nguyên Batch & FIFO Đa OB**:
  * **Mã ngoài kế hoạch (Non-UT)** nằm chung trên OB: Lập kế hoạch `CONFIRM_0` (xác nhận về 0 với mã ngoại lệ `OVP2`).
  * **Mã UT đúng batch & đủ số lượng trên 1 OB**: Lập kế hoạch `KEEP`.
  * **Mã UT đa OB đủ batch** ($A=25 + B=25 = 50$): Lập kế hoạch `KEEP_MULTI_OB` (giữ nguyên cả hai OB, không hủy nhầm).
  * **Mã UT thiếu batch đơn lẻ** (ví dụ batch 50 chỉ gắp 40, không có OB nào bù): Lập kế hoạch `CANCEL_WT` (hủy task đơn lẻ để không xuất dở dang).
  * **Mã UT vi phạm FIFO** (gắp lô mới trong khi lô cũ còn tồn khả dụng): Lập kế hoạch `CANCEL_WT` (hủy task lô mới để gom lô cũ).
- **Xuất kế hoạch**: Xuất dữ liệu trực quan lên bảng `INPUT (F:N)` và `BATCH_PLAN`, cập nhật trạng thái chi tiết cho từng dòng vật tư tại `INPUT!D`.

### 2.3. Nút 2: Thực thi theo Kế hoạch EWM (`OB2001_Execute`)
- **Kiểm tra an toàn**: Yêu cầu đã chạy Nút 1 trong phiên (`mChecked2001 = True`), kiểm tra `CONFIG!MODE = APPLY`.
- **Thực thi `CONFIRM_0`**: Gọi `MON2001_ConfirmZero` per `cfveo.vbs` (`METHODS` -> `@M00006`, tích `DIF_ALL = true`, `NISTA = 0`, `EXCCODE = OVP2`, Save Ctrl+S, Back F12, Refresh).
- **Thực thi `CANCEL_WT`**: Gọi `MON2001_CancelTask` per `can.vbs` (`METHODS` -> `@M00008`, Refresh).
- **Thực thi `KEEP` / `KEEP_MULTI_OB`**: Giữ nguyên trạng thái và ghi nhận hoàn tất.
- **Hậu kiểm & Cập nhật UI**: Tự động Refresh lưới, cập nhật cột Status `INPUT!M` sang `DONE` hoặc `FAIL`, hiển thị thông báo tổng kết số lượng hoàn thành.

---

## 3. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: PASS 100% (`control.Enabled` chuyển từ `True` sang `False`, không lỗi cú pháp).
- **Button Macro Bindings**:
  * `INPUT!OBV1__0` $ightarrow$ `'261010_OB_TOOL_V1_RC49.xlsm'!OB_Check`
  * `INPUT!OBV1__1` $ightarrow$ `'261010_OB_TOOL_V1_RC49.xlsm'!OB_Execute`
- **Khử bỏ `forceFullCalc`**: Bảo đảm `forceFullCalc="1"` bị xóa hoàn toàn khỏi `xl/workbook.xml` (`calcPr calcId="181029"/>`).
- **Mã băm SHA-256**: Tạo bảng kê SHA-256 chính xác cho workbook và handoff zip package.
