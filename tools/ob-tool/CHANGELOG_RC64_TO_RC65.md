# CHANGELOG - OB TOOL V1 RC65

## 1. Tổng quan Bản phát hành RC65
Bản phát hành RC65 tập trung giải quyết triệt để sự cố phân phối WMS (VL589), sự cố sập đọc lại OB trên VL03N và tự chủ hóa quy trình tạo OB nhánh 3001:

1. **Khôi phục triệt để cơ chế xóa dòng mã ngoài (Cancel unplanned items) trong OB 3001**:
   - Loại bỏ hoàn toàn điều kiện chặn `InputProducts().exists` trong `RemoveUnplannedItems`.
   - Bất kỳ item nào tồn tại trong OB nhưng không nằm trong kế hoạch phân bổ của lượt đó đều bị xóa bỏ qua `btnBT_POLO_T` hoặc phím tắt `Shift+F2`.
   - Thiết lập chốt chặn nghiêm ngặt **Pre-Save Gate** trong `VerifyDeliveryPlan`: Tuyệt đối không cho phép Save nếu còn bất kỳ item ngoài kế hoạch nào sót lại trong OB, ngăn chặn 100% nguy cơ lỗi WMS `VL589: Replenishment Dlv. ... could not be distributed`.

2. **Khắc phục triệt để lỗi sập đọc lại chứng từ (`CONFLICT_READBACK_UNAVAILABLE: Object variable or With block variable not set`)**:
   - Trong `OB_ResolveUnknown` (`ModuleSAP.bas`): Thêm bẫy an toàn `On Error GoTo RecoveryFail` và kiểm tra `If ses Is Nothing` / `If tbl Is Nothing` trước khi đọc bảng `VL03N`.
   - Bẫy toàn bộ pop-up phát sinh sau chuyển màn hình, ngăn ngừa lỗi runtime 91 làm sập luồng phục hồi.

3. **Cơ chế quét VL10D 1 lần & tự sinh OB cho Nhánh 3001 (Nếu hết STO thì bỏ qua batch thiếu)**:
   - Khi kho 3001 thiếu OB do 4007 vừa xóa nhả tồn hoặc DO AUTO chưa kịp sinh: Tool tự động quét `VL10D` đúng 1 lần với `Receiving Plant = 3001`, `Supplying Plant = 3093`.
   - Nếu STO còn tồn mở (`Due > 0`): Tự động tạo OB bù đúng số lượng nguyên batch -> Refresh đối soát -> Điền Batch/SLoc.
   - Nếu STO không còn tồn (`Due = 0`): Dừng và **bỏ qua batch thiếu đó (`SKIP_NO_STO_DUE_3001`)**, không gây lỗi sập tool, tiếp tục chạy trọn vẹn các mã/batch còn lại.

---

## 2. Chi tiết Kỹ thuật Các Thay Đổi

### 2.1. `ModuleSAP.bas`
- `RemoveUnplannedItems`:
  * Xóa bỏ kiểm tra `If ModuleMain.InputProducts().exists(...) Then keep(...) = True`.
  * Thêm cơ chế Fallback Menu `Edit -> Delete Item` (`Shift+F2`) nếu nút `btnBT_POLO_T` không phản hồi.
- `VerifyDeliveryPlan`:
  * Nâng cấp `PRESAVE_EXTRA_ITEM_INFO` thành `PRESAVE_GATE_FAIL: EXTRA_ITEM_NOT_CLEARED`, chặn Save nếu còn dòng lạ chưa được dọn dẹp.
- `OB_ResolveUnknown`:
  * Bổ sung bẫy Nothing cho Session, Table, và xử lý popup chuyển màn hình `VL03N`.
- `EDC_SetField ses, "ST_WERKS-LOW", edcShipTo`:
  * Động hóa giá trị Receiving Plant trên `VL10D` theo cấu hình `INPUT!R5` (hỗ trợ cả `3001` và `4007`).

### 2.2. `ModuleMain.bas`
- `EDCBuildCreatePlans`:
  * Cho phép cả `4007` và `3001` lập kế hoạch tạo OB trên `VL10D`.
- `EDCPlanOne`:
  * Bổ sung nhánh xử lý riêng cho `ShipTo = "3001"`: Khi `eligible.count = 0` (hết STO Due), ghi log `SKIP_NO_STO_DUE_3001`, ghi chú trên INPUT: `"THIEU OB (STO 3001 HET TON DUE - BO QUA)"` và thoát êm dịu, không ném ngoại lệ dừng toàn bộ tool.
- `EDCCreateGroup`:
  * Bỏ qua gọi lệnh xóa nháp 3001 (`EDCReleaseDrafts3001`) khi đang xử lý cho chính nhánh 3001.
