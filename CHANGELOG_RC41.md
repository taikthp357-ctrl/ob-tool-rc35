# NHẬT KÝ THAY ĐỔI VÀ NÂNG CẤP KỸ THUẬT: OB TOOL V1 RC41

**Ngày phát hành:** 10/10/2026  
**Phiên bản:** OB TOOL V1 RC41  
**Tập tin Workbook:** `261010_OB_TOOL_V1_RC41.xlsm`  
**Gói bàn giao:** `RC41_HANDOFF.zip`  

---

## 1. TỔNG QUAN VÀ BỐI CẢNH THỰC TẾ

Sau chuỗi thử nghiệm thực tế với phiên bản RC40:
- **Nhánh 3001**: Đạt tỷ lệ thành công tuyệt đối **100% trên toàn bộ 19 OB**! Các tính năng Fast Save Mode (lưu tức thì không qua VL03N), Direct Item Locate (không cuộn trang xuống đáy để tìm item, không quét dư thừa), Batch Splits và Smart Lock Backoff hoạt động mượt mà và ổn định.
- **Nhánh 4007**: Sau khi RC40 gỡ bỏ chặn cứng `ST_RESWK` tại màn hình chọn VL10D, tool đã gửi phím F8 thành công, vào được màn hình danh sách Due của VL10D và đọc chính xác STO `3601040480|610` với số lượng Due còn lại là **280 pcs**.
- **Điểm nghẽn cần tối ưu (RC40)**:
  - Khi INPUT yêu cầu số lượng 680 pcs (`missing = 680`), logic kiểm tra của RC40 yêu cầu STO Due phải lớn hơn hoặc bằng toàn bộ nhu cầu (`candidate(4) >= missing`).
  - Vì $280 < 680$, hệ thống coi như không có STO nào thỏa mãn (`eligible.count = 0`), dừng lại và báo:  
    `WAIT_STO_DUE_4007: need=680; OB capacity=0; missing=680`  
    khiến tool không tạo được OB cho 280 pcs sẵn có.

---

## 2. NÂNG CẤP VÀ HOÀN THIỆN TRÊN RC41

Phiên bản **RC41** triển khai cơ chế **Phân bổ STO Due Linh Hoạt (Partial STO Due Allocation & Dynamic Create)**:

### 2.1. Phân bổ STO Due từng phần (Partial STO Due Allocation)
* **Thuật toán mới:**
  $$\text{CreateQty} = \min(\text{missing}, \text{Due})$$
* **Thay đổi chi tiết trong ModuleMain.bas:**
  1. `EDCCheckOne`: Chấp nhận STO có `Due > 0`:  
     `If CDbl(a(4)) > 0 Then count = count + 1`
  2. `EDCPlanOne`:
     - Bỏ điều kiện cứng `>= missing`, chấp nhận mọi STO còn Due:  
       `If CDbl(candidate(4)) > 0 Then eligible.Add candidate`
     - Tính toán số lượng tạo thực tế:  
       `createQty = Application.Min(missing, CDbl(chosen(4)))`
     - Lưu `createQty` vào kế hoạch tạo OB `mEDCCreatePlan`.
  3. `EDCCreateInputOne` & `EDCExecuteCreatePlans`:
     - Giới hạn số lượng nhập vào VL10D theo đúng `createQty` đã lập:  
       `need = Application.Min(Application.Max(0, requiredNow - CDbl(planned(5))), CDbl(planned(6)))`
     - Đảm bảo không bao giờ yêu cầu vượt quá số lượng Due khả dụng của STO.

### 2.2. Luồng tự động hoàn chỉnh cho nhánh 4007
* **Bấm Nút 1**:
  - Tự động vào VL10D, chọn STO có Due khả dụng.
  - Tạo mới OB với đúng số lượng tối đa STO có thể đáp ứng (ví dụ: tạo OB 280 pcs).
  - Tự động làm mới danh sách VL06O để nạp ngay OB vừa tạo.
  - Phân bổ 280 pcs này cho các lô hàng trong POSTING (ví dụ: 7 lô $\times$ 40 = 280 pcs).
  - Báo cáo rõ ràng phần còn thiếu (400 pcs) do STO hết hạn mức Due.
* **Bấm Nút 2**:
  - Tiến hành điền SLoc 3093 và Batch vào OB vừa tạo như bình thường theo đúng quy trình Fast Path.

---

## 3. BẢO TOÀN CÁC TIÊU CHUẨN KỸ THUẬT VÀ INVARIANTS

1. **Non-target Item Preservation (Invariant 1):** Giữ nguyên vẹn toàn bộ các mặt hàng khác trong cùng Delivery.
2. **WholeBatchGate (Invariant 2):** Tuyệt đối không tách dở dang 1 batch; mỗi batch được phân bổ trọn vẹn số lượng.
3. **FIFO Allocation (Invariant 3):** Ưu tiên các Delivery cũ tạo trước theo đúng thứ tự thời gian.
4. **Fast Save Mode & Direct Item Locate:** Lưu tốc độ cao ngay sau thông báo VL337; định vị trực tiếp dòng cần xử lý mà không cần cuộn trang xuống đáy.
5. **VBE Debug Compile:** PASS 100% (Native compile verified trên Excel VBA).
6. **Sheet Cleanliness:** Sheet `LOG` và `LEDGER` được dọn sạch lịch sử dư thừa, đảm bảo workbook vận hành siêu tốc.
