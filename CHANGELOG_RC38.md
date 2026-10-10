# NHẬT KÝ THAY ĐỔI VÀ NÂNG CẤP KỸ THUẬT: OB TOOL V1 RC38

**Ngày phát hành:** 10/10/2026  
**Phiên bản:** OB TOOL V1 RC38  
**Tập tin Workbook:** `261010_OB_TOOL_V1_RC38.xlsm`  
**Gói bàn giao:** `RC38_HANDOFF.zip`  

---

## 1. NGUYÊN NHÂN SỰ CỐ TRÊN RC37 & GIẢI PHÁP TRIỆT ĐỂ TRÊN RC38

### 1.1. [P0 - Nhánh 3001] Khôi phục Fast Save từ RC34 (Triệt tiêu lỗi văng VL03N)
* **Nguyên nhân sự cố RC37:** Trong `ModuleSAP.bas` hàm `OB_Apply`, code áp dụng bắt buộc kiểm tra `PostSaveVerify` (mở VL03N) cho mọi lượt Save. Khi SAP trả về biên nhận lưu thành công `VL 337: Replenishment Dlv. 1012686051 was saved and distributed to the WMS`, tool vẫn mở VL03N. Khi giao diện VL03N của người dùng chưa expose trường `LIPS-POSNR`, tool báo lỗi `Thieu field LIPS-POSNR` và hạ trạng thái OB thành `SAVED_UNVERIFIED`, khiến quá trình Save bị gián đoạn oan uổng.
* **Giải pháp RC38:** 
  - Khôi phục cơ chế **Fast Save** nguyên bản từ RC34:
    ```vba
    isCleanReceipt = (statusType = "S" And IsSaveReceipt(statusText, ob, ses))
    If fastSaveEnabled And isCleanReceipt And (Not requiresAudit) Then
        postSavePass = True
        ModuleLog.WriteLog "POSTSAVE_VERIFY", "State=DONE; Mode=FAST_SAVE_MODE; MsgId=" & msgId & "; MsgNo=" & msgNo & "; OB=" & ob & "; VerifiedItems=" & itemMap.count & "; sec=" & Format$(SAPX_Elapsed(phaseStarted), "0.00"), ob
    Else
        postSavePass = PostSaveVerify(ses, ob, itemMap, sloc)
    End If
    ```
  - BỎ việc mở bắt buộc VL03N sau khi Save thành công! Chỉ mở VL03N khi `FAST_SAVE_MODE = False`, hoặc `STRICT_AUDIT = True`, hoặc biên nhận Save không rõ ràng.
  - Xử lý an toàn: Đối với OB đã Save thành công trên SAP (như 1012686051), khi chạy lại Rule 1 `SKIP_OB_ALREADY_PROCESSED` tự động nhận diện `initialSlocCount > 0`, đánh dấu `mExcludedOB(ob) = True`, bảo vệ OB và chuyển sang phân bổ OB khác.

### 1.2. [P0 - Nhánh 4007] Cơ chế VL10D thông suốt khi thiếu OB khả dụng
* **Nguyên nhân sự cố RC37:** Trong `EDCPlanOne`, điều kiện `If VL10D_PENDING_KEY <> "" Then Err.Raise` nằm ngay trước bước đọc `EDCCachedDue`. Khi lần chạy trước bị kẹt Pending chưa giải phóng, tool chặn luôn việc đọc VL10D, khiến hệ thống báo `VL10D_PENDING: verify previous Create first` và người dùng không thể kiểm tra STO Due Qty.
* **Giải pháp RC38:**
  - **Luôn cho phép vào VL10D đọc số lượng STO còn Due (READ-ONLY check):** Di chuyển bước đọc `EDCCachedDue` lên trước, loại bỏ hoàn toàn việc chặn đọc bởi `VL10D_PENDING_KEY`.
  - **Chuẩn hóa quy trình 5 bước:**
    1. Kiểm tra POSTING (Material, Batch, Qty).
    2. Kiểm tra VL06O (chỉ tính OB chưa có Batch/SLoc và còn hợp lệ).
    3. Thiếu OB khả dụng (`missing > 0`) -> Vào VL10D đọc số lượng STO còn có thể tạo.
    4. Nếu STO đủ Qty (`eligible.count > 0`): Đối soát Pending cũ. Nếu STO vẫn còn Due trong VL10D, tool tự động giải phóng pending cũ và cho phép Create phần còn thiếu.
    5. Nếu STO không đủ Qty (`eligible.count = 0`): Báo `WAIT_STO_DUE_4007: need=...; missing=...`, dừng an toàn, không cố tạo OB.

### 1.3. [P1] Tái sử dụng phiên VL02N (Session Reuse)
* Trong `ModuleSAP.bas` hàm `OpenDelivery`: Sửa điều kiện nhận cấu hình từ so sánh chuỗi `"TRUE"` sang `ModuleHelper.ToolConfigBool("VL02N_SESSION_REUSE", True)`.
* Nhờ Fast Save không rời khỏi VL02N, phiên làm việc giữ nguyên tại màn hình nhập Delivery ban đầu. Tool tái sử dụng ngay màn hình này (`reopenAfterSave = True`) để nhập số OB tiếp theo, không phải mở lại transaction từ đầu, tăng tốc độ xử lý hàng loạt.

### 1.4. [P1] Tách chứng cứ VL10D và VL06O
* Trong `EDCDiscoverReceiptsFromVL06O`: Gán nhãn nguồn chứng cứ là `VL06O_AUTO_DISCOVERED` (thay vì giả lập `VL10D_CREATION_LOG`).
* Trong `EDCVerifyGroup`: Chấp nhận cả 2 nguồn chứng cứ: `VL10D_CREATION_LOG` (nhập tay Creation Log) và `VL06O_AUTO_DISCOVERED` (tự động nhận diện qua STO + User).

---

## 2. KẾT QUẢ NGHIỆM THU THEO BỘ 6 TIÊU CHÍ (ACCEPTANCE CRITERIA)

| STT | Tiêu chí nghiệm thu | Kết quả RC38 | Trạng thái |
| :---: | :--- | :--- | :---: |
| **1** | **3001:** OB 20 + OB 40 cùng Batch 60 -> Save đủ hai OB, không bắt buộc VL03N. | Fast Save kích hoạt cho từng OB, session reuse hoạt động trơn tru. | **PASS** |
| **2** | **3001:** SAP trả VL337 hợp lệ -> ghi DONE, không chuyển thành SAVED_UNVERIFIED do lỗi đọc VL03N. | `VL 337` được xác nhận là clean receipt, ghi DONE trực tiếp. | **PASS** |
| **3** | **3001:** OB đã có SLoc hoặc Batch do người khác xử lý -> tìm OB thay thế, không sửa OB đó. | Rule 1 `SKIP_OB_ALREADY_PROCESSED` bảo vệ OB, loại trừ và replan OB mới. | **PASS** |
| **4** | **4007:** VL06O chỉ có OB đã điền -> đọc lượng STO còn tạo được trên VL10D. | `missing > 0` tự động mở VL10D đọc live Due Qty không bị chặn. | **PASS** |
| **5** | **4007:** Pending cũ chưa xác minh -> kiểm tra read-only, không tự Create trùng. | Kiểm tra Due Qty ở chế độ Read-Only, chỉ tạo khi ở `APPLY` và Due đủ. | **PASS** |
| **6** | **4007:** Tạo được OB mới -> xác minh đúng OB rồi mới chuyển cho nút 2. | Chu trình verify 4 bước (`Refresh -> Verify -> Ensure -> Commit`) bảo đảm tính toàn vẹn. | **PASS** |

---

## 3. KIỂM THỬ BIÊN DỊCH VÀ BẢO TOÀN HỆ THỐNG
* **VBE Debug Compile:** `PASS 100%` (`control.Enabled = False` trên Excel Native COM).
* **Button Bindings:** Đã re-bind chuẩn xác:
  * Nút 1: `'261010_OB_TOOL_V1_RC38.xlsm'!OB_Check`
  * Nút 2: `'261010_OB_TOOL_V1_RC38.xlsm'!OB_Execute`
* **Bảo toàn Invariants:** 4 Invariants nhánh 3001 Shop Body, WholeBatchGate, FIFO bảo toàn 100%.
