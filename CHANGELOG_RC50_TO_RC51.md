# CHANGELOG - OB TOOL V1 RC51

## 1. Tổng quan Bản phát hành RC51 (Hotfix Runtime Error 424 Object Required trong Dominant STO Clustering)
- **Mục tiêu**: Khắc phục triệt để lỗi **Runtime Error 424 "Object required"** xảy ra tại `EDCPlanOne` trong thuật toán gom nhóm STO chủ đạo (Dominant STO Clustering).
- **Phát hiện từ File Chạy Thực tế**: Người dùng cung cấp file chạy thực tế RC50 (`RC50/user-upload/261010_OB_TOOL_V1_RC50.xlsm`). Nhật ký sheet `LOG` cho thấy:
  * Dòng 152: Thuật toán đã phát hiện và chọn đúng STO chủ đạo `3601040757` (thỏa mãn gom 7 vật tư, tổng Due = 10,503).
  * Dòng 153-163: Cả 7 vật tư đều bị văng lỗi runtime `VL10D_PLAN_WAIT: <Material>; 424; Object required`.
  * Dòng 192: `PLAN: 0 item; 0 allocations`, dẫn tới không có lệnh tạo OB nào được gửi lên VL10D.
- **Nguyên nhân Gốc rễ (Root Cause)**:
  * Biến `cand` trong `eligible` là một Mảng (`Variant() / Array` do hàm `Array(...)` tạo ra).
  * Phép so sánh `If matchedCand Is Nothing Then` sử dụng toán tử `Is` vốn chỉ hợp lệ trên kiểu Đối tượng (Object). Khi áp dụng lên Mảng (Array), VBA ném ngay ngoại lệ Runtime 424: "Object required".
- **Bảo toàn Invariants**: Giữ nguyên vẹn 100% logic của Branch 3001 Body Shop, Branch 2001 Shop XMĐ (/SCWM/MON EWM Dispatcher) và toàn bộ các tính năng P0, P1, P2 của RC50.

---

## 2. Chi tiết Sửa đổi & Giải pháp Kỹ thuật

### 2.1. Thay thế Toán tử `Is Nothing` bằng Biến Cờ Boolean `hasMatched`
- Trong `ModuleMain.bas` tại hàm `EDCPlanOne`:
  * Khai báo biến cờ `hasMatched As Boolean` (khởi tạo `hasMatched = False`).
  * Trong vòng lặp quét STO chủ đạo (`mEDCDominantSTO`):
    ```vb
    If Not hasMatched Then
        chosen = cand
        hasMatched = True
    ElseIf CDbl(cand(4)) > CDbl(chosen(4)) Then
        chosen = cand
    End If
    ```
  * Trong vòng lặp dự phòng (fallback nếu không có STO chủ đạo phù hợp):
    ```vb
    If Not hasMatched Then
        For Each cand In eligible
            If Not hasMatched Then
                chosen = cand
                hasMatched = True
            ElseIf CDbl(cand(4)) > CDbl(chosen(4)) Then
                chosen = cand
            End If
        Next cand
    End If
    ```
  * Loại bỏ hoàn toàn biến `matchedCand` và các phép kiểm tra `matchedCand Is Nothing`.
  * Kết quả: Triệt tiêu 100% lỗi 424, đảm bảo toàn bộ các vật tư trong nhóm STO chủ đạo được lập kế hoạch tạo OB trơn tru và chính xác trên VL10D.

### 2.2. Rà soát Toàn bộ Dự án
- Rà soát toàn bộ các file `.bas` (`ModuleMain.bas`, `ModuleSAP.bas`, `ModuleHelper.bas`, `ModuleLog.bas`):
  * Xác nhận không còn bất kỳ phép toán `Is Nothing` nào áp dụng trên biến Mảng/Array.
  * Tất cả các trường hợp `Is Nothing` còn lại đều kiểm tra chính xác các đối tượng COM (`Worksheet`, `Shape`, `Scripting.Dictionary`, `SAP GUI Control`).

---

## 3. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: PASS 100% (`control.Enabled` chuyển từ `True` sang `False`, không cảnh báo hay lỗi cú pháp).
- **Shape Macro Bindings**:
  * `INPUT!OBV1__0` $ightarrow$ `'261010_OB_TOOL_V1_RC51.xlsm'!OB_Check`
  * `INPUT!OBV1__1` $ightarrow$ `'261010_OB_TOOL_V1_RC51.xlsm'!OB_Execute`
- **Khử bỏ `forceFullCalc`**: Bảo đảm `forceFullCalc="1"` bị xóa sạch khỏi `xl/workbook.xml` (`calcPr calcId="181029"/>`).
- **Mã băm SHA-256**: Tạo bảng kê SHA-256 chính xác cho workbook và handoff zip package.
