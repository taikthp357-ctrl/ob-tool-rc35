# OB Tool V1 RC35 R5 — Public Release & Handoff Package

Bản nâng cấp Excel Macro giải quyết luồng Plant 4007 STO (pending create) và bảo toàn 4 Invariants nhánh 3001 Shop Body cho công cụ OB Tool (SAP / Excel).

---

## 📦 Tệp đính kèm trong kho lưu trữ

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC35_R5_HANDOFF.zip`** | 970,406 bytes | `BBAAD4ABAF3C91C3B428C5CF177CEEC0D8FE94585F070ACBFFC4839CDD122A74` | **Gói bàn giao đầy đủ (All-in-one)** |
| **`261010_OB_TOOL_V1_RC35_R5.xlsm`** | 1,059,734 bytes | `98EE42AD70097067436EFD674AE069F1B9C2A72D66BEC3D4856394E2A833FF53` | Tệp Excel Macro phát hành RC35 R5 |
| **`VBA_R4_TO_R5.diff`** | 4,505 bytes | — | Chi tiết bản vá mã nguồn VBA từ R4 lên R5 |
| **`REPAIR_AND_QA.md`** | 1,917 bytes | — | Báo cáo nghiệm thu kỹ thuật và kiểm thử native |
| **`SHA256_MANIFEST.json`** | 1,299 bytes | — | Bảng mã băm xác thực toàn vẹn từng thành phần |

---

## 🔍 Điểm cải tiến & Sửa lỗi trọng tâm trên RC35 R5

1. **Khắc phục lỗi VBA Active Error Handler (`ModuleMain.bas:3740-3751`):**
   * Trong vòng lặp `targeted readback` (VL03N retry handler), bổ sung `Resume RetryWait` để giải phóng cờ lỗi của VBA trước khi chuyển sang lần thử 2.
   * Ngăn chặn việc VBA ném lỗi thô ra ngoài và đảm bảo đóng gói chính xác thông báo chẩn đoán `GROUP_OB_READBACK_WAIT: ...; no Create retry.` khi không tìm thấy OB.

2. **Cổng kiểm soát danh sách snapshot (`EDCEnsureGroupInSnapshot`):**
   * Xác thực và đối soát OB/delivery item vừa tạo trên snapshot `VL06O_DATA` trước khi ghi nhận hoàn tất.
   * Ngăn chặn việc kế hoạch phân bổ sử dụng snapshot cũ gây sai lệch.

3. **Bảo toàn tính an toàn giao dịch SAP:**
   * Giữ nguyên khóa pending `BATCH|REQ_20261010_042845_1|3601034990` (Material `BIW32012300AA`, STO item `70`, Qty `28`) để chống trùng lặp Create khi chưa có biên lai thật từ SAP Creation Log.
   * Giữ nguyên 100% 4 Invariants cho nhánh 3001 Shop Body (WholeBatchGate, skip khi có sẵn SLoc, xóa sạch non-plan items, SLoc bắt buộc 100% dòng xuất).

---

## 🧪 Kết quả kiểm thử nội bộ

* **VBE Debug Compile:** 100% PASS (0 lỗi cú pháp, Macro shape bindings nguyên vẹn).
* **Native Caller-Gate Tests:** 20/20 ca kiểm thử đạt chuẩn PASS (100%).
* **Tình trạng SAP Live:** NOT_RUN (chưa kết nối phiên SAP GUI trực tiếp; tệp ở chế độ Code Ready / Local QA Verified).
