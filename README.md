# OB TOOL V1 RC48

Canonical XLSM: `261010_OB_TOOL_V1_RC48.xlsm`  
Canonical SHA256: `AC5A70B941DB7D5E49E92ABBBAF033CC9CBE9608C7C40B13ADE82D5070F4AE3C`  
Handoff Archive: `RC48_HANDOFF.zip`  
Archive SHA256: `0C88DDD6DAEAFF234766D4A3DF4D976B319AFF5D545425417F6625D0D7D40AA7`

## Key Improvements in RC48
- **Nâng giới hạn STO từ 2 lên tối đa 4 (Plant 4007 STO)**:
  - Trong `ModuleMain.bas` (`EDCNormalizeSTOs`), hỗ trợ chuẩn hóa phân tách bằng `;`, phẩy, xuống dòng (`vbCrLf`), tab; nâng giới hạn lên tối đa 4 STO.
  - Trên sheet `INPUT`, cập nhật nhãn ô `P31` thành `STO 4007 (tối đa 4)`.
- **Gom mã đọc VL10D tối ưu cho nhiều STO (Bulk Multi-STO Read)**:
  - Trong `ModuleSAP.bas`, bổ sung `EDCSetGroupSTOFilter` kết hợp Multiple Selection `%_ST_EBELN_%_APP_%-VALU_PUSH` và `%_ST_MATNR_%_APP_%-VALU_PUSH`.
  - Mở VL10D đúng 1 lần duy nhất cho toàn bộ danh sách mã và danh sách STO (hàm `EDC4007_FindDueMultiSTO`).
  - Lưu toàn bộ kết quả Due vào bộ nhớ đệm RAM theo Material và STO (`EDCCachedDue`), loại bỏ 100% việc mở lặp lại VL10D từng mã × từng STO.
  - Khi tạo OB, gom các mã theo STO đã chọn và kiểm tra đúng dòng trước Create.
- **Quản lý Calculation Mode & Tối ưu hiệu năng Excel**:
  - Gỡ bỏ triệt để thuộc tính `forceFullCalc="1"` trong `xl/workbook.xml` và đặt `wb.ForceFullCalculation = False`, triệt tiêu tình trạng đơ/lag các file Excel khác.
  - Trong `ModuleMain.bas`, lưu trạng thái ban đầu của `Application.Calculation`, `EnableEvents`, `ScreenUpdating` tại `BeginRun` (chuyển sang Manual), khôi phục đúng trạng thái ban đầu khi kết thúc, không bao giờ đơn phương ép Automatic.
  - Loại bỏ thao tác AutoFit lặp từng dòng trong `UpdateItemStatus`.
- **Biên dịch & Bảo toàn Hệ thống**:
  - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
  - Gắn macro nút bấm chuẩn xác (`OB_Check` và `OB_Execute`).
  - Bảo toàn trọn vẹn 4 Invariants nhánh 3001 Shop Body, WholeBatchGate và nguyên tắc FIFO.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC48 (13/10/2026)

### 📦 Các tệp đính kèm RC48

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC48_HANDOFF.zip`** | 652,565 bytes | `0C88DDD6DAEAFF234766D4A3DF4D976B319AFF5D545425417F6625D0D7D40AA7` | **Gói bàn giao đầy đủ RC48 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC48.xlsm`** | 710,356 bytes | `AC5A70B941DB7D5E49E92ABBBAF033CC9CBE9608C7C40B13ADE82D5070F4AE3C` | Tệp Excel Macro phát hành RC48 (Đã dọn dẹp & Compile) |
| **`VBA_RC47_TO_RC48.diff`** | 10,720 bytes | `06BBA607E42052ED2E4F036E30EF8FD0F50B021BD182A17A60C521F2DD7899B3` | Chi tiết bản vá mã nguồn VBA từ RC47 lên RC48 |
| **`CHANGELOG_RC48.md`** | 5,502 bytes | `55EE70D2217BAE3F934C7250E109AA5EFF1864098DA1DB4C07BCF0D87FFAAD86` | Nhật ký thay đổi và giải pháp kỹ thuật RC48 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC47 (11/10/2026)
* **`RC47_HANDOFF.zip`**: SHA-256 `A9F845EC80F4DB7D0C03B78E9E5060964A3559E90A0838BDE2EAD022D002ACA8`
* **`261010_OB_TOOL_V1_RC47.xlsm`**: SHA-256 `B90CA4209A13806027C2234D32608E2F2C583F76074979BC048C3C450C3604EC`

### OB Tool V1 RC46 (11/10/2026)
* **`RC46_HANDOFF.zip`**: SHA-256 `AA012032530A6DE17D10AAA5A4C2DA8BD4A1F4798B678E3BFAC577F31A39CDAC`
* **`261010_OB_TOOL_V1_RC46.xlsm`**: SHA-256 `D3121B65845EBDFDA4AEF1DA84C9BB7B92D7E4E496D8A18918B98FD7C2F5D5A7`

### OB Tool V1 RC45 (11/10/2026)
* **`RC45_HANDOFF.zip`**: SHA-256 `58B4307C0F4A5E4F8B70DF301DE783E00DA5467F5734A1B367EF9B9B88DE75A3`
* **`261010_OB_TOOL_V1_RC45.xlsm`**: SHA-256 `C19C45244199E5C56BF47DC83BE28BD7CA36F3D431C39257C98622D97394541B`
