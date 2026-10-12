# OB TOOL V1 RC49

Canonical XLSM: `261010_OB_TOOL_V1_RC49.xlsm`  
Canonical SHA256: `A199AB73A70EEE04820C0B9F538ABE7518F39C5FED75569F064E661AA43381F3`  
Handoff Archive: `RC49_HANDOFF.zip`  
Archive SHA256: `D26D245F22606301B76DD60C4B39F0FC87235E7B1AB316660C5368D7650287CB`

## Key Improvements in RC49
- **Bộ điều phối chuyên biệt Nhánh 2001 (Shop XMĐ) trên EWM Monitor (`/SCWM/MON`)**:
  - Rẽ nhánh độc lập khi ô `INPUT!R5 = "2001"`, bảo toàn nguyên vẹn 100% logic của nhánh 3001 và 4007 khi `R5 <> "2001"`.
- **Nút 1: Đọc & Lập kế hoạch EWM (`OB2001_Check`)**:
  - Đọc danh sách mã gọi ưu tiên (UT) từ `INPUT!A6:C205`.
  - Quét `VL06O` theo Ship-to = `2001`, Shipping Point = `3093` để thu thập các Outbound Delivery (OB) chứa mã UT; tự động phát hiện OB thiếu SLoc 3093.
  - Mở `/SCWM/MON` (Warehouse 3093), truy vấn Node ODO (`N000000010`) qua Multiple Selection `S_ERP_H-LOW`, tải lưới Warehouse Tasks qua nút Toolbar `N000000096` (`MATNR`, `CHARG`, `VLPLA`, `VSOLM`, `NISTM`, `DOCNO`, `DOCNO_H`).
  - Truy vấn Node Stock Overview (`N000000139`) với Owner `3093` để đọc tồn vật lý `QUAN`, tồn khả dụng `ABSQUAN`, tồn giữ `OUT_QUAN`.
  - **Ma trận Thẩm định Nguyên Batch & FIFO Đa OB**:
    * Mã ngoài kế hoạch (Non-UT) nằm chung trên OB: Lập kế hoạch `CONFIRM_0` (mã ngoại lệ `OVP2`).
    * Mã UT đúng batch và đủ số lượng: Lập kế hoạch `KEEP`.
    * Mã UT đa OB đủ batch (A=25 + B=25 = 50): Lập kế hoạch `KEEP_MULTI_OB` (giữ nguyên cả hai OB).
    * Mã UT thiếu batch đơn lẻ: Lập kế hoạch `CANCEL_WT` (hủy task đơn lẻ).
    * Mã UT vi phạm FIFO (gắp lô mới khi lô cũ còn tồn): Lập kế hoạch `CANCEL_WT` (hủy lô mới để gom lô cũ).
  - Xuất bảng kế hoạch trực quan lên `INPUT (F:N)` và `BATCH_PLAN`.
- **Nút 2: Thực thi theo Kế hoạch EWM (`OB2001_Execute`)**:
  - Thực thi `CONFIRM_0` per `cfveo.vbs` (`METHODS` -> `@M00006`, `ODP1_CONF_FGR`, `DIF_ALL = true`, `NISTA = 0`, `EXCCODE = OVP2`, Save, Back, Refresh).
  - Thực thi `CANCEL_WT` per `can.vbs` (`METHODS` -> `@M00008`, Refresh).
  - Sửa số lượng ODO và tạo task theo batch mới per `candienmoi.vbs` (`OIP_CHANGE`, `QTY_UI`, VKey 25, tab `DEFDLV`, `DEFDLV-CHARG`, `OK_OI_DLV_CREATE_AND_SAVE`).
  - Refresh và đối soát đọc lại kết quả xác minh.
- **Biên dịch & Bảo toàn Hệ thống**:
  - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
  - Gắn macro nút bấm chuẩn xác (`INPUT!OBV1__0` -> `OB_Check`, `INPUT!OBV1__1` -> `OB_Execute`).
  - Khử bỏ triệt để `forceFullCalc="1"` trong `xl/workbook.xml`.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC49 (13/10/2026)

### 📦 Các tệp đính kèm RC49

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC49_HANDOFF.zip`** | 693,922 bytes | `D26D245F22606301B76DD60C4B39F0FC87235E7B1AB316660C5368D7650287CB` | **Gói bàn giao đầy đủ RC49 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC49.xlsm`** | 742,593 bytes | `A199AB73A70EEE04820C0B9F538ABE7518F39C5FED75569F064E661AA43381F3` | Tệp Excel Macro phát hành RC49 (Đã dọn dẹp & Compile) |
| **`VBA_RC48_TO_RC49.diff`** | 30,450 bytes | `10A92864176CB06E5BA27E8B5C98A68DD2CE06719F32D2E8ADDC0488EA88294A` | Chi tiết bản vá mã nguồn VBA từ RC48 lên RC49 |
| **`CHANGELOG_RC48_TO_RC49.md`** | 5,212 bytes | `62B801C61C5A6A21E4455FF33646E9F7C3E9C7CCEA31D6480FE195F6ECA4EF9D` | Nhật ký thay đổi và giải pháp kỹ thuật RC49 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC48 (13/10/2026)
* **`RC48_HANDOFF.zip`**: SHA-256 `0C88DDD6DAEAFF234766D4A3DF4D976B319AFF5D545425417F6625D0D7D40AA7`
* **`261010_OB_TOOL_V1_RC48.xlsm`**: SHA-256 `AC5A70B941DB7D5E49E92ABBBAF033CC9CBE9608C7C40B13ADE82D5070F4AE3C`

### OB Tool V1 RC47 (11/10/2026)
* **`RC47_HANDOFF.zip`**: SHA-256 `A9F845EC80F4DB7D0C03B78E9E5060964A3559E90A0838BDE2EAD022D002ACA8`
* **`261010_OB_TOOL_V1_RC47.xlsm`**: SHA-256 `B90CA4209A13806027C2234D32608E2F2C583F76074979BC048C3C450C3604EC`

### OB Tool V1 RC46 (11/10/2026)
* **`RC46_HANDOFF.zip`**: SHA-256 `AA012032530A6DE17D10AAA5A4C2DA8BD4A1F4798B678E3BFAC577F31A39CDAC`
* **`261010_OB_TOOL_V1_RC46.xlsm`**: SHA-256 `D3121B65845EBDFDA4AEF1DA84C9BB7B92D7E4E496D8A18918B98FD7C2F5D5A7`

### OB Tool V1 RC45 (11/10/2026)
* **`RC45_HANDOFF.zip`**: SHA-256 `58B4307C0F4A5E4F8B70DF301DE783E00DA5467F5734A1B367EF9B9B88DE75A3`
* **`261010_OB_TOOL_V1_RC45.xlsm`**: SHA-256 `C19C45244199E5C56BF47DC83BE28BD7CA36F3D431C39257C98622D97394541B`
