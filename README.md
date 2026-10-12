# OB TOOL V1 RC51

Canonical XLSM: `261010_OB_TOOL_V1_RC51.xlsm`  
Canonical SHA256: `85966347A480D4AF990881B481A6F860C3950DDEFF6991653BBA0A1615CFCBD9`  
Handoff Archive: `RC51_HANDOFF.zip`  
Archive SHA256: `DB86E5724A2153D43D50606FB0E456C3D9F101F6DBABFCCEC917DBB7AC9F9B50`

## Key Improvements in RC51
- **Vá lỗi Runtime Error 424 "Object required" trong Dominant STO Clustering (`ModuleMain.bas - EDCPlanOne`)**:
  - Khắc phục triệt để lỗi ngoại lệ khi kiểm tra `matchedCand Is Nothing` trên biến Mảng (Array).
  - Thay thế toàn diện bằng biến cờ Boolean `hasMatched As Boolean` trong thuật toán gom nhóm STO chủ đạo.
  - Đảm bảo toàn bộ các vật tư trong STO chủ đạo được lập kế hoạch tạo OB thành công và gửi lệnh lên VL10D, không còn tình trạng kế hoạch 0 item.
- **Toàn bộ cải tiến vững chắc từ RC50 được bảo toàn 100%**:
  - **P0**: Lưu create intent bất biến, lọc bỏ baseline OBs, loại bỏ lỗi gán đè Qty kế hoạch (`obQty <= 0`), cam kết chỉ commit các kế hoạch đã xác minh (`verifiedPlans.Count > 0`).
  - **P1**: Giải phóng draft 3001 dựa trên thâm hụt đã xác minh, chờ đồng bộ ATP SAP 2 giây trước khi tự động thử lại tạo VL10D.
  - **P2**: Nhật ký UoM chi tiết toàn diện (`raw`, `normalized`, `expected`, `row`, `x`, `control ID`), hàm `NormalizeUoM` dọn sạch khoảng trắng ẩn, nhóm đơn vị đếm tương đương `EA/UN/PC/ST/PCE`, hỗ trợ định dạng số thập phân trong VL02N (`FormatSAPQty`).
  - **P3**: Tự động gom cụm STO chủ đạo (Dominant STO Clustering), loại bỏ hoàn toàn các hộp thoại `Application.InputBox` và `MsgBox`.
- **Biên dịch & Bảo toàn Hệ thống**:
  - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
  - Gắn macro nút bấm chuẩn xác (`INPUT!OBV1__0` -> `OB_Check`, `INPUT!OBV1__1` -> `OB_Execute`).
  - Khử bỏ triệt để `forceFullCalc="1"` trong `xl/workbook.xml`.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC51 (13/10/2026)

### 📦 Các tệp đính kèm RC51

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC51_HANDOFF.zip`** | 694,251 bytes | `DB86E5724A2153D43D50606FB0E456C3D9F101F6DBABFCCEC917DBB7AC9F9B50` | **Gói bàn giao đầy đủ RC51 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC51.xlsm`** | 748,030 bytes | `85966347A480D4AF990881B481A6F860C3950DDEFF6991653BBA0A1615CFCBD9` | Tệp Excel Macro phát hành RC51 (Đã vá 424 & Compile) |
| **`VBA_RC50_TO_RC51.diff`** | 1,667 bytes | `D9B2B4F4B60E20E59868C40AA11B30494D43DB2509974232B2D48DA81B274AB3` | Chi tiết bản vá mã nguồn VBA từ RC50 lên RC51 |
| **`CHANGELOG_RC50_TO_RC51.md`** | 3,826 bytes | `BA8A15C32B4319D45E8B2FC0BFE1B0413A22242936378091505235F3BBF5332D` | Nhật ký thay đổi và giải pháp kỹ thuật RC51 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC50 (13/10/2026)
* **`RC50_HANDOFF.zip`**: SHA-256 `7CB1088B32B6ADEC280737F49DCC1CB8130FA997F376AD8D6810E4E860F6BC9E`
* **`261010_OB_TOOL_V1_RC50.xlsm`**: SHA-256 `545E9C5CAB9C3F0BACC34D53DFF840B3B8C3618853DE72E5760A87D29A7ECAC4`

### OB Tool V1 RC49 (13/10/2026)
* **`RC49_HANDOFF.zip`**: SHA-256 `BCFFCF5F3D1865EAC3FA770BC78DEBF3B8083884358896EEFEA3AC3DF38E90BC`
* **`261010_OB_TOOL_V1_RC49.xlsm`**: SHA-256 `A199AB73A70EEE04820C0B9F538ABE7518F39C5FED75569F064E661AA43381F3`

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
