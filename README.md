# OB TOOL V1 RC50

Canonical XLSM: `261010_OB_TOOL_V1_RC50.xlsm`  
Canonical SHA256: `545E9C5CAB9C3F0BACC34D53DFF840B3B8C3618853DE72E5760A87D29A7ECAC4`  
Handoff Archive: `RC50_HANDOFF.zip`  
Archive SHA256: `7CB1088B32B6ADEC280737F49DCC1CB8130FA997F376AD8D6810E4E860F6BC9E`

## Key Improvements in RC50
- **P0 — Sửa Xác minh Receipt và Dữ liệu Thiếu Trước Khi Cho Phép Cancel**:
  - Lưu create intent bất biến (`VL10D_INTENT_<mat>`) gồm Request ID, Material, Qty định tạo, STO/Item và Timestamp.
  - Xóa sạch pending key và receipt cũ trước mỗi đợt tạo nhóm mới; loại trừ hoàn toàn việc tái sử dụng receipt sai scope.
  - Bỏ `On Error Resume Next` bao quanh verify/commit; thay bằng bắt lỗi có cấu trúc `On Error GoTo Failed`.
  - Loại bỏ hoàn toàn lỗi thay thế số lượng kế hoạch (`obQty <= 0 Then obQty = planned(3)`).
  - Tự động lọc các OB baseline đã tồn tại từ trước trong `EDCDiscoverReceiptsFromVL06O`.
  - Phân loại trạng thái rõ ràng từng mã: `VERIFIED_FULL`, `VERIFIED_PARTIAL`, `VERIFIED_ZERO`, `UNKNOWN_READBACK`.
  - Commit riêng phần đã xác minh (`verifiedPlans.Count > 0`); loại bỏ triệt để lỗi báo `COMMITTED Materials=0`.
- **P1 — Đưa Recovery 3001 Vào Đúng Các Đường Quyết định**:
  - Kích hoạt giải phóng draft 3001 chỉ khi Due hợp lệ và có thâm hụt thực tế đã được xác minh (`deficits.Count > 0`).
  - Khi Due = 0: Chẩn đoán STO mà không tự hủy bừa bãi draft 3001 (`WAIT_STO_DUE_4007`).
  - Thẩm định live ứng viên draft 3001 (đúng Ship-to 3001, đúng Material, Batch/SLoc trống, chưa GI) ngay trước khi xóa.
  - Chờ 2 giây đồng bộ SAP ATP sau khi xóa draft 3001 (`WAIT_RELEASE_SYNC`) trước khi tự động thử lại (`AUTO_RETRY_VL10D`).
- **P2 — Sửa Đọc Classic UoM và Xử lý Triệt để Đơn vị Khác EA**:
  - Bắt buộc log UoM chi tiết: `raw`, `normalized`, `expected`, `material`, `STO/item`, `page`, `row`, tọa độ `x` và `control ID`.
  - Phân biệt rõ `UOM_EMPTY` (ô trống) và `UOM_MISMATCH` (sai lệch đơn vị thực sự).
  - Hàm `NormalizeUoM` loại bỏ ký tự trắng ẩn (`Chr(160)`, `Chr(9)`, `Chr(10)`, `Chr(13)`).
  - Hỗ trợ nhóm đơn vị đếm tương đương chuẩn SAP: `EA`, `UN`, `PC`, `ST`, `PCE` qua hàm `IsCompatibleUoM`.
  - Hỗ trợ định dạng số thập phân trong VL02N cho đơn vị khác EA qua hàm `FormatSAPQty`.
- **P3 — Dominant STO Clustering & Loại bỏ Hoàn toàn Popup**:
  - Loại bỏ hoàn toàn `Application.InputBox` và `MsgBox` trong `EDCPlanOne`.
  - Tự động tính toán và chọn STO chủ đạo (Dominant STO) gom được nhiều vật tư nhất cùng nhau trong cùng đợt chạy để tối đa hóa việc gộp chung vào 1 Outbound Delivery trên VL10D.
  - Cơ chế tie-breaker: Ưu tiên theo thứ tự STO tại R31 và Due Qty lớn hơn.
- **Biên dịch & Bảo toàn Hệ thống**:
  - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
  - Gắn macro nút bấm chuẩn xác (`INPUT!OBV1__0` -> `OB_Check`, `INPUT!OBV1__1` -> `OB_Execute`).
  - Khử bỏ triệt để `forceFullCalc="1"` trong `xl/workbook.xml`.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC50 (13/10/2026)

### 📦 Các tệp đính kèm RC50

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC50_HANDOFF.zip`** | 707,717 bytes | `7CB1088B32B6ADEC280737F49DCC1CB8130FA997F376AD8D6810E4E860F6BC9E` | **Gói bàn giao đầy đủ RC50 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC50.xlsm`** | 753,624 bytes | `545E9C5CAB9C3F0BACC34D53DFF840B3B8C3618853DE72E5760A87D29A7ECAC4` | Tệp Excel Macro phát hành RC50 (Đã dọn dẹp & Compile) |
| **`VBA_RC49_TO_RC50.diff`** | 26,795 bytes | `5832329BBF37C547A1A34798F74042F18A429AEA14A42A2A9AF081807D171307` | Chi tiết bản vá mã nguồn VBA từ RC49 lên RC50 |
| **`CHANGELOG_RC49_TO_RC50.md`** | 6,542 bytes | `17EE73D5A9E56996828D4DB72847970539894D38EFCDCC36678C6176282CE3D7` | Nhật ký thay đổi và giải pháp kỹ thuật RC50 |

---

## 📜 Lịch sử bản phát hành trước đó

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
