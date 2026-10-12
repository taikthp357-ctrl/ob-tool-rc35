# OB TOOL V1 RC52

Canonical XLSM: `261010_OB_TOOL_V1_RC52.xlsm`  
Canonical SHA256: `C506832470E005CFF848A2F9319FA9266DC3136B0F5D2683053DFD1025B99979`  
Handoff Archive: `RC52_HANDOFF.zip`  
Archive SHA256: `FE8D927A5DA88E13122159DCA2FCDFB99CFF60F5BB1EDA3789C6460449ED28E8`

## Key Improvements in RC52
- **P0 — Triển khai hàm xóa Draft 3001 theo đúng `deleteob.vbs` (`ModuleSAP.bas`)**:
  - Mở `VL02N`, điền số OB và gõ `Enter` (`sendVKey 0`).
  - Thao tác trực tiếp theo đúng bản ghi: bấm nút Xóa `wnd[0]/tbar[1]/btn[14]`, nhấn xác nhận popup `wnd[1]/usr/btnSPOP-OPTION1`.
  - Hậu kiểm live: Mở lại tra số OB trong `VL02N` để xác nhận thông báo `does not exist`/không còn tồn tại trên SAP trước khi công nhận xóa thành công.
  - Phân định rõ ràng các trạng thái trả về: `DELETE_VERIFIED`, `WAIT_OB_LOCK`, `WAIT_STO_LOCK`, `DELETE_FAILED`.
- **P0 — Khôi phục chạy tiếp khi OB sinh thiếu đã bị xóa (`ModuleMain.bas`)**:
  - Khi phát hiện receipt tồn tại trong `VL10D_DIRECT_DONE_<mat>` nhưng `capacity = 0`: Xác nhận trạng thái `OB_DELETED_CONFIRMED`.
  - Vô hiệu hóa receipt cũ (`VL10D_DIRECT_DONE_<mat> = ""`), ghi log `OB_DELETED_REPLAN`.
  - Không ném lỗi `OB_ALREADY_CREATED_WAIT` nữa; lập lại kế hoạch cho toàn bộ số lượng thực tế cần tạo theo nguồn live (`missing = need`, ví dụ cần 200 tạo 102 bị xóa thì lập lại 200 pcs).
- **P0 — Tách biệt lỗi Khóa STO (`WAIT_STO_LOCK`)**:
  - Nhận diện chính xác thông báo `Purchase order ... is currently being processed by ...`.
  - Định danh trạng thái `WAIT_STO_LOCK`, tuyệt đối không gọi xóa draft 3001 vì đây là lỗi do người khác mở, không phải do thiếu tồn ATP.
- **P1 — Tối ưu hóa tốc độ chuẩn bị VL10D trước Background (`ModuleSAP.bas`)**:
  - Staging theo trang trong `EDCStageGroup`: Gom nhóm các vật tư cần thao tác theo từng trang hiển thị, chỉ cuộn tới mỗi trang một lần, điền số lượng và tick chọn toàn bộ các dòng trên trang đó rồi mới chuyển tiếp.
  - Loại bỏ lệnh gọi `EDCStageGroup False` lặp lại dư thừa ngay trước khi bấm Create Delivery trong `EDC4007_CreateGroup`.
  - Bổ sung log đo thời gian thực tế `VL10D_TIMING`.
- **P3 — Loại bỏ hoàn toàn Popup trong `EDCExistingCapacity` (`ModuleMain.bas`)**:
  - Loại bỏ hoàn toàn các hộp thoại `Application.InputBox` và `MsgBox` khi phát hiện vật tư có OB thuộc nhiều STO khác nhau; tự động ưu tiên STO chủ đạo (`mEDCDominantSTO`) hoặc STO có sức chứa lớn nhất.
- **Biên dịch & Bảo toàn Hệ thống**:
  - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
  - Gắn macro nút bấm chuẩn xác (`INPUT!OBV1__0` -> `OB_Check`, `INPUT!OBV1__1` -> `OB_Execute`).
  - Khử bỏ triệt để `forceFullCalc="1"` trong `xl/workbook.xml`.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC52 (13/10/2026)

### 📦 Các tệp đính kèm RC52

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC52_HANDOFF.zip`** | 705,229 bytes | `FE8D927A5DA88E13122159DCA2FCDFB99CFF60F5BB1EDA3789C6460449ED28E8` | **Gói bàn giao đầy đủ RC52 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC52.xlsm`** | 752,935 bytes | `C506832470E005CFF848A2F9319FA9266DC3136B0F5D2683053DFD1025B99979` | Tệp Excel Macro phát hành RC52 (Đã tối ưu & Compile) |
| **`VBA_RC51_TO_RC52.diff`** | 19,745 bytes | `830B5C7A1AD3C3311365A2FA4CBA5B455C5CA387028B6773046C84B9DAC9DB5A` | Chi tiết bản vá mã nguồn VBA từ RC51 lên RC52 |
| **`CHANGELOG_RC51_TO_RC52.md`** | 5,375 bytes | `B6482C6DAD08CE1EC515634EBBED1EE26C9636C91C0631B50D6A262F167DFAAB` | Nhật ký thay đổi và giải pháp kỹ thuật RC52 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC51 (13/10/2026)
* **`RC51_HANDOFF.zip`**: SHA-256 `DB86E5724A2153D43D50606FB0E456C3D9F101F6DBABFCCEC917DBB7AC9F9B50`
* **`261010_OB_TOOL_V1_RC51.xlsm`**: SHA-256 `85966347A480D4AF990881B481A6F860C3950DDEFF6991653BBA0A1615CFCBD9`

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
