# OB TOOL V1 RC62

Canonical XLSM: `261010_OB_TOOL_V1_RC62.xlsm`  
Canonical SHA256: `4947174A06C0D0017BD17924803C14699B6305F6A1253C9ADECD66E0A97C6593`  
Handoff Archive: `RC62_HANDOFF.zip`  
Archive SHA256: `DB39C4BD0CD87E48DA41213DF0E305BD2B90010CD30BD3295EDAA82681C1C00D`

## Key Improvements in RC62
- **P0 — Thuật toán chọn tổ hợp OB tối ưu theo Qty (Optimal Subset-Sum)**:
  * Ưu tiên tìm tổ hợp OB giải phóng vừa đủ nhu cầu (surplus = 0); nếu không có tổ hợp vừa đủ thì chọn tổ hợp có tổng Qty vượt nhu cầu ít nhất (surplus > 0 tối thiểu).
  * Cho phép xóa thừa khi không có lựa chọn vừa khít (không chặn cứng vô lý), ví dụ cần 16 PCS nhưng chỉ có OB 32 PCS thì xóa nguyên OB 32 PCS.
  * Thuật toán nhánh và cận (Branch-and-Bound với Suffix-Sums) đảm bảo thời gian tính toán < 1ms, độ chính xác tuyệt đối.
- **P0 — Tối ưu theo từng mã vật tư riêng biệt (Individual Material Optimization)**:
  * Lần lượt lập kế hoạch giải phóng cho từng mã thiếu trong danh sách `deficits`, chỉ chọn đúng số OB cần thiết cho từng mã rồi dừng.
  * Khắc phục triệt để lỗi của RC61 khi gom xóa hàng loạt tất cả các OB cùng vật tư mà không kiểm soát số lượng.
- **P0 — Tính toán toàn bộ ảnh hưởng đa mã trên chứng từ (Multi-Material OB Impact)**:
  * Khi một OB chứa nhiều mã được chọn để giải phóng, hệ thống tự động nhận diện và khấu trừ số lượng giải phóng cho toàn bộ các mã linh kiện có trong OB đó.
  * Ngăn ngừa xóa trùng OB hoặc xóa thừa OB cho các mã đi kèm đã được giải phóng đủ.
- **P0 — Giữ nguyên cơ chế xóa nhanh toàn bộ OB (`isDraft3001Release=True`)**:
  * Giữ nguyên thao tác xóa nhanh qua nút `btn[14]` trên thanh công cụ VL02N và xác nhận popup.
  * Không duyệt từng dòng item, không giảm Qty, xác nhận trực tiếp qua thông điệp trạng thái của SAP.
- **P0 — Bảo toàn cơ chế phòng vệ xác minh kết quả VL10D**:
  * Kiểm tra đầy đủ kết quả tạo OB từ VL10D trước khi quyết định giải phóng 3001, loại bỏ nguy cơ xóa nhầm khi SAP phản hồi chậm.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC62 (15/10/2026)

### 📦 Các tệp đính kèm RC62

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC62_HANDOFF.zip`** | 906,995 bytes | `DB39C4BD0CD87E48DA41213DF0E305BD2B90010CD30BD3295EDAA82681C1C00D` | **Gói bàn giao đầy đủ RC62 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC62.xlsm`** | 921,710 bytes | `4947174A06C0D0017BD17924803C14699B6305F6A1253C9ADECD66E0A97C6593` | Tệp Excel Macro phát hành RC62 (Đã tối ưu & Compile) |
| **`VBA_RC61_TO_RC62.diff`** | 16,922 bytes | `19EF8E7CA908D5A1FDBC3FF140DB848D02B6D4C63FBC73656C6CD553BE60F9D8` | Chi tiết bản vá mã nguồn VBA từ RC61 lên RC62 |
| **`CHANGELOG_RC61_TO_RC62.md`** | 4,215 bytes | `6B6B5CE9B7A277BC3D31AC9908F81711FA8E3D7B1B6195ED1136A4143BCFDE5E` | Nhật ký thay đổi và giải pháp kỹ thuật RC62 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC61 (15/10/2026)
* **`RC61_HANDOFF.zip`**: SHA-256 `0DB0D4437D316843DC3DE7CFA517DD2ADE8432631B0E98CD78504A13519180B0`
* **`261010_OB_TOOL_V1_RC61.xlsm`**: SHA-256 `E446C065F22FD3D25CD14318AA54CB66C7599C7E6637650DE2C2A974C1DFDF5D`

### OB Tool V1 RC60 (14/10/2026)
* **`RC60_HANDOFF.zip`**: SHA-256 `8EA43D6F7A4E835D5DD6EE8A122321762EC9777CF3D35DA52E0E587B9DBA2A15`
* **`261010_OB_TOOL_V1_RC60.xlsm`**: SHA-256 `FF86BF33B8CB638755E479B7071C680E36639A25EB840688ABB6E72CD92A83FF`
