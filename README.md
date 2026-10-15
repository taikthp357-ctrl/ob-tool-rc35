# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC64 (16/10/2026)

### 📦 Các tệp đính kèm RC64

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC64_HANDOFF.zip`** | 935,918 bytes | `80DDDC1D787B7058A4F530285080E316C853D15E5FB5F67472C9332A1D86DA9F` | **Gói bàn giao đầy đủ RC64 (Streamlined VL10D & ALL IN ONE)** |
| **`261010_OB_TOOL_V1_RC64.xlsm`** | 949,326 bytes | `9A9D9A3875D037C5DF5771932C1622C993780AE41131D52B6ADA90A8FC199FC5` | Tệp Excel Macro phát hành RC64 (Đã tối ưu & Compile PASS 100%) |
| **`VBA_RC63_TO_RC64.diff`** | 10,677 bytes | `679885A193EE90D303B0725D20F0A046D6C8A4F844B79116342110A6619D427A` | Chi tiết bản vá mã nguồn VBA từ RC63 lên RC64 |
| **`CHANGELOG_RC63_TO_RC64.md`** | 4,323 bytes | `7EE3DC060A907B4F5DDF09C6F74059617689981C9B4C1D437A4A3BFB599CDD5E` | Nhật ký thay đổi và giải pháp kỹ thuật RC64 |

### 🚀 Điểm cải tiến chính RC64
1. **Kiểm tra mức giảm STO Due trực tiếp trên VL10D (F8 Refresh `btn[8]`)**:
   - Sau khi bấm `btn[19]`, bấm ngay `btn[8]` để làm mới Due list.
   - Nếu số lượng STO đã giảm đủ: **COI LÀ DONE**, bỏ qua hoàn toàn việc quét/xóa 3001, chuyển thẳng sang VL02N.
   - Chỉ khi STO giữ nguyên chưa giảm số lượng, mới rẽ sang 3001 xóa Draft OB nhả tồn.
2. **Khắc phục triệt để hiển thị nút ALL IN ONE**:
   - Tạo đủ 3 nút bấm độc lập: Nút 1 (Teal), Nút 2 (Cam), Nút 3 ALL IN ONE (Royal Blue `RGB(30, 64, 175)`) tại `P17:R19`.
   - Đảm bảo hiển thị tự động khi mở workbook qua `Workbook_Open` và tự phục hồi trong `OB_Check`.

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC63 (15/10/2026)
* **`RC63_HANDOFF.zip`**: SHA-256 `B69601E855CD1DECE016C0D597A2E9F6DFE73E29EB334D972DF79FDBC4AE5980`
* **`261010_OB_TOOL_V1_RC63.xlsm`**: SHA-256 `60EF66BDD0E44BA13E1A635CD6EB255749A5F8990F605A7FFB5254A112108D64`

### OB Tool V1 RC62 (15/10/2026)
* **`RC62_HANDOFF.zip`**: SHA-256 `DB39C4BD0CD87E48DA41213DF0E305BD2B90010CD30BD3295EDAA82681C1C00D`
* **`261010_OB_TOOL_V1_RC62.xlsm`**: SHA-256 `4947174A06C0D0017BD17924803C14699B6305F6A1253C9ADECD66E0A97C6593`
