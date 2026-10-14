# CHANGELOG - OB TOOL V1 RC62

## 1. Tổng quan Bản phát hành RC62 (Thuật toán Chọn OB Tối Ưu Theo Qty, Bù Vừa Đủ & Tính Toàn Bộ Ảnh Hưởng Đa Mã)
- **Vấn đề cốt lõi trong RC61**:
  Ở RC60 và RC61, nhánh giải phóng Draft OB 3001 (`SAPX_FindAndReleaseDraft3001OB_Batch`) gom toàn bộ các OB nháp không có SLoc của các mã thiếu rồi xóa sạch toàn bộ hàng loạt (đã từng xóa tới 5 OB mã thứ nhất và 11 OB mã thứ hai). Điều này làm xóa vượt quá mức cần thiết, lãng phí chứng từ nháp và làm xáo trộn kế hoạch STO.
- **Mục tiêu giải pháp kỹ thuật RC62**:
  1. Giữ nguyên 100% cơ chế xóa nhanh toàn bộ OB qua `btn[14]` (isDraft3001Release=True) không mở từng dòng item kiểm tra, không giảm Qty.
  2. Bổ sung thuật toán lựa chọn tổ hợp OB tối ưu theo Qty (Subset-Sum): Ưu tiên xóa vừa đủ (surplus = 0); nếu không có tổ hợp vừa đủ thì chọn tổ hợp có tổng Qty vượt nhu cầu ít nhất (surplus > 0 tối thiểu).
  3. Tối ưu theo từng mã riêng biệt: Không gom xóa hàng loạt, chỉ chọn đúng số OB cần giải phóng cho từng mã rồi dừng.
  4. Tính toán toàn bộ ảnh hưởng đa mã: Nếu OB chứa nhiều mã, khi xóa sẽ đồng thời giải phóng tất cả các mã đó. Thuật toán tự động trừ số lượng được giải phóng của các mã còn lại, không tính trùng OB hai lần hoặc xóa thừa OB của mã đi kèm.
  5. Bảo toàn cơ chế xác minh kết quả VL10D trước khi quyết định giải phóng, tránh trường hợp SAP tạo chậm khiến tool xóa nhầm.

---

## 2. Bảng Phép Thử Kiểm Chứng Thuật Toán Lựa Chọn OB (Subset-Sum)

| Nhu cầu cần giải phóng | Danh sách OB hiện có | Tổ hợp OB được chọn | Tổng Qty giải phóng | Đánh giá / Trạng thái |
| :---: | :---: | :---: | :---: | :---: |
| **16** | 16, 32 | **16** | 16 | Khớp chính xác (surplus = 0) |
| **16** | 10, 12 | **10 + 12** | 22 | Vượt tối thiểu (+6 PCS) |
| **16** | 32 | **32** | 32 | Cho phép xóa thừa (+16 PCS), không chặn cứng |
| **16** | 8, 16, 32 | **16** | 16 | Ưu tiên OB đơn lẻ khớp vừa đủ |
| **32** | 16, 16, 80 | **16 + 16** | 32 | Tổ hợp 2 OB vừa đủ 32 |
| **32** | 20, 32, 50 | **32** | 32 | Khớp chính xác 1 OB 32 |
| **32** | 16, 40, 80 | **40** | 40 | Vượt tối thiểu (+8 PCS) thay vì chọn 80 (+48 PCS) |

---

## 3. Chi tiết Kỹ thuật Các Thay Đổi Cốt Lõi (`ModuleSAP.bas`)

### 3.1. Động cơ Tìm Tổ Hợp Tối Ưu (`OptimalOBSubset` & `BacktrackSubset`)
- Sử dụng thuật toán đệ quy nhánh và cận (Branch-and-Bound với Suffix-Sums Pruning).
- Sắp xếp các ứng viên giảm dần theo số lượng để cắt tỉa nhánh siêu tốc (thời gian tính toán < 1ms).
- Hàm mục tiêu 2 tầng:
  * Tầng 1: $\min(	ext{Surplus}) = \min(\sum Q - Q_{	ext{need}}) \ge 0$.
  * Tầng 2: $\min(	ext{Count})$ (Ưu tiên số lượng OB ít nhất khi thặng dư bằng nhau).
- Trường hợp tổng tồn khả dụng của tất cả các OB $< Q_{	ext{need}}$, tự động chọn toàn bộ các OB hiện có để giải phóng tối đa có thể.

### 3.2. Tối Ưu Hóa Riêng Theo Từng Mã & Trừ Dồn Ảnh Hưởng Đa Mã
- Quét ALV Grid VL06O thu thập cấu trúc chứng từ hoàn chỉnh:
  * `obItemsByOB(ob)`: Lưu từ điển các mã vật tư và số lượng tương ứng trong từng OB.
  * `obMatMap(mat)`: Danh sách các OB có chứa vật tư.
- Lần lượt xét từng mã vật tư thiếu trong `deficits`:
  * Nếu `remDeficits(mat) <= 0`: Đã được giải phóng gián tiếp từ trước $ightarrow$ Bỏ qua.
  * Lọc các OB chứa mã chưa bị đưa vào `finalDeleteOBs`.
  * Chạy `OptimalOBSubset` chọn đúng tổ hợp tối ưu cho mã đó.
  * **Cập nhật ảnh hưởng đa mã**: Với mỗi OB được chọn, duyệt toàn bộ các mã linh kiện khác có trong OB đó và tự động trừ số lượng tương ứng vào `remDeficits`.

### 3.3. Giữ Nguyên Cơ Chế Xóa Nhanh Toàn Bộ OB (`isDraft3001Release=True`)
- Trong `SAPX_DeleteDraftOB`, giữ nguyên luồng `isDraft3001Release=True` chuyển thẳng tới `DeleteWholeOB:`.
- Bấm nút xóa trên thanh công cụ `btn[14]`, xác nhận popup qua `btnOption1` (`btnSPOP-OPTION1` / `btn[11]`).
- Đọc kết quả trực tiếp từ SAP Status Bar (`deleted`, `gelscht`, `da xoa`, Type="S") mà không cần mở lại chứng từ hay duyệt từng item.

---

## 4. Bằng chứng Kiểm chứng Kỹ thuật
- **Native VBE Debug Compile**: PASS 100% (`before=True`, `after=False`).
- **Shape Macro Bindings**:
  * `INPUT!OBV1__0` -> `'261010_OB_TOOL_V1_RC62.xlsm'!OB_Check`
  * `INPUT!OBV1__1` -> `'261010_OB_TOOL_V1_RC62.xlsm'!OB_Execute`
- **Version Labels**:
  * `INPUT!A1` = `"OB TOOL V1 RC62"`
  * `CONFIG!B2` = `"RC62"`
  * `INPUT!R22` = `"RC62 OPTIMAL OB SUBSET SELECTION & MULTI-MAT ENGINE"`
  * `INPUT!R25` = `"62.0"`
- **CalcPr Optimization**: Thuộc tính `forceFullCalc="1"` tuyệt đối không tồn tại trong XML.
