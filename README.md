# OB TOOL V1 RC57

Canonical XLSM: `261010_OB_TOOL_V1_RC57.xlsm`  
Canonical SHA256: `FCBB9AB80DD612D0C03CBBD2F503253EAEDF1EC123E06F6DEEFDAD9D69FC1C99`  
Handoff Archive: `RC57_HANDOFF.zip`  
Archive SHA256: `A098F2C8E96713F3FD316E030D5A45CDBEC95800FA0F9C8970A3E98F78287A72`

## Key Improvements in RC57
- **P0 — Cơ chế Checkpoint State Machine (`ModuleMain.bas`)**:
  - Giữ nguyên trạng thái các mã đã tạo thành công (`VERIFIED_FULL`), khóa ghi chú cột D sheet `INPUT`, loại trừ hoàn toàn khỏi danh sách tạo mới của VL10D.
  - Chỉ tính toán số lượng bù cho các mã thiếu (`Deficit > 0`): $Need_{bù} = TotalNeed - CreatedQty$.
  - Khi người dùng sửa dữ liệu (`Fingerprint(False) != saved`), tự động Clean Reset 100%: xóa sạch toàn bộ cấu hình checkpoint cũ (`VL10D_DIRECT_DONE_*`, `VL10D_PENDING_KEY`, `VL10D_GROUP_OB_RECEIPTS`), xóa trắng `INPUT!D6:D205` và `INPUT!F6:N505`, sinh `NEW_REQUEST`.
- **P0 — Item-Level Delete trên VL02N cho Chứng Từ Đa Mã (`ModuleSAP.bas`)**:
  - Sửa đổi hàm `SAPX_DeleteDraftOB`: Khi chứng từ OB có từ 2 item trở lên (`OverviewSnapshot.count >= 2`), tuyệt đối không bấm `btn[14]` xóa cả chứng từ.
  - Định vị chính xác dòng tuyệt đối chứa mã thiếu, chọn dòng và gọi lệnh Menu `Edit -> Delete Item` (`Shift+F2`), xác nhận popup và bấm Save `btn[11]`. Bảo toàn 100% các dòng mã đã đủ.
  - Chỉ xóa cả chứng từ qua Toolbar `btn[14]` khi OB chỉ có đúng 1 item.
- **P0 — Gom Batch Xóa Draft 3001 & Quét Nhả Tồn Chủ Động Khi Due = 0 (`ModuleSAP.bas`, `ModuleMain.bas`)**:
  - Gom toàn bộ danh sách mã thiếu vào 1 lần mở `VL06O` duy nhất qua `SAPX_FindAndReleaseDraft3001OB_Batch`.
  - Nhập danh sách mã vào Multiple Selection của `MATNR`, quét ALV 1 lượt, sang `VL02N` xóa tuần tự và đợi 2s đồng bộ ATP kho 1 lần duy nhất sau khi hoàn tất toàn bộ batch.
  - Kích hoạt chủ động trong `EDCPlanOne`: Nếu phát hiện STO Due = 0 hoặc tổng Due < Need do nhánh 3001 chiếm giữ, tự động tìm và xóa draft 3001 ngay tại bước lập kế hoạch, sau đó làm mới Due list để tiếp tục tạo VL10D.
- **P1 — Tối ưu VL10D - Loại bỏ mở lặp 2 lần & Tận dụng F8 Refresh tại chỗ (`ModuleSAP.bas`)**:
  - Tái sử dụng màn hình Due List từ Lần 1 khi đã mở danh sách với các STO (`canReuseScreen = True`), không thoát ra gõ lại `/nVL10D`.
  - Bổ sung hàm `EDC4007_RefreshDueList` gửi lệnh F8 Refresh (`btn[8]`) trực tiếp trên thanh công cụ của bảng Due List khi cần cập nhật số lượng Due.
- **P1 — Phòng vệ đa tầng chống văng lỗi Runtime & Bảo toàn Pending Key (`ModuleMain.bas`)**:
  - Trong `EDCVerifyGroup`, cơ chế chờ mềm (retry đọc lại VL06O sau 2s) và tuyệt đối không gọi `Err.Raise` làm sập luồng sang khối `Failed:` khi chưa thấy receipt, cho phép nhánh xóa 3001 tiếp tục chạy khép kín vòng lặp.
  - Bảo toàn `VL10D_PENDING_KEY` trong `EDCCommitGroup` khi còn kế hoạch `AUTO_RETRY_VL10D`.
- **Biên dịch & Bảo toàn Hệ thống**:
  - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
  - Gắn macro nút bấm chuẩn xác (`INPUT!OBV1__0` -> `OB_Check`, `INPUT!OBV1__1` -> `OB_Execute`).
  - Khử bỏ triệt để `forceFullCalc="1"` trong `xl/workbook.xml`.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC57 (14/10/2026)

### 📦 Các tệp đính kèm RC57

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC57_HANDOFF.zip`** | 747,492 bytes | `A098F2C8E96713F3FD316E030D5A45CDBEC95800FA0F9C8970A3E98F78287A72` | **Gói bàn giao đầy đủ RC57 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC57.xlsm`** | 794,892 bytes | `FCBB9AB80DD612D0C03CBBD2F503253EAEDF1EC123E06F6DEEFDAD9D69FC1C99` | Tệp Excel Macro phát hành RC57 (Đã tối ưu & Compile) |
| **`VBA_RC56_TO_RC57.diff`** | 35,969 bytes | `6A71333A3C67E57983AB4B0EA3C2C0C98B153B7EE3B86CEAE74F3FCFAFE1AD0E` | Chi tiết bản vá mã nguồn VBA từ RC56 lên RC57 |
| **`CHANGELOG_RC56_TO_RC57.md`** | 9,260 bytes | `3B12E33F735EAD44C577EAA14F577485FC673A6BF4602DB7FA6BBBE63EBE38BD` | Nhật ký thay đổi và giải pháp kỹ thuật RC57 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC56 (14/10/2026)
* **`RC56_HANDOFF.zip`**: SHA-256 `68407EB7CECE7C312DA446C86FAAA8BCF3C474D90C5AC57E94B34C4806A329F5`
* **`261010_OB_TOOL_V1_RC56.xlsm`**: SHA-256 `D1C86081F039C5D74217A5264636E26F17D382C1650815BCB1508BAD55A64763`

### OB Tool V1 RC55 (14/10/2026)
* **`261010_OB_TOOL_V1_RC55.xlsm`**: SHA-256 `547A36809C42263FCA5050C6CA162E089A0C3B77D0A4729FF3A5EEE2300C99AF`
