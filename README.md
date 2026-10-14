# OB TOOL V1 RC58

Canonical XLSM: `261010_OB_TOOL_V1_RC58.xlsm`  
Canonical SHA256: `298C458259268CC602EBEA890840C5877F4F6F397B5157C68E9D675DF04C4973`  
Handoff Archive: `RC58_HANDOFF.zip`  
Archive SHA256: `A706786D3E846F711C457C50B35914C892158239F0CC9D3AA1C60C4EA1DE315E`

## Key Improvements in RC58
- **P0 — Khắc phục Lỗi ngầm 1: Nhận diện trạng thái `DELETE_ITEM_OK` trong Batch Release (`ModuleSAP.bas`)**:
  - Tại dòng 5753 trong `SAPX_FindAndReleaseDraft3001OB_Batch`, chấp nhận cả `DELETE_VERIFIED` và `DELETE_ITEM_OK`.
  - Công nhận giải phóng tồn thành công khi xóa 1 item trên draft OB đa mã, tăng biến đếm `countSuccess` và kích hoạt chờ đồng bộ ATP chính xác.
- **P0 — Khắc phục Lỗi ngầm 2: Truyền `targetMat` và bảo vệ chứng từ đa mã trong Single Release (`ModuleSAP.bas`)**:
  - Tại dòng 5896 trong `SAPX_FindAndReleaseDraft3001OB`, truyền đầy đủ `targetMat` vào hàm `SAPX_DeleteDraftOB(ses, bestOB, targetMat)`.
  - Chấp nhận cả `DELETE_VERIFIED` và `DELETE_ITEM_OK`, ngăn chặn hoàn toàn rủi ro bấm `btn[14]` xóa cả chứng từ đa mã.
- **P0 — Khắc phục Lỗi ngầm 3: Xóa cấp dòng item dở dang trong `PARTIAL_STO_CLEANUP` (`ModuleMain.bas`)**:
  - Khi OB chứa mã đã hoàn thành 100% (`obHasCompletedMat = True`), kích hoạt `SAPX_DeleteDraftOB(pSes, pTargetOB, CStr(pCleanMat))` để xóa riêng dòng thiếu.
  - Giải phóng Due dở dang về STO để tạo bù thành công, đồng thời bảo toàn tuyệt đối các dòng hàng đã đủ số lượng.
- **P1 — Tối ưu 1: Tự động đóng Pop-up cảnh báo sau khi bấm Lưu trên VL02N (`ModuleSAP.bas`)**:
  - Sau lệnh Save `btn[11]` trong Item Delete, tự động xác nhận pop-up thông tin/cảnh báo `wnd[1]`, chống treo phiên SAP.
- **P1 — Tối ưu 2: Short-circuit thoát sớm trong `EDCReleaseDrafts3001` (`ModuleMain.bas`)**:
  - Bỏ qua mở VL06O quét tồn khi tổng số lượng thiếu $\le 0$.
- **P1 — Tối ưu 3: Chuẩn hóa ghi chú trạng thái cột D sheet `INPUT` (`ModuleMain.bas`)**:
  - Phân biệt rõ ràng giữa xóa dòng item dở dang vs xóa toàn bộ chứng từ OB.
- **Biên dịch & Bảo toàn Hệ thống**:
  * Native VBA Debug Compile PASS 100% (`before=True`, `after=False`).
  * Gắn macro nút bấm chuẩn xác (`INPUT!OBV1__0` -> `OB_Check`, `INPUT!OBV1__1` -> `OB_Execute`).
  * Khử bỏ triệt để `forceFullCalc="1"` trong `xl/workbook.xml`.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC58 (14/10/2026)

### 📦 Các tệp đính kèm RC58

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC58_HANDOFF.zip`** | 887,203 bytes | `A706786D3E846F711C457C50B35914C892158239F0CC9D3AA1C60C4EA1DE315E` | **Gói bàn giao đầy đủ RC58 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC58.xlsm`** | 906,844 bytes | `298C458259268CC602EBEA890840C5877F4F6F397B5157C68E9D675DF04C4973` | Tệp Excel Macro phát hành RC58 (Đã tối ưu & Compile) |
| **`VBA_RC57_TO_RC58.diff`** | 6,565 bytes | `09F28D316DBE29A5918CA8374044E7524118A4201D3063797C21D90D06021402` | Chi tiết bản vá mã nguồn VBA từ RC57 lên RC58 |
| **`CHANGELOG_RC57_TO_RC58.md`** | 6,315 bytes | `6E2C61B277BE1EB58C981C48FCC163A4285CE665D2BDD0306564B7F9FECDF845` | Nhật ký thay đổi và giải pháp kỹ thuật RC58 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC57 (14/10/2026)
* **`RC57_HANDOFF.zip`**: SHA-256 `E1D019013F7EC27D861D02C1F9A7ED739F57908B17135C5BDAC01A4EBA2EBB3B`
* **`261010_OB_TOOL_V1_RC57.xlsm`**: SHA-256 `16FE33682976C325B6AF97E59489BD5198C4BA862F8A4FF957D71C1E699B5C98`

### OB Tool V1 RC56 (14/10/2026)
* **`RC56_HANDOFF.zip`**: SHA-256 `68407EB7CECE7C312DA446C86FAAA8BCF3C474D90C5AC57E94B34C4806A329F5`
* **`261010_OB_TOOL_V1_RC56.xlsm`**: SHA-256 `D1C86081F039C5D74217A5264636E26F17D382C1650815BCB1508BAD55A64763`
