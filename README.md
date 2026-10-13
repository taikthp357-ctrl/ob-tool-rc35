# OB TOOL V1 RC56

Canonical XLSM: `261010_OB_TOOL_V1_RC56.xlsm`  
Canonical SHA256: `4B09DCCA751E62D7E7E57A6638E79F4DC744AD80A47A815E0F9658735F64EBE2`  
Handoff Archive: `RC56_HANDOFF.zip`  
Archive SHA256: `3A29952001E62B7B277FF07C85868FC96E693ACFE01F9719E2AE1121769585FE`

## Key Improvements in RC56
- **P0 — Vá lỗi Schema Deserializer `VL10D_PENDING_KEY` 6 trường (`ModuleMain.bas`)**:
  - Khi tạo OB theo nhóm, lưu 6 trường (`Request | Material | Qty | STO | Item | Timestamp`).
  - Sửa `EDCResolvePendingGroup` và `EDCResolvePendingCreate` chấp nhận `UBound(parts) >= 4` thay vì bắt buộc `= 4`.
  - Không ném lỗi `Invalid pending item identity` và không bao giờ xóa nhầm pending key trên lỗi đọc tạm thời.
- **P0 — Bảo toàn Pending Key trong quy trình `AUTO_RETRY_VL10D` (`ModuleMain.bas`)**:
  - Không xóa pending key trong `EDCCommitGroup` khi vẫn còn deficits cần tạo bù.
  - Thiết lập lại đầy đủ `VL10D_GROUP_PENDING` và `VL10D_PENDING_KEY` trước khi chạy Background Create cho `retryPlans`.
  - Thêm cơ chế fallback xác định STO/Request trong `EDCVerifyGroup`, loại bỏ vĩnh viễn lỗi `GROUP_RECEIPT_MISSING: pending identity invalid`.
- **P0 — Bổ sung Nhánh Xử lý Tạo 1 Nửa STO & Tự Động Xóa OB Thiếu (`ModuleMain.bas`)**:
  - Khi VL10D chỉ tạo được 1 phần số lượng STO và nhánh 3001 không giữ tồn hoặc không thể sinh thêm để đạt nguyên batch:
  - Tự động gọi `SAPX_DeleteDraftOB` per `deleteob.vbs` để xóa bỏ OB dở dang ("xóa tạo bừa đó"), khôi phục lại Due về STO.
  - Loại bỏ mã khỏi danh sách `verifiedPlans`, ghi log `VL10D_PARTIAL_OB_CLEANUP` và cập nhật ghi chú rõ ràng trên sheet INPUT.
- **P1 — Gỡ bỏ Bẫy `PRESAVE_GATE_FAIL: EXTRA_ITEM` trong VL02N (`ModuleSAP.bas`)**:
  - Khi OB có nhiều line item từ STO (ví dụ items 10, 20, 30, 40) mà lượt chạy chỉ phân bổ cho 10, 30, 40 (item 20 chưa có kế hoạch):
  - Chuyển `EXTRA_ITEM` sang log thông tin `PRESAVE_EXTRA_ITEM_INFO`, không chặn đứng việc lưu chứng từ; kiểm tra nghiêm ngặt 100% tất cả các item có trong kế hoạch `itemMap`.
  - Tương tự trong `SavedSnapshotMatchesPlan`, cho phép lưu khi tồn tại item chưa có kế hoạch trên overview.
- **P1 — Khắc phục Lỗi Đọc `TableField` trong VL02N (`ModuleSAP.bas`)**:
  - Trả về `Nothing` an toàn thay vì ném lỗi `Thieu field ... row 3. Can MAP layout VL02N.`.
  - `OverviewSnapshot` kết thúc vòng lặp an toàn khi gặp `Nothing`.
- **P1 — Đối Chiếu An Toàn Draft 3001 & Mở Rộng Phạm Vi Ngày Planned GI (`ModuleSAP.bas`, `ModuleMain.bas`)**:
  - Đối soát chính xác theo người tạo (`currentUser` ERNAM) và STO tham chiếu (`targetSTO` VGBEL). Ưu tiên tuyệt đối draft do chính user tạo ra.
  - Mở rộng `IT_WADAT-HIGH` từ `Date` thành `DateAdd("d", 60, Date)` để bao quát cả các OB có Planned GI trong tương lai.
  - `EDCPlanOne`: Ưu tiên chọn STO có Due đáp ứng trọn vẹn nguyên batch trước khi chọn theo STO chủ đạo. Không bỏ dở lập kế hoạch khi `missing > 0`.
- **Biên dịch & Bảo toàn Hệ thống**:
  - Native VBA Debug Compile PASS 100% (`compile_enabled_before = True`, `compile_enabled_after = False`).
  - Gắn macro nút bấm chuẩn xác (`INPUT!OBV1__0` -> `OB_Check`, `INPUT!OBV1__1` -> `OB_Execute`).
  - Khử bỏ triệt để `forceFullCalc="1"` trong `xl/workbook.xml`.

---

# OB Tool Releases — Public Release & Handoff Repository

Kho lưu trữ công khai các bản phát hành và gói bàn giao Excel Macro cho công cụ **OB Tool (SAP / Excel)**.

---

## 🌟 Bản phát hành mới nhất: OB Tool V1 RC56 (14/10/2026)

### 📦 Các tệp đính kèm RC56

| Tên tệp | Kích thước | SHA-256 Hash | Mô tả |
| :--- | :--- | :--- | :--- |
| **`RC56_HANDOFF.zip`** | 849,458 bytes | `3A29952001E62B7B277FF07C85868FC96E693ACFE01F9719E2AE1121769585FE` | **Gói bàn giao đầy đủ RC56 (All-in-one)** |
| **`261010_OB_TOOL_V1_RC56.xlsm`** | 793,292 bytes | `4B09DCCA751E62D7E7E57A6638E79F4DC744AD80A47A815E0F9658735F64EBE2` | Tệp Excel Macro phát hành RC56 (Đã tối ưu & Compile) |
| **`VBA_RC55_TO_RC56.diff`** | 654,054 bytes | `DAB5C0117FA48330DDA8740D535595992A3EC94779AD6B960C0FAED7D522040A` | Chi tiết bản vá mã nguồn VBA từ RC55 lên RC56 |
| **`CHANGELOG_RC55_TO_RC56.md`** | 7,526 bytes | `7A746BCF0BCE4D651A469D4CC62F961D0875C10EE26FBF99DC1B456C7E2C82C8` | Nhật ký thay đổi và giải pháp kỹ thuật RC56 |

---

## 📜 Lịch sử bản phát hành trước đó

### OB Tool V1 RC52 (13/10/2026)
* **`RC52_HANDOFF.zip`**: SHA-256 `60155252EB31DC512FBD58804ACF827CFE3F40E7D2E6872BC6779AE613FBCF0F`
* **`261010_OB_TOOL_V1_RC52.xlsm`**: SHA-256 `87F072D409DFD86144DE91E3477142426992D252CEEF8EE99AC6E18B3DC14885`

### OB Tool V1 RC51 (13/10/2026)
* **`RC51_HANDOFF.zip`**: SHA-256 `DB86E5724A2153D43D50606FB0E456C3D9F101F6DBABFCCEC917DBB7AC9F9B50`
* **`261010_OB_TOOL_V1_RC51.xlsm`**: SHA-256 `85966347A480D4AF990881B481A6F860C3950DDEFF6991653BBA0A1615CFCBD9`

### OB Tool V1 RC50 (13/10/2026)
* **`RC50_HANDOFF.zip`**: SHA-256 `7CB1088B32B6ADEC280737F49DCC1CB8130FA997F376AD8D6810E4E860F6BC9E`
* **`261010_OB_TOOL_V1_RC50.xlsm`**: SHA-256 `545E9C5CAB9C3F0BACC34D53DFF840B3B8C3618853DE72E5760A87D29A7ECAC4`
