# Production & Logistics Automation Tools Suite

Kho lưu trữ tập trung các công cụ tự động hóa, giám sát và điều hành sản xuất, tồn kho xưởng và tích hợp SAP ERP.

---

## 📂 Cấu trúc danh mục công cụ (`tools/`)

```text
tools/
├── ob-tool/              # Công cụ tự động hóa tạo & xử lý Outbound Delivery (OB / DO) trên SAP ERP
├── stock-fg-ma-body/     # Công cụ cảnh báo tồn kho thành phẩm cho Xưởng BODY & Xe Máy Điện (XMĐ)
├── ma-control-fg/        # Hệ thống phần mềm điều hành lắp ráp và logistics MA Control FG (Portable & Web)
└── utilities/            # Các công cụ tiện ích phụ trợ (Biên bản bàn giao BBBG, Cân điện...)
```

---

## 🚀 Truy cập nhanh các bản phát hành mới nhất

### 1. 📊 [Stock FG MA BODY V3.7](./tools/stock-fg-ma-body/) (Cập nhật 16/10/2026)
*Tách biệt hoàn toàn Shop BODY và XMĐ, hiển thị rõ ràng mã Trước/Sau sơn, sửa cảnh báo "Ưu tiên kéo line", font Times New Roman & kẻ ô toàn diện.*
* **Tệp Excel**: [`tools/stock-fg-ma-body/261008_Stock_FG_MA_BODY_V3_7.xlsm`](./tools/stock-fg-ma-body/261008_Stock_FG_MA_BODY_V3_7.xlsm) — SHA-256: `84C5F2FFF853E119052CAFC3CDC1609D713C5B4BBF052A1C84F752AE077CD07F`
* **Gói nén ZIP**: [`tools/stock-fg-ma-body/261008_Stock_FG_MA_BODY_V3_7.zip`](./tools/stock-fg-ma-body/261008_Stock_FG_MA_BODY_V3_7.zip) — SHA-256: `7DA6952E3DF9EEDF6665405C531758A1D42B735476A2332706CDCDEE937F0D83`
* **Ghi chú phát hành**: [`tools/stock-fg-ma-body/RELEASE_NOTES_V3_7.md`](./tools/stock-fg-ma-body/RELEASE_NOTES_V3_7.md)

---

### 2. ⚡ [OB Tool V1 RC64](./tools/ob-tool/) (Cập nhật 16/10/2026)
*Tự động hóa VL10D STO, giải phóng OB nháp kho 3001, phân bổ Batch VL02N và nút bấm điều phối ALL IN ONE.*
* **Tệp Excel**: [`tools/ob-tool/261010_OB_TOOL_V1_RC64.xlsm`](./tools/ob-tool/261010_OB_TOOL_V1_RC64.xlsm) — SHA-256: `9A9D9A3875D037C5DF5771932C1622C993780AE41131D52B6ADA90A8FC199FC5`
* **Gói bàn giao Handoff**: [`tools/ob-tool/RC64_HANDOFF.zip`](./tools/ob-tool/RC64_HANDOFF.zip) — SHA-256: `80DDDC1D787B7058A4F530285080E316C853D15E5FB5F67472C9332A1D86DA9F`
* **Vá mã VBA & Changelog**: [`tools/ob-tool/CHANGELOG_RC63_TO_RC64.md`](./tools/ob-tool/CHANGELOG_RC63_TO_RC64.md)

---

### 3. 🖥️ [MA Control FG Suite](./tools/ma-control-fg/)
*Nền tảng kiểm soát lắp ráp và dữ liệu điều hành xưởng (Zero-Port Architecture).*
* **Bản Portable v1.1.0**: [`tools/ma-control-fg/MA_CONTROL_FG_PORTABLE_ZERO_PORT_v1.1.0.zip`](./tools/ma-control-fg/MA_CONTROL_FG_PORTABLE_ZERO_PORT_v1.1.0.zip)
* **Gói Handoff**: [`tools/ma-control-fg/MA_CONTROL_FG_V1_0_0_HANDOFF.zip`](./tools/ma-control-fg/MA_CONTROL_FG_V1_0_0_HANDOFF.zip)

---

### 4. 🧰 [Tiện ích phụ trợ](./tools/utilities/)
* **Biên bản bàn giao**: [`tools/utilities/261008_BBBG_FG_V2_2_14.zip`](./tools/utilities/261008_BBBG_FG_V2_2_14.zip)
* **Cân điện**: [`tools/utilities/candienmoi.zip`](./tools/utilities/candienmoi.zip)

---

## 🔒 Quản lý quyền truy cập & Phân quyền tài khoản

Repository hỗ trợ phân quyền chi tiết cho từng tài khoản GitHub:
1. Khi repository ở chế độ **Private**, chỉ những tài khoản được mời trong danh sách **Collaborators** mới có quyền xem và tải xuống các tệp.
2. Để cấp quyền cho một tài khoản mới:
   - Vào **Settings** -> **Collaborators** -> Nhấn **Add people**.
   - Tìm theo username hoặc email GitHub của nhân sự cần cấp quyền.
   - Chọn vai trò **Read** (chỉ xem & tải về) hoặc **Write** (được đẩy code/cập nhật).
