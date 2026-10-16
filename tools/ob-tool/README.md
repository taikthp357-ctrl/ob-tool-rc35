# OB Tool — SAP Outbound Delivery Automation Suite

Bộ công cụ tự động hóa quy trình tạo, xử lý và phân bổ Outbound Delivery (OB / DO) trên SAP ERP qua giao diện Excel Macro (VBA).

---

## 📌 Các phiên bản phát hành

### 🌟 Bản mới nhất: RC65 (17/10/2026)
* **Tệp Excel**: [`261010_OB_TOOL_V1_RC65.xlsm`](./261010_OB_TOOL_V1_RC65.xlsm) (SHA-256: `FF7D14D2572FCABE42A9E3459A48AD29A1658CCC61789770BBE05B975427FA22`)
* **Gói bàn giao đầy đủ**: [`RC65_HANDOFF.zip`](./RC65_HANDOFF.zip) (SHA-256: `8576D07EC21BD560E7F76D2120CE4E5D97FAA52491B11B2BBA9881446C35ED0E`)
* **Tính năng đột phá**:
  * Tự động hủy/xóa dòng mã ngoài (cancel unplanned items) trong OB 3001 trước Save, chốt chặn Pre-Save Gate triệt tiêu 100% lỗi WMS `VL589`.
  * Khắc phục triệt để lỗi sập đọc lại chứng từ (`CONFLICT_READBACK_UNAVAILABLE: Object variable or With block variable not set`) trên VL03N.
  * Tích hợp cơ chế quét VL10D 1 lần cho nhánh 3001: nếu STO còn tồn mở (`Due > 0`) thì tự sinh OB; nếu hết STO thì bỏ qua batch thiếu (`SKIP_NO_STO_DUE_3001`), không gây sập hay kẹt vòng lặp.
* **Ghi chú kỹ thuật & Vá lỗi**:
  * [`CHANGELOG_RC64_TO_RC65.md`](./CHANGELOG_RC64_TO_RC65.md)
  * [`VBA_RC64_TO_RC65.diff`](./VBA_RC64_TO_RC65.diff)
  * [`RC65/RC65_3001_VL10D_DO_AUTO_AND_RC64LAN1_VL589_RECOVERY_PLAN.md`](./RC65/RC65_3001_VL10D_DO_AUTO_AND_RC64LAN1_VL589_RECOVERY_PLAN.md)

### 📜 Các bản trước đó
* **RC64 (16/10/2026)**:
  * [`261010_OB_TOOL_V1_RC64.xlsm`](./261010_OB_TOOL_V1_RC64.xlsm) (SHA-256: `9A9D9A3875D037C5DF5771932C1622C993780AE41131D52B6ADA90A8FC199FC5`)
  * [`RC64_HANDOFF.zip`](./RC64_HANDOFF.zip)
  * [`CHANGELOG_RC63_TO_RC64.md`](./CHANGELOG_RC63_TO_RC64.md)

### 📜 Các bản trước đó
* **RC63 (15/10/2026)**:
  * [`261010_OB_TOOL_V1_RC63.xlsm`](./261010_OB_TOOL_V1_RC63.xlsm) (SHA-256: `60EF66BDD0E44BA13E1A635CD6EB255749A5F8990F605A7FFB5254A112108D64`)
  * [`RC63_HANDOFF.zip`](./RC63_HANDOFF.zip)
  * [`CHANGELOG_RC62_TO_RC63.md`](./CHANGELOG_RC62_TO_RC63.md)
  * [`RC63/references/vl10d1(1).vbs`](./RC63/references/vl10d1%281%29.vbs)
* **RC62 (15/10/2026)**:
  * [`261010_OB_TOOL_V1_RC62.xlsm`](./261010_OB_TOOL_V1_RC62.xlsm)
  * [`RC62_HANDOFF.zip`](./RC62_HANDOFF.zip)
  * [`CHANGELOG_RC61_TO_RC62.md`](./CHANGELOG_RC61_TO_RC62.md)

---

## 🛠 Chức năng cốt lõi
1. **Tạo Outbound Delivery qua VL10D**: Tự động lọc STO, gom nhiều mã vật tư, tối ưu phân bổ Qty và xác thực trạng thái Due list.
2. **Quản lý & Giải phóng DO nháp (Kho 3001)**: Tự động phát hiện và xóa nhanh các DO nháp chưa xuất để nhả tồn kho cho việc sinh đơn mới.
3. **Phân bổ Batch qua VL02N**: Kiểm tra số lượng tồn, điền Batch và Storage Location theo đúng nguyên tắc FIFO và toàn vẹn lô (Whole Batch Gate).
4. **Điều phối ALL IN ONE**: Tự động nhận diện Ship-to (3001 hoặc 4007) để kích hoạt toàn bộ chu trình xử lý khép kín chỉ bằng 1 thao tác bấm nút.
