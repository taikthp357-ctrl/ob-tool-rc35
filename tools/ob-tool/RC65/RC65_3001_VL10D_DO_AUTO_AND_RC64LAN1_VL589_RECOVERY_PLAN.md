# RC65 — Kế hoạch nâng cấp từ RC64Lan1: VL10D nhánh 3001, DO_AUTO & phục hồi VL589

**Trạng thái:** Technical plan / chưa patch XLSM, chưa test SAP live.  
**Nguồn phân tích:** `261010_OB_TOOL_V1_RC64Lan1.xlsm`, ngày 17/10/2026, SHA256 `4097c974fe8c95787627e06a7cc8dc26b3f2c96fadd99f77a73d3bee89ba44ae`.  
**Baseline repo:** `tools/ob-tool/261010_OB_TOOL_V1_RC64.xlsm` (bản phát hành khác với RC64Lan1 người dùng gửi). Giữ nguyên release RC64 và hash; không ghi đè.  
**Tài liệu đối chiếu:** `tools/ob-tool/VBA_RC63_TO_RC64.diff`, `tools/ob-tool/RC63/references/vl10d1(1).vbs`, LOG/INPUT/BATCH_PLAN/LEDGER trong RC64Lan1.

## 1. Kết quả kiểm tra lượt RC64Lan1: chạy 01:03–01:04 ngày 17/10/2026

- Ship-to = **3001**, TEST_SCOPE=ALL, MODE=APPLY, nguyên batch bật.
- `PLAN 13 item; 18 allocations`; có **13 lượt SAVE_RESULT trên 13 OB**: **11 Save ACK VL337**, **2 SAVE UNKNOWN VL589**. LOG cuối có `EXCEL_SAVE_FINAL_SEC`, chưa xác nhận workbook bị crash hoàn toàn trong đoạn log có sẵn.
- **1012785994**, STO **3601045143**, material **BIW32011340AB**, gồm batch `261016-154` **51 PCS** và `261016-168` **70 PCS**. `REMOVE_UNPLANNED_ITEMS Removed=6; KeptIDs=80,50`; `PRESAVE_EXTRA_ITEM_INFO Item 50` vẫn còn ngoài kế hoạch; Save `VL589: could not be distributed`; VL03N cho biết OB còn tồn tại; `CONFLICT_READBACK_UNAVAILABLE: Object variable or With block variable not set`.
- **1012785656**, STO **3601044959**, material **BIW32025100AB**, batch `261016-155` **50 PCS** và `261016-122` **50 PCS**. `REMOVE_UNPLANNED_ITEMS Removed=4; KeptIDs=50,40`; item 40 ngoài kế hoạch vẫn hiện; cùng lỗi **VL589**, **VL03N EXISTS**, **CONFLICT_READBACK_UNAVAILABLE**.
- Ledger và BATCH_PLAN đánh dấu 4 allocations của hai OB trên là **UNSAVED**, không phải DONE. Sau đó tool chuyển lỗi Save thành `WAIT_FULL_BATCH_CAPACITY`, chờ 6+5 giây nhưng không đối soát được item hiện hữu rồi báo `DEFER_FINAL_FAIL`.
- Còn **7 batch INPUT** báo `WAIT_FULL_BATCH_CAPACITY` (bao gồm các dòng có ảnh hưởng từ hai OB lỗi và những mã thiếu OB ban đầu). Trạng thái cuối màn hình `INPUT MOI - NHAN NUT 1` có thể gây hiểu nhầm kết quả hoàn thành.
- Hiện **chưa đủ bằng chứng** để kết luận item ngoài kế hoạch là nguyên nhân duy nhất của VL589, hoặc hai OB đã bị SAP xóa/rollback. Không được xóa chúng hay tạo bù trước khi đối soát live.

## 2. Root cause cần vá: P0

### P0-A: Save VL589 và item ngoài kế hoạch

`REMOVE_UNPLANNED_ITEMS` đã xóa nhiều dòng nhưng `PRESAVE_EXTRA_ITEM_INFO` vẫn tìm được item nằm ngoài phân bổ hiện tại. Chính sách “bỏ qua không can thiệp” trước Save không đủ an toàn: lỗi phân phối WMS có thể khiến toàn bộ OB không được phân phối. Cần pre-save gate bao quát **mọi item** trong OB, không chỉ item được phân bổ.

- Item không thuộc plan nhưng do cùng mã INPUT / nguồn DO_AUTO: không tự suy luận nó vô hại.
- Nếu item phụ không cần, chỉ xóa khi xác nhận **chưa có Batch, chưa SLoc, chưa Pick/WT** và đúng phạm vi quyền quản lý; nếu không bảo đảm: đánh dấu `WAIT_EXTRA_ITEM`, **không Save OB đó**, vẫn chạy các OB khác.
- Trước Save, đọc lại toàn bộ item sống, so kế hoạch / trạng thái, không bỏ qua item phụ chưa đủ dữ kiện.

### P0-B: Recovery readback object null

`CONFLICT_READBACK_UNAVAILABLE: Object variable or With block variable not set` xuất hiện cả hai OB sau VL03N EXISTS. Sau thao tác điều hướng, phải lấy lại SAP session / wnd[0] / container / grid thực tế. Thêm guard Nothing, kiểm tra tcode và retry đọc lại có giới hạn, mỗi lần mở lại phải refresh control handle.

- `VL03N EXISTS` **không** nghĩa Save VL02N thành công; `VL589` không được ghi DONE.
- `SAVE_UNKNOWN` -> `READBACK_PENDING`; kiểm tra item/Batch/SLoc/Qty + trạng thái WMS, sau đó mới phân `SAVED_CONFIRMED`, `UNSAVED_CONFIRMED`, `EXISTS_UNVERIFIED` hoặc `ABSENT_CONFIRMED`.
- `EXISTS_UNVERIFIED` phải khóa **chỉ OB liên quan** cho đến khi xác minh được, không tự sửa và không tạo OB khác có thể trùng cùng STO.
- Không dùng `WAIT_LOCK_BACKOFF` chung cho mọi lỗi: khóa STO/OB khác với VL589/readback exception.

### P0-C: Chống queue khi một OB nhiều item

Một delivery 3001 có thể chứa nhiều item và nhiều material. Phải giữ mapping `OB -> {Item, Material, STO, STO Item, Qty, Batch, SLoc, Status}`. Nếu một item không được xử lý và có thể sinh WT/queue: cảnh báo `QUEUE_RISK`; không được chỉ trả `DONE` cho phần đã điền mà bỏ qua item phụ trong cùng OB.

## 3. Nâng nhánh 3001 bằng cơ chế VL10D nhưng khác 4007

### 3.1. Route theo Ship-to

- **4007**: giữ pipeline đặc thù đã có — VL10D STO Due trước/sau Refresh -> xác minh OB -> nếu thiếu tồn đã được xác nhận mới chọn xóa nguyên OB nháp 3001 theo Qty thực tế đã xóa -> tạo bù -> VL02N. Không coi Due giảm đơn độc là chứng minh do chính lượt Create.
- **3001**: **VL06O tồn tại -> xác định STO/Item chính xác -> kiểm tra VL10D hiện trạng -> phát hiện DO_AUTO / OB người khác / OB của user hiện tại theo STO -> chỉ tạo lượng thực sự thiếu -> Refresh + so sánh snapshot trước/sau -> VL02N**. Không sử dụng thủ tục “cancel 3001 để nhả tồn” dành riêng 4007 cho luồng 3001.
- **Ship-to khác**: giữ nhánh chuẩn hiện tại; không tự áp dụng logic 3001 hoặc 4007 nếu chưa có quy tắc cấu hình.

### 3.2. Đặc điểm chứng cứ RC64Lan1

Sheet `VL06O_DATA` có cột `H=STO`, `I=STO ITEM`, `O=CREATED_BY`, `P=CREATED_ON`, nên có thể xác định `DO_AUTO` thật từ **CREATED_BY**, không dựa vào tên hoặc đoán theo Qty. Có nhiều OB `CREATED_BY=DO_AUTO` khác nhau cùng Material; một số cùng STO, Qty phân tán trên nhiều OB.

Ví dụ snapshot:
- `1012774758` / `1012775446` cùng material `BIW32025800AC`, cùng STO `3601043797`, Qty **48+2=50**.
- `1012770379` / `1012776855` cùng `BIW32025100AB`, STO `3601044139`, Qty **49+1=50**.
Vì thế so sánh **tổng Qty theo STO + STO Item + Material**, không đòi từng OB có Qty đúng bằng nhu cầu.

### 3.3. Phân loại OB để tránh tạo trùng hoặc điền nhầm

Với khóa `(ShipTo, STO, STOItem, Material, UoM)`, snapshot VL06O ghi tất cả OB/item và `CREATED_BY`:
- `DO_AUTO`: nguồn OB tự động, có thể dùng để **đánh giá capacity đã tồn tại**; chỉ phân bổ nếu item còn rỗng Batch/SLoc, không Pick/WT, đúng nghiệp vụ, và không có người khác đang xử lý.
- `CURRENT_USER`: OB do chính user hiện tại tạo; phân biệt OB mới thuộc request với OB cũ bằng `CreateIntent + beforeSnapshot + exact STO/Item + Qty + time`. Không lấy bừa OB cũ cùng mã.
- `OTHER_USER`: **không auto chỉnh sửa/xóa** OB mà người khác đang thao tác; vẫn đưa lượng liên quan vào đối soát để ngăn Create trùng. Nếu không chắc tương quan, gắn `EXTERNAL_OB_CONFLICT`.
- `UNKNOWN_CREATOR`: không đồng nhất với DO_AUTO.
- Nếu STO có nhiều item cùng material, phải tách theo STO Item. Nếu INPUT không cung cấp STO/Item mà có nhiều ứng viên: yêu cầu lựa chọn/logic map đã xác thực, **không tùy ý chọn STO**.

### 3.4. Quy trình 3001 kiểu Confirmed Create / Do-auto reconciliation

1. Kiểm tra POSTING, nguyên Batch và INPUT fingerprint. Ghi `RunID`, request row, batch key, ship-to.
2. `VL06O_BEFORE`: lọc theo Material và 3001, đọc tất cả item + creator + STO + STO Item + Qty + SLoc/Batch/GM.
3. `VL10D_BEFORE`: chỉ đọc Due/Open đúng STO+Item+Material, chụp baseline. Đối chiếu tổng OB DO_AUTO/Current/Other cùng STO, bao gồm nhiều OB cộng lại.
4. `NEED_CREATE` = phần cần phục vụ chưa có **OB hợp lệ khả dụng** hoặc đã được chính request này phân bổ. Không cộng đúp Qty giữa snapshot, ledger và CREATE_PENDING. DO_AUTO/Other-user có OB không đồng nghĩa Batch của người dùng đã DONE.
5. Nếu cần Create, ghi durable `CREATE_INTENT` **trước nút tạo**; gửi Create đúng Qty hợp lệ. Sau đó Refresh VL10D và `VL06O_AFTER` đúng scope.
6. Đối soát bộ ba **DueDecrease + (OB set after - before) + receipt/exact STO Item**. So khớp Qty từng OB/item; không nhận OB do người khác tạo trùng thời điểm là kết quả của tool nếu không đủ tương quan.
7. `CREATE_FULL` khi có đủ capacity OB hợp lệ xác minh; `CREATE_PARTIAL` khi chỉ một số Batch đủ. Chỉ `VL02N_DONE` nếu Batch+SLoc đã Save thành công; không gộp hai loại DONE.
8. Nếu Create result không rõ, `CREATE_PENDING_VERIFY` và khóa tạo lặp trên cùng STO/item cho đến khi đối soát; không tự gửi nút Create lần hai.
9. Nếu một OB có 55 PCS nhưng thiếu OB tiếp theo, vẫn xử lý các Batch nguyên phù hợp trong 55; phần thiếu `WAIT_NEXT_OB`, không văng cả batch run.
10. Sau bất kỳ thao tác thêm/xóa OB trong cùng STO: làm mới riêng STO đó và tính lại việc còn thiếu; không chạy lại tất cả Material đã DONE.

**Lưu ý:** `STO Due` giảm có thể do DO_AUTO hoặc user khác cùng thời gian, hoặc do Refresh sai scope; không được dùng **một mình** để kết luận `CREATE_DONE`. Nếu dòng STO biến mất trên màn hình, trạng thái là `DUE_NOT_VISIBLE`, không tự gán `remaining=0` nếu chưa xác minh cùng selection/layout.

## 4. ALL IN ONE: Isolation, retry đúng loại, chặn DONE giả

- Giữ 1 nút `OB_RunAll`, đọc Ship-to trước khi chạy; không ép mặc định sang 4007.
- Dùng state machine theo `RequestRow + Batch + STO/Item`: `NEW -> SOURCE_OK -> OB_READY -> ITEM_READY -> SAVING -> DONE`; side states `WAIT_NEXT_OB`, `CREATE_PENDING_VERIFY`, `WAIT_EXTRA_ITEM`, `WAIT_SAP_LOCK`, `SAVE_UNKNOWN`, `QUEUE_RISK`.
- Khối xử lý mỗi OB có error boundary riêng. Bất kỳ exception nào lưu `Err.Number/Description/Erl/Phase`, checkpoint và tiếp tục **OB khác không phụ thuộc cùng delivery**. Không `Err.Raise` toàn lượt sau khi ghi lỗi local; trừ lỗi gây mất tính toàn vẹn toàn cục.
- Retry theo nguyên nhân: `WAIT_SAP_LOCK` có backoff có giới hạn; `SAVE_UNKNOWN` chỉ readback; `WAIT_NEXT_OB` tìm candidate theo STO và chỉ Create sau xác minh; `VL589` chặn ghi lại đến khi phân tích WMS.
- Khi một OB có 2 material và một item không xử lý được: không DONE cho nguyên OB; không đánh dấu item chưa Save là DONE.
- Kết thúc ALL IN ONE bắt buộc hiện popup **TOTAL INPUT batch, DONE, PARTIAL, WAIT, ERROR, QUEUE_RISK**, liệt kê OB và Qty còn thiếu. Cấm UI chỉ hiển thị `INPUT MOI` thay báo cáo thất bại.
- `LEDGER` lưu trạng thái theo **mỗi allocation/batch** và `RunID`. Nếu INPUT chỉnh sau lần chạy, tách request mới nhưng giữ ledger/unknown cũ để chống lặp và phục hồi.

## 5. Hạng mục sửa và nghiệm thu RC65

| Ưu tiên | Module/khu vực | Yêu cầu |
|---|---|---|
| P0 | `ModuleSAP` / SAVE | VL589 -> readback đáng tin; handle missing object; không DONE giả |
| P0 | `RemoveUnplannedItems` + presave gate | Không Save nếu item ngoài plan chưa rõ; ngăn queue |
| P0 | `ModuleMain` / recovery | Per-OB isolation; không biến SAVE_UNKNOWN thành capacity/lock error |
| P0 | `ModuleMain` / RC64 Due | Dòng Due không hiển thị != 0; Due giảm chưa đủ xác nhận chính request đã tạo |
| P1 | New 3001 VL10D planner | Baseline STO/Item, DO_AUTO/Other/Current creator, anti-duplicate |
| P1 | `OB_RunAll` + `LEDGER` | Retry theo từng batch, không làm lại những Batch DONE |
| P1 | UI INPUT/RESULT | Cảnh báo không hoàn tất nổi bật, không chỉ đổi R22 |
| P2 | Performance | Cache VL06O/POSTING theo request; refresh đúng STO và chỉ 1–2 vòng có điều kiện |

### Bộ test tối thiểu

1. RC64Lan1: OB `1012785994`, `1012785656` Save VL589; readback không crash; giữ nguyên hai OB đến khi rõ trạng thái.
2. OB có thêm item ngoài plan; không phát sinh queue ẩn, không bỏ qua item khi Save.
3. 3001: DO_AUTO có 48+2 trên cùng STO, INPUT cần 50 -> không Create thêm 50.
4. 3001: DO_AUTO có 17+2 trên **hai STO khác nhau**, INPUT cần 70 -> không cộng thành 19 vào cùng STO.
5. 3001: Other-user tạo OB cùng STO/Qty cùng thời điểm -> tool không nhận nhầm thành OB do Create của mình.
6. 3001: CREATE đã gửi, Refresh chưa đổi Due -> không Create lần 2 khi pending chưa xác minh.
7. 3001: OB 55/100 -> Save phần Batch nguyên khả thi, WAIT 45, xử lý những mã khác.
8. 4007: Due giảm đúng nhưng không có OB nào xác minh được -> CREATE_PENDING_VERIFY, không DONE và không cancel 3001.
9. 4007: STO Due row biến mất do filter/layout -> không ngộ nhận giảm hết và tự DONE.
10. Một OB có 2 Material, một lỗi -> các allocation ngoài OB đó vẫn chạy; trong OB không đánh dấu toàn bộ DONE.
11. Rerun sau lỗi/Excel crash -> ledger chỉ resume phần chưa xong, không ghi đè các Batch DONE.
12. SAP lock user chính mình -> không trừ Qty chưa xóa, nhận diện và thoát WAIT có kiểm soát.

## 6. Gate phát hành

- **NO-GO** nếu có bất kỳ `VL589`, `SAVE_UNKNOWN`, `PRESAVE_EXTRA_ITEM_INFO` chưa giải thích, `CREATE_PENDING_VERIFY`, hoặc tổng Batch DONE thấp hơn INPUT nhưng UI báo thành công toàn bộ.
- **GO** chỉ sau SAP live thử 3001 DO_AUTO/OTHER_USER cùng STO và 4007 Create/Refresh/Release, kiểm tra WMS queue và rerun idempotency.
- **Không tuyên bố đã fix** dựa vào tài liệu này. Cần patch, native Excel VBE Compile, test file-level ZIP/XML, và chạy SAP thực tế.
