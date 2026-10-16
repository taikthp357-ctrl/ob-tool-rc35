# Independent historical baseline audit — 2026-10-13

Read-only source review of the actual VBA embedded in `baseline/fg/extracted/261008_BBBG_FG_V2_2_14.xlsm` using oletools, workbook cells using openpyxl, and RC52 `.bas` sources. No SAP posting, mail, printing or attachment operations were run. This is source evidence, not live certification.

## Where historical BBBG data comes from

`Mod_SAP_Integration` reads MB51 and VL06I for the exact requested OB set. `Mod_Data_Reconciliation` compares SAP results; the input grid supplies metadata rather than authoritative quantities. BBBG lines use net MB51 quantity, material, batch, OB, unique joined material document IDs and inbound delivery IDs. A manual physical CF gate in the Python application is useful but does not replace this source comparison.

MB51 fields: OB=XBLNR, material=MATNR, batch=CHARG, quantity=MENGE, material document=MBLNR, unit=MEINS, movement=BWART, optional purchase order=EBELN. VL06I fields: OB=LIFEX, material=MATNR, batch=CHARG, delivery quantity=LFIMG, inbound delivery=VBELN, plant=WERKS, storage location=LGORT, goods movement status=WBSTK/WBSTA, unit=VRKME/MEINS. Imported rows must preserve identifiers as text.

The comparison key is `OB|Material|Batch`. Supported positive MB51 movements are 641, 643, 645, 647, 601, 351; corresponding reversals 642, 644, 646, 648, 602, 352 subtract absolute quantity. Unknown movements fail, including 101/102 in this baseline. Quantities are summed per key, units must match, and both sources must have equal key sets and quantities within 0.0001. Negative net quantity fails. Zero-net key absent from VL06I or zero on both sides is skipped. OB absent from both sources is skipped with NO_HISTORY and never inferred canceled. A failed or incomplete SAP read must never be treated as an empty successful result.

VL06I statuses A/B/C are all read. C is informational (already GR), and no second GR is performed. Missing batch, unit or key, unstable ALV row counts and negative VL06I quantity fail. Plant/SLoc determines shop, and mixed shops in a batch are rejected.

Metadata priority: exact OB/material input metadata, then OB-only manual metadata, then material mapping. Sender comes from CẤU_HÌNH G6:G10; receiver is ĐẦU_VÀO I8:I10 with manual values preserved; footer signer is K13. Master append dedup uses `Material|OB|Batch|IB|DOC`, rather than slip number alone.

## Configuration evidence

Actual source table CẤU_HÌNH: BODY=3001/1003, EDC=4007/*, Cơ khí=3098/*, XMĐ=2001/*. Default_Unit is EA, Default_Model empty, Outlook Preview_Only=1. The workbook contains 15 broken #REF! formulas, including CONFIG_SHOP aliases, so configuration must be imported from the central source table and checked instead of blindly copying alias sheets. Old path values belong to another workstation and require operator selection.

## Reusable SAP connector paths

COM entry is GetObject("SAPGUI").GetScriptingEngine. RC52 retains the chosen session for the entire run and fails on disconnect instead of silently switching. A specific session override is supported; AUTO uses the first connection/session. Python must initialize COM on one dedicated worker thread and serialize all session interactions.

MB51: clear MATNR/LGORT/CHARG/BWART/MBLNR ranges and multi-selection, set source plant 3093 (not receiving plant), layout /khoma, open `wnd[0]/usr/btn%_XBLNR_%_APP_%-VALU_PUSH`, clear popup selection, paste requested OBs, F8. VL06I: press `wnd[0]/usr/btnBUTTON7`, clear IT_LFDAT LOW/HIGH and IT_WBSTK LOW/HIGH, use IT_LIFEX multiple selection, F8, press `wnd[0]/tbar[1]/btn[18]` for item view. Both use `wnd[0]/usr/cntlGRID1/shellcont/shell`. Resolve columns by technical IDs and unambiguous captions, scroll pages, and verify stable complete row capture. Unknown popups and status E must fail with evidence.

## Bugs found in the pre-fix Python implementation

1. Stock summary and query aggregate every historical snapshot, inflating available inventory after refresh. Latest authoritative snapshot must be selected per scope; history remains separately accessible, including empty snapshots that clear a scope.
2. `unrestricted_qty or qty` and required quantity equivalents replace explicit zero with fallback values. Presence/None checks are required. Non-finite, negative and malformed quantities and blank identifiers need validation.
3. `/ob/batch-save` matches global `row-{idx}` or raw IDs without branch scope, letting one tab overwrite another. External IDs must be scoped or rejected; changing existing quantities must invalidate old validation.
4. Worker declares queue/COM isolation but executes synchronous callables, lacks CoInitialize usage and durable RUNNING records. A crash after an external side effect cannot be safely resumed without RUNNING persistence and reconciliation. Current APIs bypass pipeline lock entirely. Do not enable external write automation for this pilot.
5. Resume gathers all successful step names over a business key, not a pinned job/config/input revision. Old success can skip changed work. Uncertain takes precedence over historical success.
6. Health reported CONNECTED unconditionally and BBBG marked UNC/GOS/log success without artifacts. Fixes need truthful states and real path/hash evidence.

## Evening pilot acceptance

Feasible scope: local single-operator pilot with real exported MB51/VL06I data, explicit source completeness, reconciliation, physical CF, downloadable BBBG and operator comparison with original Excel. Connected status requires an actual read-only session probe. Native SAP reads require selector validation on the actual workstation. SAP posting, mail, printing, UNC and GOS remain unverified until separately exercised and recorded.

Required cases: full match, mismatch, missing inbound key, no-history OB, positive/reversal net zero, unsupported movement, unit mismatch, mixed shop, leading-zero identifiers, incomplete source, zero physical CF, stock refresh twice, empty latest stock snapshot, two branches with row 1, repeated BBBG request, restart after uncertain step, and artifact contents matching source lines. Pilot readiness must distinguish offline PASS from SAP live PASS.

## Independent reconciliation regression result

The new `tests/test_source_reconciliation.py` runs against isolated temporary SQLite databases and creates actual pick orders plus explicit physical CF. All 36 cases passed after fixing three review findings: missing/nonfinite/boolean quantities, missing/invalid VL06I GR status, and a VL06I-only zero key that previously disappeared. Coverage includes six positive/reversal pairs, 101/102 rejection, units, missing reciprocal keys, no-history/zero-net output rejection, duplicates, mixed shops, source metadata/completeness, pick quantity mismatch, modified pick revalidation and expired source revalidation. This verifies imported-source rules without proving SAP GUI capture works live.

The BBBG_NEW template has no hidden rows/columns or historical personnel values outside its print area. Personnel formulas A35/C35 must be overwritten and all workbook defined names pointing to deleted input/config sheets removed during standalone export. Sender/receiver ID, phone, title and approver should stay blank until explicitly configured. The formal template document ID/version/effective date are static form metadata. Its logo image requires explicit copying because openpyxl copy_worksheet omits images.
