# BBBG baseline and current export

Reviewed `baseline/fg/extracted/261008_BBBG_FG_V2_2_14.xlsm` with openpyxl (read) and oletools olevba (VBA extraction). Authoritative generation module is `Mod_BBBG_Generation.bas`; reconciliation is `Mod_Data_Reconciliation.bas`.

## Historical field mapping

| Cell / column | Historical meaning | Current exported value |
|---|---|---|
| I1 | SoPhieu | slip_no |
| E1 | ViTri | storage_bin |
| A4 | Date/time | frozen handover_date; no volatile NOW formula |
| C5 | Sender from config G6 | sender_name |
| G5 | Receiver from input I8 | receiver_name |
| C7 / G7 | Sender / receiver department | Kho MA / shop |
| A13 onward | Index | sequential index |
| B | Material | text material_code, leading zeros preserved |
| C | Batch | text batch or `/` |
| D | Bill/container | `/` (not supplied) |
| E | Inbound / material document | reconciled IB / Doc; draft uses supplied doc_no, no fabricated SAP number |
| F | Unit | reconciled source unit, draft PCS default |
| G | Handover quantity | reconciled net source quantity; draft explicit cf_actual_qty |
| H / I / J | Received qty / GR tick / GR material doc | empty; never infer GR from CF |
| K | Note | source reconciliation or draft caveat |
| A31:K35 | Total and signatures | copied styled footer, relocated after max(18, line_count) rows |

No historical employee codes, phone numbers, titles or approvals are invented. Footer sender/receiver comes from current request. Historical formula references are replaced by fixed values so no dangling links or stale historical people appear. Only the copied approved `BBBG_NEW` form is exported as macro-free .xlsx; historical input/config/email/run-state sheets and VBA are excluded. Borders, merged cells, row/column geometry and page setup are retained. This is a runtime export implementation using openpyxl to copy the historical Excel form, not a redesigned artifact.

## The operational gate

V2.2.14 compares MB51 net movement quantities with VL06I by OB|Material|Batch, in both directions, with equal units and tolerance 0.0001; net quantities must be positive. The source supplies real IB, Doc, Plant and SLoc. Pick/CF quantity is not sufficient evidence of a SAP posting. No SAP session, attachment, print, email or UNC success is claimed by this exporter.

If imported-source reconciliation exists and matches the current pick/CF, export uses those authoritative rows and sets `VERIFIED_IMPORTED` / `IMPORTED_RECONCILED`; the visible form still says GR/SAP attachment is unconfirmed. Otherwise export is prominently marked draft, stored as `DRAFT_CREATED`, and reports `is_draft=true`, `reconciliation_status=NOT_VERIFIED`. The pick stays confirmed until genuine handover evidence is available.

## Reliability and evidence

Export validates all actual quantities, writes a temporary workbook, reopens it to check totals, then publishes the file without overwriting another file. Database transaction records the file path and SHA-256 in job_events evidence_json only after successful export. Ordinary failures remove the uncommitted file and roll back records. An OS/process crash between file publication and DB commit can leave an unreferenced file; it cannot mark a failed operation as successful. File integrity is checked before download and idempotent retry. Retries reuse the existing record for the same pick; different slips/documents on that pick are rejected. Legacy records without file evidence need operator recovery, rather than fabricated success.

For the pilot, operators may review a draft before source import. After successful source reconciliation, repeating create for the same pick upgrades that draft to `VERIFIED_IMPORTED` using the same ID, slip and dedup key. It publishes a new file version and updates evidence atomically, preserving the old draft file for history. Further retries reuse the verified version. UNC/general/shop ledgers and GOS flags remain PENDING/0. External items do not bypass reconciliation. Tests use a temporary SQLite database and export directory, never production SAP.
