# -*- coding: utf-8 -*-
"""Workflow API routes for MA CONTROL 5-step process.

Step 1: Stock Ingestion & Warnings
Step 2: OB Checks & Fill (RC52 logic)
Step 3: Pick Order Generation (Status: 'Chờ người dùng CF')
Step 4: Manual CF Verification (POST /api/pick-orders/{id}/verify-cf)
Step 5: BBBG Generation (Locked until status == 'Đã CF')
"""
from typing import List, Dict, Any, Optional
import base64
from fastapi import APIRouter, HTTPException, Depends, Query, Path, Request
from fastapi.responses import FileResponse
from pathlib import Path as FilePath
from app.models.schemas import (
    StockUploadRequest,
    OBBatchSaveRequest,
    OBCheckRequest,
    OBFillRequest,
    PickOrderCreateRequest,
    PickVerifyCFRequest,
    ConfirmPickLegacyRequest,
    WaitingCFLineSaveRequest,
    WaitingCFBulkSaveRequest,
    WaitingCFDirectPasteRequest,
    WaitingCFUpdateLineRequest,
    LowStockConfigUpdate,
    LowStockInspectRequest,
    LowStockApplyRequest,
    BBBGCreateRequest,
    SAPDownloadMB52Request,
    DirectReconcileSAPRequest,
    LowStockTransferRequest,
    normalize_pick_status,
    normalize_ob_status
)
from app.services.stock_service import StockService
from app.services.ob_service import OBService
from app.services.pick_cf_service import PickCFService
from app.services.bbbg_service import BBBGService
from app.services.sap_worker import sap_worker
from app.core.database import get_connection
from app.services.reconciliation_service import save_reconciliation
from app.services.source_import import parse_source
router = APIRouter(prefix="/api", tags=["Workflow"])
# -------------------------------------------------------------
# STEP 1: STOCK ENDPOINTS
# -------------------------------------------------------------
@router.post("/stock/upload")
def upload_stock_snapshot(req: StockUploadRequest):
    service = StockService()
    try:
        res = service.create_snapshot(source_type=req.source, plant=req.plant or "3001", sloc=req.sloc or "1003", items=req.items)
    except (ValueError, TypeError) as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    return res
@router.get("/stock/summary")
def get_stock_summary(plant: Optional[str] = None):
    service = StockService()
    return service.get_zone_summary(plant=plant)
@router.get("/stock/items")
def query_stock_items(
    material_code: Optional[str] = None,
    plant: Optional[str] = None,
    sloc: Optional[str] = None,
    pull_line_only: bool = False
):
    service = StockService()
    return service.query_stock(
        material_code=material_code,
        plant=plant,
        sloc=sloc,
        pull_line_only=pull_line_only
    )
# -------------------------------------------------------------
# STEP 2: OB ENDPOINTS (RC52 LOGIC)
# -------------------------------------------------------------
@router.get("/ob/items")
def get_ob_items(branch: str = Query("BODY_3001", alias="branch_tab"), status: Optional[str] = Query(None)):
    conn = get_connection()
    try:
        rows = conn.execute("""
            SELECT * FROM ob_items
            WHERE branch_tab = ? OR branch = ?
            ORDER BY row_order ASC, id ASC
        """, (branch, branch)).fetchall()
        items = []
        for r in rows:
            d = dict(r)
            # Harmonize required_qty / req_qty
            d["required_qty"] = d["required_qty"] or d["req_qty"] or 0.0
            d["req_qty"] = d["required_qty"]
            norm = normalize_ob_status(d.get("status"))
            d.update(norm)
            if status and status.strip():
                st = status.strip()
                if d.get("status") != st and d.get("status_code") != st and d.get("status_group") != st:
                    continue
            items.append(d)
        return items
    finally:
        conn.close()
@router.post("/ob/batch-save")
def batch_save_ob_items(req: OBBatchSaveRequest):
    import math
    conn = get_connection()
    branch = req.branch_tab or req.branch or "BODY_3001"
    if branch not in {"BODY_3001", "EDC_4007"}:
        conn.close()
        raise HTTPException(status_code=400, detail="Nhánh không hợp lệ")
    saved_ids = []
    try:
        with conn:
            conn.execute("BEGIN IMMEDIATE")
            for idx, it in enumerate(req.items, 1):
                raw_id = it.get("id")
                raw_uid = str(it.get("item_uid") or f"row-{idx}")
                uid = raw_uid if raw_uid.startswith(branch + ":") else f"{branch}:{raw_uid}"
                mat = str(it.get("material_code") or "").strip()
                model = str(it.get("model") or "").strip()
                qty_raw = it["required_qty"] if "required_qty" in it else it.get("req_qty", 0)
                qty = float(qty_raw or 0)
                if not mat:
                    if not model and qty == 0 and not it.get("ob_number") and not it.get("batch") and not raw_id:
                        continue
                    raise ValueError("Dòng OB thiếu mã hàng")
                if isinstance(qty_raw, bool) or not math.isfinite(qty) or qty < 0:
                    raise ValueError("Số lượng yêu cầu không hợp lệ")
                batch = str(it.get("batch") or "").strip()
                shop = str(it.get("shop") or ("BODY" if "BODY" in branch else "EDC")).strip()
                sto = str(it.get("sto_number") or "").strip()
                ob = str(it.get("ob_number") or "").strip()
                existing = None
                if isinstance(raw_id, int):
                    existing = conn.execute("SELECT * FROM ob_items WHERE id = ?", (raw_id,)).fetchone()
                    if existing and existing["branch_tab"] != branch:
                        raise ValueError("ID dòng thuộc tab khác; không được ghi đè")
                if not existing:
                    existing = conn.execute("SELECT * FROM ob_items WHERE item_uid = ? AND branch_tab = ?", (uid, branch)).fetchone()
                if existing:
                    changed = any(existing[k] != v for k, v in [("material_code", mat), ("model", model), ("required_qty", qty), ("batch", batch), ("shop", shop), ("sto_number", sto), ("ob_number", ob)])
                    if changed and existing["pick_order_id"]:
                        raise ValueError("Dòng đã tạo pick; không sửa nguồn sau xác nhận")
                    core_changed = any(existing[k] != v for k, v in [("material_code", mat), ("required_qty", qty), ("batch", batch), ("shop", shop)])
                    if core_changed:
                        status = "Chưa kiểm tra"
                    elif existing["status"] in ("Sẵn sàng điền", "READY", "READY_FOR_PICK", "Đủ điều kiện tạo pick"):
                        status = existing["status"]
                    else:
                        status = "Chưa kiểm tra" if changed else existing["status"]
                    conn.execute("""
                        UPDATE ob_items SET item_uid=?, row_order=?,material_code=?,model=?,req_qty=?,required_qty=?,batch=?,shop=?,sto_number=?,ob_number=?,status=?,selected=?,updated_at=CURRENT_TIMESTAMP
                        WHERE id=? AND branch_tab=?
                    """, (uid,idx,mat,model,qty,qty,batch,shop,sto,ob,status,int(bool(it.get("selected", True))),existing["id"],branch))
                    saved_ids.append(existing["id"])
                else:
                    cur = conn.execute("""
                        INSERT INTO ob_items(item_uid,branch_tab,branch,row_order,material_code,model,req_qty,required_qty,batch,shop,sto_number,ob_number,status,selected)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (uid,branch,branch,idx,mat,model,qty,qty,batch,shop,sto,ob,"Chưa kiểm tra",int(bool(it.get("selected", True)))))
                    saved_ids.append(cur.lastrowid)
        return {"status":"success","count":len(saved_ids),"ids":saved_ids}
    except (ValueError, TypeError) as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    finally:
        conn.close()
@router.post("/ob/check")
def check_ob_items_route(req: OBCheckRequest):
    service = OBService()
    updated = service.check_ob_items(branch_tab=req.branch_tab, item_ids=req.item_ids, source_plant=req.source_plant, source_sloc=req.source_sloc)
    return {"status": "success", "checked_count": len(updated), "items": updated}
@router.post("/ob/fill")
def fill_ob_items_route(req: OBFillRequest):
    service = OBService()
    try:
        filled = service.fill_ob_items(branch_tab=req.branch_tab, item_ids=req.item_ids)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    return {"status": "success", "filled_count": len(filled), "items": filled}
@router.post("/ob/delete-draft-3001")
def delete_draft_3001_route(ob_number: str = Query(...), simulated_condition: Optional[str] = Query(None)):
    service = OBService()
    return service.delete_draft_3001(ob_number, simulated_condition=simulated_condition)
# -------------------------------------------------------------
# STEP 3: PICK ORDER CREATION
# -------------------------------------------------------------
@router.post("/pick/create")
@router.post("/pick-orders")
def create_pick_order_route(req: PickOrderCreateRequest):
    service = PickCFService()
    try:
        res = service.create_pick_order(
            shop=req.shop,
            model=req.model,
            items=req.items,
            branch_tab=req.branch_tab
        )
        return res
    except ValueError as err:
        raise HTTPException(status_code=400, detail=str(err))
@router.get("/pick/orders")
@router.get("/pick-orders")
def get_pick_orders_route():
    service = PickCFService()
    return service.list_pick_orders()
@router.get("/pick-orders/{id}")
def get_pick_order_detail(id: str):
    conn = get_connection()
    try:
        if id.isdigit():
            row = conn.execute("SELECT * FROM pick_orders WHERE id = ?", (int(id),)).fetchone()
        else:
            row = conn.execute("SELECT * FROM pick_orders WHERE pick_order_code = ? OR pick_order_no = ?", (id, id)).fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Không tìm thấy đơn pick")
        po = dict(row)
        po.update(normalize_pick_status(po.get("status")))
        items = conn.execute("SELECT * FROM pick_order_items WHERE pick_order_id = ?", (po["id"],)).fetchall()
        po["items"] = [dict(it) for it in items]
        return po
    finally:
        conn.close()
# -------------------------------------------------------------
# STEP 4: MANUAL CF VERIFICATION GATE
# -------------------------------------------------------------
@router.post("/pick-orders/{id}/verify-cf")
def verify_cf_route(id: str, req: PickVerifyCFRequest):
    """BƯỚC 4: NGƯỜI DÙNG CF (XÁC NHẬN THỦ CÔNG BẮT BUỘC).
    Chỉ kích hoạt khi người dùng bấm 'Kiểm tra CF' trên giao diện.
    Tuyệt đối không tự động CF ngầm!
    """
    service = PickCFService()
    try:
        res = service.verify_cf(
            pick_order_ref=id,
            verified_by=req.verified_by,
            actual_quantities=req.actual_quantities,
            notes=req.notes
        )
        return res
    except LookupError as err:
        raise HTTPException(status_code=404, detail=str(err))
    except Exception as err:
        raise HTTPException(status_code=400, detail=str(err))
@router.post("/pick/{pick_order_no}/confirm")
def confirm_pick_order_legacy_route(pick_order_no: str, req: ConfirmPickLegacyRequest):
    """Legacy alias for Step 4 manual CF confirmation."""
    service = PickCFService()
    try:
        res = service.verify_cf(
            pick_order_ref=pick_order_no,
            verified_by=req.confirmed_by,
            actual_quantities=req.actual_quantities
        )
        return res
    except LookupError as err:
        raise HTTPException(status_code=404, detail=str(err))
    except Exception as err:
        raise HTTPException(status_code=400, detail=str(err))

@router.get("/waiting-cf")
@router.get("/pick-orders/waiting-cf/lines")
def list_waiting_cf_route(
    ob_number: Optional[str] = Query(None),
    pick_order_code: Optional[str] = Query(None),
    material_code: Optional[str] = Query(None),
    batch: Optional[str] = Query(None),
    status: Optional[str] = Query(None),
    include_completed: bool = Query(False),
    q: Optional[str] = Query(None),
):
    service = PickCFService()
    return service.list_waiting_cf_lines(
        ob_number=ob_number,
        pick_order_code=pick_order_code,
        material_code=material_code,
        batch=batch,
        status=status,
        include_completed=include_completed,
        q=q,
    )

@router.post("/waiting-cf/save-line")
def save_waiting_cf_line_route(req: WaitingCFLineSaveRequest):
    service = PickCFService()
    try:
        return service.save_line_cf(
            line_id=req.line_id,
            actual_qty=req.actual_qty,
            verified_by=req.verified_by,
            reason=req.reason,
            evidence=req.evidence,
        )
    except LookupError as err:
        raise HTTPException(status_code=404, detail=str(err))
    except ValueError as err:
        raise HTTPException(status_code=400, detail=str(err))

@router.post("/waiting-cf/bulk-save")
def bulk_save_waiting_cf_route(req: WaitingCFBulkSaveRequest):
    service = PickCFService()
    results = []
    try:
        for item in req.lines:
            line_id = int(item["line_id"])
            actual_qty = float(item["actual_qty"])
            reason = item.get("reason") or item.get("shortage_reason")
            evidence = item.get("evidence") or req.notes
            res = service.save_line_cf(
                line_id=line_id,
                actual_qty=actual_qty,
                verified_by=req.verified_by,
                reason=reason,
                evidence=evidence,
            )
            results.append(res)
        return {"status": "success", "saved_count": len(results), "results": results}
    except LookupError as err:
        raise HTTPException(status_code=404, detail=str(err))
    except (ValueError, KeyError, TypeError) as err:
        raise HTTPException(status_code=400, detail=str(err))

@router.get("/waiting-cf/history")
def get_waiting_cf_history_route(pick_order_id: Optional[int] = Query(None), limit: int = Query(100)):
    service = PickCFService()
    return service.list_cf_history(pick_order_id=pick_order_id, limit=limit)

@router.post("/waiting-cf/paste-excel")
def paste_waiting_cf_excel_route(req: WaitingCFDirectPasteRequest):
    service = PickCFService()
    try:
        return service.paste_excel_waiting_cf(
            raw_text=req.raw_text,
            items=req.items,
            allow_duplicates=req.allow_duplicates,
            direct_order_ref=req.direct_order_ref,
        )
    except ValueError as err:
        raise HTTPException(status_code=400, detail=str(err))

@router.put("/waiting-cf/lines/{line_id}")
def update_waiting_cf_line_route(line_id: int, req: WaitingCFUpdateLineRequest):
    service = PickCFService()
    try:
        return service.update_waiting_cf_line(
            line_id=line_id,
            ob_number=req.ob_number,
            material_code=req.material_code,
            plant=req.plant,
            batch=req.batch,
            qty=req.qty,
            packing_box=req.packing_box,
            rack=req.rack,
        )
    except LookupError as err:
        raise HTTPException(status_code=404, detail=str(err))
    except ValueError as err:
        raise HTTPException(status_code=400, detail=str(err))

@router.delete("/waiting-cf/lines/{line_id}")
def delete_waiting_cf_line_route(line_id: int):
    service = PickCFService()
    try:
        return service.delete_waiting_cf_line(line_id=line_id)
    except LookupError as err:
        raise HTTPException(status_code=404, detail=str(err))
    except ValueError as err:
        raise HTTPException(status_code=400, detail=str(err))

# -------------------------------------------------------------
# LOW STOCK & MULTI-BATCH PRIORITY ALLOCATION ROUTES
# -------------------------------------------------------------
@router.get("/stock/low-stock/config")
def get_low_stock_config_route():
    service = StockService()
    return service.get_low_stock_config()

@router.post("/stock/low-stock/config")
def update_low_stock_config_route(req: LowStockConfigUpdate):
    service = StockService()
    try:
        return service.update_low_stock_config(
            default_threshold=req.default_threshold,
            max_age_hours=req.max_age_hours,
            thresholds=req.thresholds,
        )
    except ValueError as err:
        raise HTTPException(status_code=400, detail=str(err))

@router.get("/stock/low-stock/alerts")
def get_low_stock_alerts_route(
    plant: Optional[str] = Query(None),
    scope_sloc: Optional[str] = Query(None),
    shop: Optional[str] = Query(None),
):
    service = StockService()
    return service.list_low_stock_alerts(plant=plant, scope_sloc=scope_sloc, shop=shop)

@router.post("/stock/low-stock/inspect")
def inspect_low_stock_priority_route(req: LowStockInspectRequest):
    service = StockService()
    try:
        return service.inspect_low_stock_priority(
            material_code=req.material_code,
            plant=req.plant,
            required_qty=req.required_qty,
            scope_sloc=req.scope_sloc,
            shop=req.shop,
            max_batches=req.max_batches,
        )
    except ValueError as err:
        raise HTTPException(status_code=400, detail=str(err))

@router.post("/stock/low-stock/apply")
def apply_low_stock_allocation_route(req: LowStockApplyRequest):
    service = StockService()
    try:
        return service.apply_low_stock_allocation(
            material_code=req.material_code,
            plant=req.plant,
            required_qty=req.required_qty,
            allocations=req.allocations,
            shop=req.shop,
            scope_sloc=req.scope_sloc,
            target=req.target,
            ob_number=req.ob_number,
            branch_tab=req.branch_tab,
            evidence_note=req.evidence_note,
            other_shop_reserves=req.other_shop_reserves,
        )
    except ValueError as err:
        raise HTTPException(status_code=400, detail=str(err))
# -------------------------------------------------------------
# STEP 5: BBBG GENERATION
# -------------------------------------------------------------
@router.post("/bbbg/generate")
@router.post("/bbbg/create")
def generate_bbbg_route(req: BBBGCreateRequest):
    """BƯỚC 5: TẠO BBBG THEO CHUẨN V2.2.14.
    BẤT BIẾN NGHIỆP VỤ: Chặn cứng mọi đơn pick chưa được xác nhận CF ('Đã CF').
    """
    po_ref = req.pick_order_id or req.pick_order_code or req.pick_order_no
    if not po_ref:
        raise HTTPException(status_code=400, detail="Thiếu thông tin đơn pick để tạo BBBG")
    sender = req.sender_name or req.handover_person or ""
    receiver = req.receiver_name or req.receiver_person or ""
    bin_note = req.storage_bin or req.storage_location_note or ""
    service = BBBGService()
    try:
        res = service.create_bbbg(
            pick_order_ref=po_ref,
            sender_name=sender,
            receiver_name=receiver,
            storage_bin=bin_note,
            slip_no=req.slip_no,
            doc_no=req.doc_no,
            is_external=req.is_external,
            template_mode=req.template_mode,
            font_size=req.font_size
        )
        res["download_url"] = f"/api/bbbg/{res['id']}/download"
        return res
    except LookupError as err:
        raise HTTPException(status_code=404, detail=str(err))
    except ValueError as err:
        raise HTTPException(status_code=400, detail=str(err))
    except Exception as err:
        raise HTTPException(status_code=500, detail=f"Lỗi tạo BBBG: {err}")
@router.get("/bbbg/records")
def get_bbbg_records():
    service = BBBGService()
    return service.list_bbbg_records()


@router.get("/bbbg/{record_id}/download")
def download_bbbg(record_id: int):
    try:
        filename = FilePath(BBBGService().get_export_path(record_id))
        return FileResponse(filename, filename=filename.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except LookupError as exc:
        raise HTTPException(status_code=404, detail=str(exc))
    except (ValueError, OSError) as exc:
        raise HTTPException(status_code=409, detail=str(exc))


@router.post("/pick-orders/{pick_id}/reconcile")
def reconcile_pick_sources(pick_id: int, payload: Dict[str, Any]):
    conn = get_connection()
    try:
        with conn:
            conn.execute("BEGIN IMMEDIATE")
            evidence = save_reconciliation(conn, pick_id, payload)
        return evidence
    except LookupError as exc:
        raise HTTPException(status_code=404, detail=str(exc))
    except (ValueError, TypeError) as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    finally:
        conn.close()


@router.post("/pick-orders/{pick_id}/sources/{kind}")
async def upload_pick_source(pick_id: int, kind: str, request: Request, filename: str, captured_at: str, complete: bool = False):
    import hashlib
    import json
    if kind not in {"mb51", "vl06i"}:
        raise HTTPException(status_code=400, detail="Chọn nguồn MB51 hoặc VL06I")
    if not complete:
        raise HTTPException(status_code=400, detail="Cần xác nhận đã xuất đủ OB của đơn pick")
    raw = await request.body()
    conn = get_connection()
    try:
        if conn.execute("SELECT 1 FROM bbbg_records WHERE pick_order_id=? AND status='VERIFIED_IMPORTED'", (pick_id,)).fetchone():
            raise HTTPException(status_code=409, detail="Nguồn của BBBG đã xuất được khóa; không thay bằng file khác")
        obs = {str(r["ob_number"]) for r in conn.execute("SELECT ob_number FROM pick_order_items WHERE pick_order_id=?", (pick_id,)).fetchall()}
        if not obs:
            raise LookupError("Không tìm thấy đơn pick")
        rows = parse_source(kind, raw, filename)
        source = {"rows": rows, "metadata": {"filename": filename, "captured_at": captured_at, "complete": True,
                  "requested_obs": sorted(obs), "file_sha256": hashlib.sha256(raw).hexdigest()}}
        from app.services.reconciliation_service import _metadata
        _metadata(source["metadata"], obs)
        if any(r["ob_number"] not in obs for r in rows):
            raise ValueError("File chứa OB khác ngoài đơn pick; xuất riêng danh sách cần đối chiếu")
        with conn:
            conn.execute("INSERT INTO pick_source_uploads(pick_order_id,kind,source_json) VALUES(?,?,?) ON CONFLICT(pick_order_id,kind) DO UPDATE SET source_json=excluded.source_json", (pick_id, kind, json.dumps(source, ensure_ascii=False)))
            conn.execute("DELETE FROM source_reconciliations WHERE pick_order_id=?", (pick_id,))
        return {"status": "SOURCE_IMPORTED", "kind": kind, "row_count": len(rows), "file_sha256": source["metadata"]["file_sha256"], "reconciled": False}
    except LookupError as exc:
        raise HTTPException(status_code=404, detail=str(exc))
    except (ValueError, TypeError, OSError) as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    finally:
        conn.close()


@router.post("/pick-orders/{pick_id}/reconcile-uploaded")
def reconcile_uploaded_sources(pick_id: int):
    import json
    conn = get_connection()
    try:
        sources = {r["kind"]: json.loads(r["source_json"]) for r in conn.execute("SELECT * FROM pick_source_uploads WHERE pick_order_id=?", (pick_id,)).fetchall()}
        if set(sources) != {"mb51", "vl06i"}:
            raise ValueError("Cần tải đủ hai nguồn MB51 và VL06I")
        payload = {k: v["rows"] for k, v in sources.items()}
        payload["metadata"] = {k: v["metadata"] for k, v in sources.items()}
        with conn:
            conn.execute("BEGIN IMMEDIATE")
            evidence = save_reconciliation(conn, pick_id, payload)
        return evidence
    except (ValueError, TypeError) as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    finally:
        conn.close()


@router.post("/stock/import-file")
async def import_stock_file(request: Request, filename: str):
    try:
        rows = parse_source("mb52", await request.body(), filename)
        return StockService().create_snapshot("SAP_MB52_IMPORT", rows[0]["plant"], rows[0]["sloc"], rows, notes=f"Nguồn: {filename}")
    except (ValueError, TypeError, OSError) as exc:
        raise HTTPException(status_code=400, detail=str(exc))


# -------------------------------------------------------------
# SAP MB52 AUTOMATED DIRECT DOWNLOAD & SNAPSHOT SYNC
# -------------------------------------------------------------
@router.post("/stock/sap/download-mb52")
def download_sap_mb52_route(req: Optional[SAPDownloadMB52Request] = None):
    service = StockService()
    plants = req.plants if req else None
    sloc = req.sloc if req else None
    layout = req.layout if req else "/KHOMA"
    test_bytes = base64.b64decode(req.test_bytes_b64) if (req and req.test_bytes_b64) else None
    test_filename = req.test_filename if req else None
    return service.download_and_sync_sap_mb52(
        plants=plants,
        sloc=sloc,
        layout=layout,
        test_bytes=test_bytes,
        test_filename=test_filename
    )


# -------------------------------------------------------------
# DIRECT SAP BBBG & MB51 RECONCILIATION
# -------------------------------------------------------------
@router.post("/pick-orders/{pick_id}/direct-reconcile-sap")
def direct_reconcile_sap_route(pick_id: int, req: Optional[DirectReconcileSAPRequest] = None):
    conn = get_connection()
    try:
        from app.services.reconciliation_service import reconcile_direct_sap
        test_mb51 = base64.b64decode(req.test_mb51_bytes_b64) if (req and req.test_mb51_bytes_b64) else None
        test_vl06i = base64.b64decode(req.test_vl06i_bytes_b64) if (req and req.test_vl06i_bytes_b64) else None
        fname_mb51 = req.test_mb51_filename if req else None
        fname_vl06i = req.test_vl06i_filename if req else None
        return reconcile_direct_sap(
            conn,
            pick_id,
            test_mb51_bytes=test_mb51,
            test_vl06i_bytes=test_vl06i,
            test_mb51_filename=fname_mb51,
            test_vl06i_filename=fname_vl06i
        )
    except LookupError as exc:
        raise HTTPException(status_code=404, detail=str(exc))
    except (ValueError, TypeError) as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    finally:
        conn.close()


# -------------------------------------------------------------
# VISUAL STOCK MATRIX (LINE + KHO) BODY REPORT & DRILLDOWN
# -------------------------------------------------------------
@router.get("/stock/visual-matrix")
def get_visual_stock_matrix_route():
    service = StockService()
    return service.get_visual_stock_matrix()


@router.get("/stock/visual-matrix/materials")
def get_visual_matrix_materials_route(
    model: Optional[str] = Query(None),
    bucket: Optional[str] = Query(None)
):
    service = StockService()
    return service.get_visual_matrix_materials(model_key=model, bucket_key=bucket)


# -------------------------------------------------------------
# TWO STOCK TABLES (PRIORITY & TOTAL STOCK)
# -------------------------------------------------------------
@router.get("/stock/two-tables")
def get_two_stock_tables_route(
    plant: str = Query("3001")
):
    service = StockService()
    return service.get_two_stock_tables(plant=plant)


# -------------------------------------------------------------
# KHU GỢI Ý TỒN THẤP & CHUYỂN GỢI Ý SANG OB
# -------------------------------------------------------------
@router.get("/stock/low-stock/suggestions")
def get_low_stock_suggestions_route(
    plant: str = Query("3001"),
    scope_sloc: Optional[str] = Query(None),
    shop: str = Query("BODY")
):
    service = StockService()
    return service.get_low_stock_suggestions(plant=plant, scope_sloc=scope_sloc, shop=shop)


@router.post("/stock/low-stock/transfer-to-ob")
def transfer_low_stock_suggestions_route(req: LowStockTransferRequest):
    service = StockService()
    try:
        return service.transfer_suggestions_to_ob(items=req.items, branch_tab=req.branch_tab)
    except ValueError as err:
        raise HTTPException(status_code=400, detail=str(err))
