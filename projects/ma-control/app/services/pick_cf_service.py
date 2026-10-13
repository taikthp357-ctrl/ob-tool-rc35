# -*- coding: utf-8 -*-
"""Pick Order & Manual CF Verification Service (Step 3 & Step 4).

STRICT SYSTEM INVARIANT:
- Step 3 creates pick orders in status 'Chờ người dùng CF' (never auto-CF).
- Step 4 requires an explicit manual trigger via verify_cf() when user clicks 'Kiểm tra CF'.
- Only after status == 'Đã CF' is Step 5 (BBBG) unlocked.
"""

import sqlite3
import uuid
import datetime
import math
import hashlib
import json
from typing import List, Dict, Any, Optional, Union
from app.core.database import get_connection
from app.models.schemas import normalize_pick_status, normalize_ob_status

class PickCFService:
    def __init__(self, conn: Optional[sqlite3.Connection] = None):
        self._conn = conn

    def _get_conn(self) -> sqlite3.Connection:
        return self._conn or get_connection()

    def create_pick_order(
        self,
        shop: str,
        model: str,
        items: List[Dict[str, Any]],
        branch_tab: Optional[str] = None
    ) -> Dict[str, Any]:
        """Step 3: Create a Pick Order.

        Invariant: Always starts in 'Chờ người dùng CF' status. Zero silent auto-CF!
        """
        if not items:
            raise ValueError("Danh sách mã hàng tạo đơn pick không được rỗng")

        items = [dict(item) for item in items]
        seen_sources = set()
        for item in items:
            raw_qty = next((item[key] for key in ("req_qty", "required_qty", "qty") if key in item and item[key] is not None), None)
            if isinstance(raw_qty, bool):
                raise ValueError("Số lượng yêu cầu phải là số lớn hơn 0")
            try:
                qty = float(raw_qty)
            except (TypeError, ValueError):
                raise ValueError("Số lượng yêu cầu phải là số lớn hơn 0")
            if not math.isfinite(qty) or qty <= 0:
                raise ValueError("Số lượng yêu cầu phải hữu hạn và lớn hơn 0")
            if not str(item.get("material_code") or "").strip() or not str(item.get("ob_number") or "").strip():
                raise ValueError("Cần mã hàng và số OB thật cho mỗi dòng pick")
            item["req_qty"] = qty
            source_key = ("id", str(item["id"])) if item.get("id") is not None else ("line", str(item.get("material_code", "")).strip(), str(item.get("batch", "")).strip(), str(item.get("ob_number", "")).strip())
            if source_key in seen_sources:
                raise ValueError("Dòng OB nguồn bị lặp trong cùng đơn pick")
            seen_sources.add(source_key)

        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            today_str = datetime.date.today().strftime("%Y%m%d")
            today_iso = datetime.date.today().isoformat()
            now_iso = datetime.datetime.now().isoformat()
            branch = branch_tab or ("BODY_3001" if "BODY" in shop.upper() else "EDC_4007")
            canonical_lines = sorted((
                str(item.get("material_code", "")).strip(),
                str(item.get("batch", "")).strip(),
                str(item.get("ob_number", "")).strip(),
                float(item["req_qty"]),
                str(item.get("id", "")),
                str(item.get("sto_number", "")).strip(),
                str(item.get("model") or model).strip(),
                str(item.get("shop") or shop).strip(),
            ) for item in items)
            source_key = hashlib.sha256(json.dumps(
                [branch, shop.strip(), model.strip(), canonical_lines], ensure_ascii=False,
                separators=(",", ":"), allow_nan=False
            ).encode("utf-8")).hexdigest()

            if close_conn:
                conn.execute("BEGIN IMMEDIATE")
            with conn:
                conn.execute("UPDATE pick_orders SET id = id WHERE 0")
                has_source_key = any(col["name"] == "source_key" for col in conn.execute("PRAGMA table_info(pick_orders)"))
                if has_source_key:
                    prior = conn.execute("SELECT * FROM pick_orders WHERE source_key = ?", (source_key,)).fetchone()
                    if prior:
                        code = prior["pick_order_code"] or prior["pick_order_no"]
                        return {
                            "status": "success", "idempotent": True,
                            "id": prior["id"], "pick_order_id": prior["id"],
                            "pick_order_code": code, "pick_order_no": code,
                            **normalize_pick_status(prior["status"]),
                            "item_count": prior["item_count"], "total_qty": prior["total_qty"]
                        }
                # Generate unique sequence number
                row_cnt = conn.execute(
                    "SELECT COUNT(*) FROM pick_orders WHERE pick_order_code LIKE ?",
                    (f"PICK-{today_str}-%",)
                ).fetchone()[0]
                seq = row_cnt + 1
                order_code = f"PICK-{today_str}-{seq:03d}"
                while conn.execute("SELECT 1 FROM pick_orders WHERE pick_order_code = ?", (order_code,)).fetchone():
                    seq += 1
                    order_code = f"PICK-{today_str}-{seq:03d}"

                total_qty = sum(
                    float(it.get("req_qty") or it.get("required_qty") or it.get("qty") or 0.0)
                    for it in items
                )

                # STRICT INVARIANT: Status MUST be 'Chờ người dùng CF'
                cur = conn.execute("""
                    INSERT INTO pick_orders (
                        pick_order_code, pick_order_no, branch_tab, shop, model,
                        item_count, total_items, total_qty, status, notes, created_date, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Chờ người dùng CF', ?, ?, ?, ?)
                """, (
                    order_code, order_code, branch, shop, model,
                    len(items), len(items), total_qty,
                    "Đơn mới tạo, chờ người dùng kiểm tra thực tế và bấm Kiểm tra CF",
                    today_iso, now_iso, now_iso
                ))
                pick_order_id = cur.lastrowid
                if has_source_key:
                    conn.execute("UPDATE pick_orders SET source_key = ? WHERE id = ?", (source_key, pick_order_id))

                for it in items:
                    mat = str(it.get("material_code", "")).strip()
                    batch = str(it.get("batch", "")).strip()
                    sto = str(it.get("sto_number", "")).strip()
                    ob = str(it.get("ob_number", "")).strip()
                    qty = float(it.get("req_qty") or it.get("required_qty") or it.get("qty") or 0.0)
                    item_model = str(it.get("model") or model).strip()
                    item_shop = str(it.get("shop") or shop).strip()

                    # Upsert or update ob_items
                    raw_id = it.get("id")
                    ob_item_id = None
                    if isinstance(raw_id, int) or (isinstance(raw_id, str) and raw_id.isdigit()):
                        ob_item_id = int(raw_id)
                        existing = conn.execute("SELECT * FROM ob_items WHERE id = ?", (ob_item_id,)).fetchone()
                        if not existing:
                            raise ValueError("Không tìm thấy dòng OB nguồn")
                        if existing["branch_tab"] != branch or existing["material_code"] != mat or (existing["batch"] or "") != batch or (existing["ob_number"] or "") != ob:
                            raise ValueError("Dòng OB nguồn không khớp nhánh, mã hàng, batch hoặc số OB")
                        if existing["pick_order_id"] or existing["status"] in ("Đã CF", "USER_CONFIRMED", "CF_COMPLETED", "Đã lập BBBG", "HANDED_OVER", "READY_FOR_BBBG"):
                            raise ValueError("Dòng OB đã có đơn pick; không tạo lại")
                        norm_ex = normalize_ob_status(existing["status"])
                        if norm_ex["status_code"] not in ("READY_FOR_PICK", "WAITING_CF") and existing["rc52_code"] not in ("RC52_READY", "RC52_WAIT_PICK"):
                            raise ValueError(f"Dòng OB {mat} (ID {ob_item_id}) chưa qua kiểm tra tồn hợp lệ hoặc không ở trạng thái READY (trạng thái hiện tại: {existing['status']}); không thể tạo đơn pick")
                        conn.execute("""
                            UPDATE ob_items
                            SET status = 'Chờ pick / CF', pick_order_id = ?, updated_at = ?
                            WHERE id = ?
                        """, (pick_order_id, now_iso, ob_item_id))
                    else:
                        uid = str(raw_id) if raw_id else f"uid-{uuid.uuid4().hex[:8]}"
                        existing = conn.execute("SELECT * FROM ob_items WHERE item_uid = ?", (uid,)).fetchone()
                        if existing:
                            if existing["branch_tab"] != branch or existing["material_code"] != mat or (existing["batch"] or "") != batch or (existing["ob_number"] or "") != ob:
                                raise ValueError("Dòng OB nguồn không khớp nhánh, mã hàng, batch hoặc số OB")
                            if existing["pick_order_id"] or existing["status"] in ("Đã CF", "USER_CONFIRMED", "CF_COMPLETED", "Đã lập BBBG", "HANDED_OVER", "READY_FOR_BBBG"):
                                raise ValueError("Dòng OB đã có đơn pick; không tạo lại")
                            norm_ex = normalize_ob_status(existing["status"])
                            if norm_ex["status_code"] not in ("READY_FOR_PICK", "WAITING_CF") and existing["rc52_code"] not in ("RC52_READY", "RC52_WAIT_PICK"):
                                raise ValueError(f"Dòng OB {mat} chưa qua kiểm tra tồn hợp lệ hoặc không ở trạng thái READY (trạng thái hiện tại: {existing['status']}); không thể tạo đơn pick")
                            ob_item_id = existing["id"]
                            conn.execute("""
                                UPDATE ob_items
                                SET status = 'Chờ pick / CF', pick_order_id = ?, updated_at = ?
                                WHERE id = ?
                            """, (pick_order_id, now_iso, ob_item_id))
                        else:
                            if it.get("unverified") is True or it.get("verified") is False or it.get("status") in ("UNVERIFIED", "Chưa kiểm tra", "Thiếu SL OB", "LOCKED_SHORTAGE", "LOCKED", "SHORTAGE", "INVALID", "CANCELLED", "INVALID_CANCELLED"):
                                raise ValueError(f"Chặn nhận OB nhập tay chưa có bằng chứng kiểm chứng ({mat} / OB {ob}); chỉ cho phép tạo pick cho dòng OB đã kiểm tra tồn và có trạng thái hợp lệ/READY")
                            c_ob = conn.execute("""
                                INSERT INTO ob_items (
                                    item_uid, branch_tab, branch, material_code, model,
                                    req_qty, required_qty, batch, shop, sto_number, ob_number,
                                    status, pick_order_id, created_at, updated_at
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Chờ pick / CF', ?, ?, ?)
                            """, (
                                uid, branch, branch, mat, item_model,
                                qty, qty, batch, item_shop, sto, ob,
                                pick_order_id, now_iso, now_iso
                            ))
                            ob_item_id = c_ob.lastrowid

                    poi_cols = {r[1] for r in conn.execute("PRAGMA table_info(pick_order_items)").fetchall()}
                    if "missing_qty" in poi_cols:
                        conn.execute("""
                            INSERT INTO pick_order_items (
                                pick_order_id, pick_order_no, ob_item_id, material_code, model,
                                batch, shop, sto_number, ob_number, qty, cf_status, cf_actual_qty,
                                line_confirmed, missing_qty, cf_source
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'PENDING_CF', 0.0, 0, ?, 'MANUAL_ACTUAL')
                        """, (
                            pick_order_id, order_code, ob_item_id, mat, item_model,
                            batch, item_shop, sto, ob, qty, qty
                        ))
                    else:
                        conn.execute("""
                            INSERT INTO pick_order_items (
                                pick_order_id, pick_order_no, ob_item_id, material_code, model,
                                batch, shop, sto_number, ob_number, qty, cf_status, cf_actual_qty, line_confirmed
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'PENDING_CF', 0.0, 0)
                        """, (
                            pick_order_id, order_code, ob_item_id, mat, item_model,
                            batch, item_shop, sto, ob, qty
                        ))

                # Log creation event
                conn.execute("""
                    INSERT INTO job_events (
                        job_id, business_key, job_type, step_name, event_type,
                        target_ref, status, evidence, created_at
                    ) VALUES (?, ?, 'PICK_ORDER', 'CREATE_PICK', 'CREATE_PICK', ?, 'SUCCESS', ?, ?)
                """, (
                    str(uuid.uuid4()), f"PICK_CREATE_{order_code}", order_code,
                    f"Created pick order {order_code} with {len(items)} items, status='Chờ người dùng CF'",
                    now_iso
                ))

            return {
                "status": "success",
                "id": pick_order_id,
                "pick_order_id": pick_order_id,
                "pick_order_code": order_code,
                "pick_order_no": order_code,
                "pick_status": "WAITING_USER_CF",
                "status_code": "WAITING_USER_CF",
                "status_label": "Chờ người dùng CF",
                "item_count": len(items),
                "total_qty": total_qty
            }
        finally:
            if close_conn:
                conn.close()

    def verify_cf(
        self,
        pick_order_ref: Union[int, str],
        verified_by: str,
        actual_quantities: Optional[Dict[str, float]] = None,
        notes: Optional[str] = None
    ) -> Dict[str, Any]:
        """Step 4: Explicit Manual CF Verification Gate.

        Activated ONLY when user clicks 'Kiểm tra CF' on the UI.
        Transitions pick order status to 'Đã CF' and unlocks Step 5 (BBBG).
        """
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            now_iso = datetime.datetime.now().isoformat()
            with conn:
                # Acquire the SQLite writer lock before reading status (avoid stale CF races).
                conn.execute("UPDATE pick_orders SET id = id WHERE 0")
                if isinstance(pick_order_ref, int) or (isinstance(pick_order_ref, str) and pick_order_ref.isdigit()):
                    row = conn.execute("SELECT * FROM pick_orders WHERE id = ?", (int(pick_order_ref),)).fetchone()
                else:
                    row = conn.execute(
                        "SELECT * FROM pick_orders WHERE pick_order_code = ? OR pick_order_no = ?",
                        (str(pick_order_ref), str(pick_order_ref))
                    ).fetchone()

                if not row:
                    raise LookupError(f"Không tìm thấy đơn pick: {pick_order_ref}")

                po_id = row["id"]
                po_code = row["pick_order_code"] or row["pick_order_no"]

                if not verified_by or not verified_by.strip():
                    raise ValueError("Cần nhập tên người kiểm tra CF")
                lines = conn.execute(
                    "SELECT id, qty FROM pick_order_items WHERE pick_order_id = ? ORDER BY id", (po_id,)
                ).fetchall()
                if not lines or not isinstance(actual_quantities, dict):
                    raise ValueError("Cần nhập số lượng thực tế cho từng dòng theo ID dòng pick")
                expected_keys = {str(line["id"]) for line in lines}
                if set(actual_quantities) != expected_keys:
                    raise ValueError("Số lượng thực tế phải có đúng và đủ ID các dòng pick")
                validated = {}
                for line in lines:
                    key = str(line["id"])
                    raw = actual_quantities[key]
                    if isinstance(raw, bool) or not isinstance(raw, (int, float)):
                        raise ValueError(f"Dòng {key}: số lượng thực tế phải là số")
                    actual = float(raw)
                    requested = float(line["qty"])
                    if not math.isfinite(actual) or actual <= 0 or not math.isfinite(requested) or requested <= 0:
                        raise ValueError(f"Dòng {key}: số lượng phải hữu hạn và lớn hơn 0")
                    if actual != requested:
                        raise ValueError(f"Dòng {key}: thực tế {actual:g} chưa khớp yêu cầu {requested:g}; chưa thể CF")
                    validated[key] = actual

                if row["status"] in ("Đã CF", "USER_CONFIRMED", "CF_COMPLETED"):
                    return {
                        "status": "already_confirmed",
                        "id": po_id,
                        "pick_order_code": po_code,
                        "pick_order_no": po_code,
                        "cf_status": "Đã CF",
                        "pick_status": "USER_CONFIRMED",
                        "status_code": "USER_CONFIRMED",
                        "status_label": "Đã CF"
                    }

                if row["status"] not in ("Chờ người dùng CF", "WAITING_USER_CF", "WAITING_CF", "PARTIAL_CF", "CF một phần"):
                    raise ValueError("Chỉ đơn đang chờ người dùng CF mới được xác nhận")

                conn.execute("""
                    UPDATE pick_orders
                    SET status = 'Đã CF',
                        cf_verified_by = ?,
                        cf_verified_at = ?,
                        confirmed_by = ?,
                        confirmed_at = ?,
                        notes = COALESCE(?, notes),
                        updated_at = ?
                    WHERE id = ?
                """, (verified_by, now_iso, verified_by, now_iso, notes, now_iso, po_id))

                # Update pick_order_items
                poi_cols = {r[1] for r in conn.execute("PRAGMA table_info(pick_order_items)").fetchall()}
                if "missing_qty" in poi_cols:
                    conn.executemany("""
                        UPDATE pick_order_items
                        SET cf_status = 'CONFIRMED',
                            cf_actual_qty = ?,
                            missing_qty = 0.0,
                            line_confirmed = 1,
                            verified_by = ?,
                            verified_at = ?,
                            evidence_note = COALESCE(?, evidence_note),
                            cf_source = 'MANUAL_ACTUAL'
                        WHERE id = ? AND pick_order_id = ?
                    """, [(validated[str(line["id"])], verified_by, now_iso, notes, line["id"], po_id) for line in lines])
                else:
                    conn.executemany("""
                        UPDATE pick_order_items
                        SET cf_status = 'CONFIRMED',
                            cf_actual_qty = ?,
                            line_confirmed = 1
                        WHERE id = ? AND pick_order_id = ?
                    """, [(validated[str(line["id"])], line["id"], po_id) for line in lines])

                # Update linked ob_items
                conn.execute("""
                    UPDATE ob_items
                    SET status = 'Đã CF', updated_at = ?
                    WHERE pick_order_id = ?
                """, (now_iso, po_id))

                # Audit log
                conn.execute("""
                    INSERT INTO job_events (
                        job_id, business_key, job_type, step_name, event_type,
                        target_ref, status, evidence, created_at
                    ) VALUES (?, ?, 'MANUAL_CF', 'VERIFY_CF', 'USER_CF', ?, 'SUCCESS', ?, ?)
                """, (
                    str(uuid.uuid4()), f"CF_VERIFY_{po_code}", po_code,
                    f"Manual CF (Xác nhận thực tế thủ công/nguồn nhập) verified by {verified_by} at {now_iso}; actual quantities by line: {validated}",
                    now_iso
                ))
                if conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='cf_check_history'").fetchone():
                    full_lines = conn.execute("SELECT * FROM pick_order_items WHERE pick_order_id = ? ORDER BY id", (po_id,)).fetchall()
                    for fl in full_lines:
                        conn.execute("""
                            INSERT INTO cf_check_history (
                                pick_order_id, pick_order_item_id, ob_number, material_code, batch,
                                req_qty, actual_qty, missing_qty, reason, status, verified_by, verified_at, evidence, cf_source
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, 0.0, ?, 'CF_COMPLETED', ?, ?, ?, 'MANUAL_ACTUAL')
                        """, (
                            po_id, fl["id"], fl["ob_number"], fl["material_code"], fl["batch"],
                            float(fl["qty"]), validated[str(fl["id"])], notes, verified_by, now_iso,
                            notes or "Xác nhận thực tế đủ 100%"
                        ))

            return {
                "status": "success",
                "id": po_id,
                "pick_order_id": po_id,
                "pick_order_code": po_code,
                "pick_order_no": po_code,
                "cf_status": "Đã CF",
                "pick_status": "USER_CONFIRMED",
                "status_code": "USER_CONFIRMED",
                "status_label": "Đã CF",
                "cf_source_label": "Xác nhận thực tế thủ công/nguồn nhập (Chưa xác nhận SAP live)",
                "sap_cf_verified": False,
                "verified_by": verified_by,
                "verified_at": now_iso,
                "message": "Đã xác nhận CF thành công! Đơn pick đã đủ điều kiện chuyển sang bước 5: BBBG."
            }
        finally:
            if close_conn:
                conn.close()

    def list_pick_orders(self) -> List[Dict[str, Any]]:
        """List all pick orders sorted by newest first."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            rows = conn.execute("SELECT * FROM pick_orders ORDER BY id DESC").fetchall()
            results = []
            for r in rows:
                d = dict(r)
                d.update(normalize_pick_status(d.get("status")))
                d["cf_source_label"] = "Xác nhận thực tế thủ công/nguồn nhập (Không báo SAP đã CF)"
                d["sap_cf_verified"] = False
                p_items = conn.execute(
                    "SELECT * FROM pick_order_items WHERE pick_order_id = ? ORDER BY id ASC",
                    (d["id"],)
                ).fetchall()
                d["items"] = [dict(pi) for pi in p_items]
                obs = [str(pi["ob_number"]).strip() for pi in p_items if pi["ob_number"]]
                mats = [str(pi["material_code"]).strip() for pi in p_items if pi["material_code"]]
                racks = [str(pi["rack"]).strip() for pi in p_items if pi["rack"]]
                d["ob_summary"] = ", ".join(dict.fromkeys(obs)) or "--"
                d["material_summary"] = ", ".join(dict.fromkeys(mats)) or "--"
                d["location_summary"] = ", ".join(dict.fromkeys(racks)) or "KHO BODY"
                d["slip_no"] = d.get("pick_order_no") or d.get("pick_order_code") or str(d["id"])
                results.append(d)
            return results
        finally:
            if close_conn:
                conn.close()

    def list_waiting_cf_lines(
        self,
        ob_number: Optional[str] = None,
        pick_order_code: Optional[str] = None,
        material_code: Optional[str] = None,
        batch: Optional[str] = None,
        plant: Optional[str] = None,
        packing_box: Optional[str] = None,
        rack: Optional[str] = None,
        status: Optional[str] = None,
        include_completed: bool = False,
        is_direct_only: bool = False,
        q: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Query waiting CF / partial CF lines with filtering.

        Invariant: Completed 100% CF orders leave the waiting CF list by default.
        """
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            sql = """
                SELECT
                    poi.*,
                    po.pick_order_code as po_code,
                    po.pick_order_no as po_no,
                    po.status as po_status,
                    po.branch_tab as po_branch_tab,
                    po.created_date as po_created_date
                FROM pick_order_items poi
                LEFT JOIN pick_orders po ON poi.pick_order_id = po.id
                WHERE 1=1
            """
            params: List[Any] = []
            if not include_completed:
                sql += " AND (po.id IS NULL OR po.status NOT IN ('Đã CF', 'USER_CONFIRMED', 'CF_COMPLETED', 'Đã lập BBBG', 'HANDED_OVER', 'REJECTED', 'Từ chối'))"
                sql += " AND (poi.line_confirmed = 0 OR poi.cf_status NOT IN ('CONFIRMED', 'CF_COMPLETED'))"
            if ob_number and ob_number.strip():
                sql += " AND poi.ob_number LIKE ?"
                params.append(f"%{ob_number.strip()}%")
            if pick_order_code and pick_order_code.strip():
                sql += " AND (po.pick_order_code LIKE ? OR po.pick_order_no LIKE ? OR poi.direct_order_ref LIKE ?)"
                params.extend([f"%{pick_order_code.strip()}%", f"%{pick_order_code.strip()}%", f"%{pick_order_code.strip()}%"])
            if material_code and material_code.strip():
                sql += " AND poi.material_code LIKE ?"
                params.append(f"%{material_code.strip()}%")
            if batch and batch.strip():
                sql += " AND poi.batch LIKE ?"
                params.append(f"%{batch.strip()}%")
            if plant and plant.strip():
                sql += " AND (poi.plant LIKE ? OR po.branch_tab LIKE ?)"
                params.extend([f"%{plant.strip()}%", f"%{plant.strip()}%"])
            if packing_box and packing_box.strip():
                sql += " AND poi.packing_box LIKE ?"
                params.append(f"%{packing_box.strip()}%")
            if rack and rack.strip():
                sql += " AND poi.rack LIKE ?"
                params.append(f"%{rack.strip()}%")
            if is_direct_only:
                sql += " AND poi.is_direct_paste = 1"
            if status and status.strip():
                st = status.strip()
                if st in ("PARTIAL_CF", "CF một phần"):
                    sql += " AND (poi.cf_status = 'PARTIAL_CF' OR po.status IN ('PARTIAL_CF', 'CF một phần'))"
                elif st in ("WAITING_CF", "PENDING_CF", "Chờ CF"):
                    sql += " AND poi.cf_status IN ('PENDING_CF', 'WAITING_CF')"
                elif st in ("CF_COMPLETED", "CONFIRMED", "Đã CF đầy đủ", "Đã CF"):
                    sql += " AND poi.cf_status IN ('CONFIRMED', 'CF_COMPLETED')"

            sql += " ORDER BY poi.id ASC"
            rows = conn.execute(sql, params).fetchall()
            results = []
            q_lower = q.strip().lower() if q and q.strip() else None
            for r in rows:
                d = dict(r)
                req_qty = float(d.get("qty") or 0.0)
                act_qty = float(d.get("cf_actual_qty") or 0.0)
                miss_qty = d.get("missing_qty")
                if miss_qty is None:
                    miss_qty = max(0.0, req_qty - act_qty)
                else:
                    miss_qty = max(0.0, req_qty - act_qty)
                reason = d.get("shortage_reason") or ""
                cf_st = d.get("cf_status") or "PENDING_CF"
                if cf_st == "CONFIRMED":
                    cf_st_code = "CF_COMPLETED"
                    cf_st_label = "Đã CF đầy đủ"
                elif cf_st == "PARTIAL_CF" or (0 < act_qty < req_qty):
                    cf_st_code = "PARTIAL_CF"
                    cf_st_label = "CF một phần"
                else:
                    cf_st_code = "WAITING_CF"
                    cf_st_label = "Chờ CF"

                item_out = {
                    "id": d["id"],
                    "line_id": d["id"],
                    "pick_order_id": d["pick_order_id"],
                    "pick_order_code": d.get("po_code") or d.get("po_no") or d.get("direct_order_ref") or "Dán trực tiếp",
                    "direct_order_ref": d.get("direct_order_ref") or "",
                    "is_direct_paste": bool(d.get("is_direct_paste")),
                    "ob_item_id": d.get("ob_item_id"),
                    "ob_number": d.get("ob_number") or "",
                    "material_code": d.get("material_code") or "",
                    "plant": d.get("plant") or ("4007" if "4007" in str(d.get("po_branch_tab") or "") else "3001"),
                    "model": d.get("model") or "",
                    "shop": d.get("shop") or "",
                    "batch": d.get("batch") or "",
                    "packing_box": d.get("packing_box") or "",
                    "rack": d.get("rack") or "",
                    "sto_number": d.get("sto_number") or "",
                    "req_qty": req_qty,
                    "qty": req_qty,
                    "cf_actual_qty": act_qty,
                    "missing_qty": round(miss_qty, 6),
                    "shortage_reason": reason,
                    "reason": reason,
                    "cf_status": cf_st_code,
                    "cf_status_label": cf_st_label,
                    "line_confirmed": int(d.get("line_confirmed") or 0),
                    "verified_by": d.get("verified_by") or "",
                    "verified_at": d.get("verified_at") or "",
                    "evidence_note": d.get("evidence_note") or "",
                    "cf_source": d.get("cf_source") or "MANUAL_ACTUAL",
                    "cf_source_label": d.get("cf_source_label") or "Xác nhận thực tế thủ công/nguồn nhập (Không báo SAP đã CF)",
                    "sap_cf_verified": False,
                    "order_status": d.get("po_status") or "Người dùng quản lý (Dán trực tiếp)",
                }
                if q_lower:
                    haystack = " ".join([
                        str(item_out["pick_order_code"]),
                        str(item_out["ob_number"]),
                        str(item_out["material_code"]),
                        str(item_out["plant"]),
                        str(item_out["batch"]),
                        str(item_out["packing_box"]),
                        str(item_out["rack"]),
                        str(item_out["shortage_reason"]),
                        str(item_out["model"]),
                    ]).lower()
                    if q_lower not in haystack:
                        continue
                results.append(item_out)
            return results
        finally:
            if close_conn:
                conn.close()

    def save_line_cf(
        self,
        line_id: int,
        actual_qty: float,
        verified_by: str,
        reason: Optional[str] = None,
        evidence: Optional[str] = None
    ) -> Dict[str, Any]:
        """Save per-line CF inspection progress (supports PARTIAL_CF and CF_COMPLETED).

        Invariants:
        - Never unlocks BBBG until 100% of lines in the pick order have actual_qty == req_qty.
        - Persists partial progress across tab switches, refreshes, and restarts.
        - Idempotent on repeated identical saves.
        - Records user, timestamp, and evidence history.
        """
        if not verified_by or not verified_by.strip():
            raise ValueError("Cần nhập tên người kiểm tra thực tế")
        if isinstance(actual_qty, bool) or not isinstance(actual_qty, (int, float)):
            raise ValueError("Số lượng thực tế phải là số")
        actual = float(actual_qty)
        if not math.isfinite(actual) or actual < 0:
            raise ValueError("Số lượng thực tế phải hữu hạn và >= 0")

        verified_by = verified_by.strip()
        reason_str = (reason or "").strip()
        evidence_str = (evidence or "").strip()
        now_iso = datetime.datetime.now().isoformat()

        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            with conn:
                conn.execute("UPDATE pick_orders SET id = id WHERE 0")
                line = conn.execute("SELECT * FROM pick_order_items WHERE id = ?", (int(line_id),)).fetchone()
                if not line:
                    raise LookupError(f"Không tìm thấy dòng pick ID {line_id}")
                po_id = line["pick_order_id"]
                if po_id is not None:
                    po = conn.execute("SELECT * FROM pick_orders WHERE id = ?", (po_id,)).fetchone()
                    if not po:
                        raise LookupError(f"Không tìm thấy đơn pick {po_id}")
                    po_code = po["pick_order_code"] or po["pick_order_no"]
                    po_status = po["status"]
                else:
                    po = None
                    po_code = line["direct_order_ref"] or "Dán trực tiếp"
                    po_status = "Người dùng quản lý (Dán trực tiếp)"
                req_qty = float(line["qty"] or 0.0)

                if actual > req_qty:
                    raise ValueError(f"Dòng {line_id}: SL thực tế ({actual:g}) không được vượt quá SL yêu cầu ({req_qty:g})")

                missing = round(max(0.0, req_qty - actual), 6)
                if missing > 0 and not reason_str:
                    raise ValueError(f"Dòng {line_id}: Còn thiếu {missing:g} PCS so với yêu cầu ({req_qty:g}); vui lòng nhập lý do thiếu")

                if po and po_status in ("Đã CF", "USER_CONFIRMED", "CF_COMPLETED", "Đã lập BBBG", "HANDED_OVER"):
                    if abs(float(line["cf_actual_qty"] or 0.0) - actual) < 1e-6:
                        return {
                            "status": "already_confirmed",
                            "idempotent": True,
                            "line_id": line_id,
                            "pick_order_id": po_id,
                            "pick_order_code": po_code,
                            "cf_status": "CF_COMPLETED",
                            "order_status": po["status"],
                            "bbbg_unlocked": True
                        }
                    raise ValueError("Đơn pick đã xác nhận CF đủ 100%; không thể sửa lùi")

                if po and po_status in ("REJECTED", "Từ chối"):
                    raise ValueError("Đơn pick đã bị từ chối; không thể xác nhận CF")

                is_line_full = abs(actual - req_qty) < 1e-6 and req_qty > 0
                if is_line_full:
                    line_cf_status = "CONFIRMED"
                    line_cf_code = "CF_COMPLETED"
                    line_confirmed = 1
                    ob_status = "Đã CF"
                elif actual > 0 or reason_str:
                    line_cf_status = "PARTIAL_CF"
                    line_cf_code = "PARTIAL_CF"
                    line_confirmed = 0
                    ob_status = "PARTIAL_CF"
                else:
                    line_cf_status = "PENDING_CF"
                    line_cf_code = "WAITING_CF"
                    line_confirmed = 0
                    ob_status = "Chờ pick / CF"

                # Idempotency check if identical state already saved
                prev_actual = float(line["cf_actual_qty"] or 0.0)
                prev_reason = (line["shortage_reason"] or "") if "shortage_reason" in line.keys() else ""
                prev_status = line["cf_status"] or "PENDING_CF"
                if abs(prev_actual - actual) < 1e-6 and prev_reason == reason_str and prev_status == line_cf_status:
                    all_lines = conn.execute("SELECT line_confirmed FROM pick_order_items WHERE pick_order_id = ?", (po_id,)).fetchall() if po_id else []
                    all_done = all(int(l["line_confirmed"] or 0) == 1 for l in all_lines) if all_lines else is_line_full
                    return {
                        "status": "success",
                        "idempotent": True,
                        "line_id": line_id,
                        "pick_order_id": po_id,
                        "pick_order_code": po_code,
                        "req_qty": req_qty,
                        "cf_actual_qty": actual,
                        "missing_qty": missing,
                        "shortage_reason": reason_str,
                        "cf_status": line_cf_code,
                        "order_status": po_status,
                        "bbbg_unlocked": bool(all_done and po),
                        "verified_by": verified_by,
                        "verified_at": line["verified_at"] if "verified_at" in line.keys() else now_iso,
                        "cf_source_label": "Xác nhận thực tế thủ công/nguồn nhập (Không báo SAP đã CF)",
                    }

                conn.execute("""
                    UPDATE pick_order_items
                    SET cf_status = ?,
                        cf_actual_qty = ?,
                        missing_qty = ?,
                        shortage_reason = ?,
                        line_confirmed = ?,
                        verified_by = ?,
                        verified_at = ?,
                        evidence_note = ?,
                        cf_source = 'MANUAL_ACTUAL'
                    WHERE id = ?
                """, (
                    line_cf_status, actual, missing, reason_str, line_confirmed,
                    verified_by, now_iso, evidence_str or "Kiểm tra thực tế tại kho/line",
                    line_id
                ))

                if line["ob_item_id"]:
                    conn.execute("""
                        UPDATE ob_items
                        SET status = ?, missing_qty = ?, note = COALESCE(?, note), updated_at = ?
                        WHERE id = ?
                    """, (ob_status, missing, reason_str or None, now_iso, line["ob_item_id"]))

                if po_id is not None:
                    # Evaluate all lines of the pick order
                    all_lines = conn.execute(
                        "SELECT id, qty, cf_actual_qty, line_confirmed FROM pick_order_items WHERE pick_order_id = ? ORDER BY id",
                        (po_id,)
                    ).fetchall()
                    all_done = len(all_lines) > 0 and all(
                        int(l["line_confirmed"] or 0) == 1 and abs(float(l["cf_actual_qty"] or 0.0) - float(l["qty"] or 0.0)) < 1e-6
                        for l in all_lines
                    )

                    if all_done:
                        new_po_status = "Đã CF"
                        conn.execute("""
                            UPDATE pick_orders
                            SET status = 'Đã CF',
                                cf_verified_by = ?,
                                cf_verified_at = ?,
                                confirmed_by = ?,
                                confirmed_at = ?,
                                updated_at = ?
                            WHERE id = ?
                        """, (verified_by, now_iso, verified_by, now_iso, now_iso, po_id))
                        conn.execute("""
                            UPDATE ob_items
                            SET status = 'Đã CF', missing_qty = 0.0, updated_at = ?
                            WHERE pick_order_id = ?
                        """, (now_iso, po_id))
                    else:
                        any_progress = any(float(l["cf_actual_qty"] or 0.0) > 0 or int(l["line_confirmed"] or 0) == 1 for l in all_lines) or bool(reason_str)
                        new_po_status = "PARTIAL_CF" if any_progress else "Chờ người dùng CF"
                        conn.execute("""
                            UPDATE pick_orders
                            SET status = ?,
                                updated_at = ?
                            WHERE id = ?
                        """, (new_po_status, now_iso, po_id))
                    bbbg_unlocked = all_done
                    left_waiting = all_done
                else:
                    new_po_status = "Dán trực tiếp"
                    bbbg_unlocked = False
                    left_waiting = is_line_full

                # Record CF history and job_events
                if conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='cf_check_history'").fetchone():
                    conn.execute("""
                        INSERT INTO cf_check_history (
                            pick_order_id, pick_order_item_id, ob_number, material_code, batch,
                            req_qty, actual_qty, missing_qty, reason, status, verified_by, verified_at, evidence, cf_source
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'MANUAL_ACTUAL')
                    """, (
                        po_id or 0, line_id, line["ob_number"], line["material_code"], line["batch"],
                        req_qty, actual, missing, reason_str, line_cf_code,
                        verified_by, now_iso, evidence_str or "Kiểm tra thực tế thủ công"
                    ))

                conn.execute("""
                    INSERT INTO job_events (
                        job_id, business_key, job_type, step_name, event_type,
                        target_ref, status, evidence, created_at
                    ) VALUES (?, ?, 'MANUAL_CF', ?, 'LINE_CF_SAVE', ?, 'SUCCESS', ?, ?)
                """, (
                    str(uuid.uuid4()),
                    f"CF_LINE_{po_code}_{line_id}",
                    "CF_COMPLETED" if (all_done if po_id else is_line_full) else "PARTIAL_CF",
                    po_code,
                    f"CF Line {line_id} (OB {line['ob_number']}, Mat {line['material_code']}): req={req_qty:g}, actual={actual:g}, missing={missing:g}, reason='{reason_str}', by={verified_by}, source=MANUAL_ACTUAL",
                    now_iso
                ))

            return {
                "status": "success",
                "idempotent": False,
                "line_id": line_id,
                "pick_order_id": po_id,
                "pick_order_code": po_code,
                "ob_number": line["ob_number"],
                "material_code": line["material_code"],
                "batch": line["batch"],
                "req_qty": req_qty,
                "cf_actual_qty": actual,
                "missing_qty": missing,
                "shortage_reason": reason_str,
                "cf_status": line_cf_code,
                "order_status": new_po_status,
                "bbbg_unlocked": bbbg_unlocked,
                "left_waiting_cf": left_waiting,
                "verified_by": verified_by,
                "verified_at": now_iso,
                "cf_source_label": "Xác nhận thực tế thủ công/nguồn nhập (Không báo SAP đã CF)",
                "sap_cf_verified": False,
            }
        finally:
            if close_conn:
                conn.close()

    def list_cf_history(self, pick_order_id: Optional[int] = None, limit: int = 100) -> List[Dict[str, Any]]:
        """List CF check history entries with operator, timestamp, and evidence."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            if not conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='cf_check_history'").fetchone():
                return []
            if pick_order_id:
                rows = conn.execute(
                    "SELECT * FROM cf_check_history WHERE pick_order_id = ? ORDER BY id DESC LIMIT ?",
                    (int(pick_order_id), limit)
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM cf_check_history ORDER BY id DESC LIMIT ?",
                    (limit,)
                ).fetchall()
            return [dict(r) for r in rows]
        finally:
            if close_conn:
                conn.close()

    def paste_excel_waiting_cf(
        self,
        raw_text: Optional[str] = None,
        items: Optional[List[Dict[str, Any]]] = None,
        allow_duplicates: bool = False,
        direct_order_ref: Optional[str] = None
    ) -> Dict[str, Any]:
        """Paste items directly from Excel clipboard or structured list into Chờ CF.

        Strict invariants per USER_SCOPE_LOW_STOCK_20261013:
        1. Does NOT require prior pick orders.
        2. Preserves leading zeroes for string codes (OB, material, plant, batch, box, rack).
        3. Real DB anti-duplicate check (both in-batch and against active uncompleted DB entries).
        4. No auto-calculation of packing box / rack; user managed.
        5. Does not auto-recognize SAP CF; starts in PENDING_CF / Chờ CF.
        """
        parsed_items: List[Dict[str, Any]] = []

        if raw_text and raw_text.strip():
            lines = [l.strip() for l in raw_text.strip().splitlines() if l.strip()]
            header_idx = {}
            data_start = 0
            if lines:
                first_row = [c.strip() for c in (lines[0].split("\t") if "\t" in lines[0] else lines[0].split(","))]
                first_lower = [c.lower() for c in first_row]
                def _is_num(val: str) -> bool:
                    try:
                        float(val)
                        return True
                    except ValueError:
                        return False

                has_numeric = any(_is_num(c) for c in first_row if c)
                header_keywords = {"ob", "so ob", "số ob", "mã", "mã hàng", "material", "plant", "batch", "số lượng", "sl", "qty", "box", "rack", "packing box", "packing_box"}
                matches = sum(1 for col in first_lower if any(col == k or col.startswith(k + " ") or col.endswith(" " + k) for k in header_keywords))

                if not has_numeric and matches >= 2:
                    data_start = 1
                    for idx, col in enumerate(first_lower):
                        if any(k in col for k in ["ob", "so ob"]):
                            header_idx["ob_number"] = idx
                        elif any(k in col for k in ["mã", "mat"]):
                            header_idx["material_code"] = idx
                        elif "plant" in col:
                            header_idx["plant"] = idx
                        elif "batch" in col:
                            header_idx["batch"] = idx
                        elif any(k in col for k in ["số lượng", "qty", "sl"]):
                            header_idx["qty"] = idx
                        elif any(k in col for k in ["box", "pack"]):
                            header_idx["packing_box"] = idx
                        elif "rack" in col:
                            header_idx["rack"] = idx

            for line_str in lines[data_start:]:
                parts = [p.strip() for p in (line_str.split("\t") if "\t" in line_str else line_str.split(","))]
                if not parts or not any(parts):
                    continue
                if header_idx:
                    row_dict = {
                        "ob_number": parts[header_idx["ob_number"]] if "ob_number" in header_idx and header_idx["ob_number"] < len(parts) else "",
                        "material_code": parts[header_idx["material_code"]] if "material_code" in header_idx and header_idx["material_code"] < len(parts) else "",
                        "plant": parts[header_idx["plant"]] if "plant" in header_idx and header_idx["plant"] < len(parts) else "3001",
                        "batch": parts[header_idx["batch"]] if "batch" in header_idx and header_idx["batch"] < len(parts) else "",
                        "qty": parts[header_idx["qty"]] if "qty" in header_idx and header_idx["qty"] < len(parts) else 0.0,
                        "packing_box": parts[header_idx["packing_box"]] if "packing_box" in header_idx and header_idx["packing_box"] < len(parts) else "",
                        "rack": parts[header_idx["rack"]] if "rack" in header_idx and header_idx["rack"] < len(parts) else "",
                    }
                else:
                    if len(parts) >= 7:
                        row_dict = {
                            "ob_number": parts[0], "material_code": parts[1], "plant": parts[2],
                            "batch": parts[3], "qty": parts[4], "packing_box": parts[5], "rack": parts[6]
                        }
                    elif len(parts) == 6:
                        row_dict = {
                            "ob_number": parts[0], "material_code": parts[1], "plant": parts[2],
                            "batch": parts[3], "qty": parts[4], "packing_box": parts[5], "rack": ""
                        }
                    elif len(parts) == 5:
                        row_dict = {
                            "ob_number": parts[0], "material_code": parts[1], "plant": parts[2],
                            "batch": parts[3], "qty": parts[4], "packing_box": "", "rack": ""
                        }
                    elif len(parts) == 4:
                        row_dict = {
                            "ob_number": parts[0], "material_code": parts[1], "plant": "3001",
                            "batch": parts[2], "qty": parts[3], "packing_box": "", "rack": ""
                        }
                    else:
                        row_dict = {
                            "ob_number": parts[0] if len(parts) > 0 else "",
                            "material_code": parts[1] if len(parts) > 1 else "",
                            "plant": "3001",
                            "batch": parts[2] if len(parts) > 2 else "",
                            "qty": parts[3] if len(parts) > 3 else 0.0,
                            "packing_box": "",
                            "rack": ""
                        }
                parsed_items.append(row_dict)

        if items:
            for it in items:
                parsed_items.append(dict(it))

        if not parsed_items:
            raise ValueError("Không có dữ liệu dán hợp lệ để nhập vào Chờ CF")

        cleaned_items = []
        seen_keys = set()
        for idx, row in enumerate(parsed_items, 1):
            mat = str(row.get("material_code") or "").strip()
            if not mat:
                raise ValueError(f"Dòng {idx}: Thiếu mã hàng")

            raw_qty = row.get("qty") if "qty" in row else row.get("req_qty")
            if isinstance(raw_qty, bool):
                raise ValueError(f"Dòng {idx}: Số lượng không được là boolean")
            try:
                qty = float(raw_qty)
            except (TypeError, ValueError):
                raise ValueError(f"Dòng {idx}: Số lượng '{raw_qty}' không hợp lệ")
            if not math.isfinite(qty) or qty <= 0:
                raise ValueError(f"Dòng {idx}: Số lượng phải hữu hạn và > 0")

            ob = str(row.get("ob_number") or "").strip()
            plant = str(row.get("plant") or "3001").strip()
            batch = str(row.get("batch") or "").strip()
            box = str(row.get("packing_box") or "").strip()
            rack = str(row.get("rack") or "").strip()
            shop = str(row.get("shop") or ("BODY" if "3001" in plant or "BODY" in str(row.get("branch", "")) else "EDC")).strip()
            model = str(row.get("model") or "").strip()

            # Anti-duplicate within payload:
            dedup_key = f"{ob}|{mat}|{plant}|{batch}"
            if dedup_key in seen_keys and not allow_duplicates:
                raise ValueError(f"Phát hiện dòng dán trùng trong danh sách Excel: OB '{ob}', mã '{mat}', plant '{plant}', batch '{batch}'")
            seen_keys.add(dedup_key)

            cleaned_items.append({
                "ob_number": ob,
                "material_code": mat,
                "plant": plant,
                "batch": batch,
                "qty": qty,
                "packing_box": box,
                "rack": rack,
                "shop": shop,
                "model": model,
                "dedup_key": f"{dedup_key}|{box}|{rack}"
            })

        conn = self._get_conn()
        close_conn = self._conn is None
        today_str = datetime.date.today().strftime("%Y%m%d")
        now_iso = datetime.datetime.now().isoformat()
        saved_lines = []

        try:
            if close_conn:
                conn.execute("BEGIN IMMEDIATE")
            with conn:
                ref = direct_order_ref or f"EXCEL-PASTE-{today_str}-{uuid.uuid4().hex[:6].upper()}"
                for it in cleaned_items:
                    # Check active duplicate in DB:
                    if not allow_duplicates:
                        existing = conn.execute("""
                            SELECT id, ob_number, material_code, plant, batch FROM pick_order_items
                            WHERE ob_number = ? AND material_code = ? AND plant = ? AND batch = ?
                              AND (cf_status != 'CF_COMPLETED' AND line_confirmed = 0)
                        """, (it["ob_number"], it["material_code"], it["plant"], it["batch"])).fetchone()
                        if existing:
                            raise ValueError(
                                f"Dòng đã tồn tại trong danh sách Chờ CF (chống dán trùng): "
                                f"OB '{it['ob_number']}', mã hàng '{it['material_code']}', plant '{it['plant']}', batch '{it['batch']}'"
                            )

                    cur = conn.execute("""
                        INSERT INTO pick_order_items (
                            pick_order_id, direct_order_ref, is_direct_paste,
                            ob_number, material_code, plant, batch, qty,
                            packing_box, rack, shop, model,
                            cf_status, cf_actual_qty, line_confirmed, missing_qty,
                            cf_source, dedup_key, created_at, updated_at
                        ) VALUES (
                            NULL, ?, 1,
                            ?, ?, ?, ?, ?,
                            ?, ?, ?, ?,
                            'PENDING_CF', 0, 0, ?,
                            'EXCEL_DIRECT', ?, ?, ?
                        )
                    """, (
                        ref,
                        it["ob_number"], it["material_code"], it["plant"], it["batch"], it["qty"],
                        it["packing_box"], it["rack"], it["shop"], it["model"],
                        it["qty"],
                        it["dedup_key"], now_iso, now_iso
                    ))
                    line_id = cur.lastrowid
                    saved_lines.append({
                        "id": line_id,
                        "line_id": line_id,
                        "direct_order_ref": ref,
                        **it,
                        "cf_status": "WAITING_CF",
                        "status_label": "Chờ CF",
                        "is_direct_paste": True
                    })
            return {
                "status": "success",
                "direct_order_ref": ref,
                "count": len(saved_lines),
                "lines": saved_lines
            }
        finally:
            if close_conn:
                conn.close()

    def update_waiting_cf_line(
        self,
        line_id: int,
        ob_number: Optional[str] = None,
        material_code: Optional[str] = None,
        plant: Optional[str] = None,
        batch: Optional[str] = None,
        qty: Optional[float] = None,
        packing_box: Optional[str] = None,
        rack: Optional[str] = None
    ) -> Dict[str, Any]:
        """Update fields of a waiting CF line before it is confirmed."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            with conn:
                line = conn.execute("SELECT * FROM pick_order_items WHERE id = ?", (int(line_id),)).fetchone()
                if not line:
                    raise LookupError(f"Không tìm thấy dòng Chờ CF ID {line_id}")
                if line["line_confirmed"] == 1 or line["cf_status"] in ("CONFIRMED", "CF_COMPLETED"):
                    raise ValueError(f"Dòng {line_id} đã hoàn tất CF; không sửa sau khi xác nhận")

                cur_qty = float(line["qty"] or 0.0)
                new_qty = cur_qty
                if qty is not None:
                    if isinstance(qty, bool):
                        raise ValueError("Số lượng không được là boolean")
                    f_qty = float(qty)
                    if not math.isfinite(f_qty) or f_qty <= 0:
                        raise ValueError("Số lượng phải hữu hạn và > 0")
                    new_qty = f_qty

                new_ob = str(ob_number).strip() if ob_number is not None else (line["ob_number"] or "")
                new_mat = str(material_code).strip() if material_code is not None else (line["material_code"] or "")
                if not new_mat:
                    raise ValueError("Mã hàng không được rỗng")
                new_plant = str(plant).strip() if plant is not None else (line["plant"] or "3001")
                new_batch = str(batch).strip() if batch is not None else (line["batch"] or "")
                new_box = str(packing_box).strip() if packing_box is not None else (line["packing_box"] or "")
                new_rack = str(rack).strip() if rack is not None else (line["rack"] or "")

                act_qty = float(line["cf_actual_qty"] or 0.0)
                new_missing = round(max(0.0, new_qty - act_qty), 6)
                now_iso = datetime.datetime.now().isoformat()
                dedup_key = f"{new_ob}|{new_mat}|{new_plant}|{new_batch}|{new_box}|{new_rack}"

                conn.execute("""
                    UPDATE pick_order_items
                    SET ob_number = ?, material_code = ?, plant = ?, batch = ?, qty = ?,
                        packing_box = ?, rack = ?, missing_qty = ?, dedup_key = ?, updated_at = ?
                    WHERE id = ?
                """, (new_ob, new_mat, new_plant, new_batch, new_qty, new_box, new_rack, new_missing, dedup_key, now_iso, line_id))

                return {
                    "status": "success",
                    "line_id": line_id,
                    "ob_number": new_ob,
                    "material_code": new_mat,
                    "plant": new_plant,
                    "batch": new_batch,
                    "qty": new_qty,
                    "packing_box": new_box,
                    "rack": new_rack,
                    "missing_qty": new_missing
                }
        finally:
            if close_conn:
                conn.close()

    def delete_waiting_cf_line(self, line_id: int) -> Dict[str, Any]:
        """Delete an unconfirmed line from Chờ CF."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            with conn:
                line = conn.execute("SELECT * FROM pick_order_items WHERE id = ?", (int(line_id),)).fetchone()
                if not line:
                    raise LookupError(f"Không tìm thấy dòng Chờ CF ID {line_id}")
                if line["line_confirmed"] == 1 or line["cf_status"] in ("CONFIRMED", "CF_COMPLETED"):
                    raise ValueError(f"Dòng {line_id} đã hoàn tất CF; không thể xóa")
                conn.execute("DELETE FROM pick_order_items WHERE id = ?", (int(line_id),))
                return {"status": "success", "deleted_line_id": line_id}
        finally:
            if close_conn:
                conn.close()
