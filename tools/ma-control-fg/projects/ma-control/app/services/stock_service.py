# -*- coding: utf-8 -*-
"""Stock FG V3.6 Service.

Implements multi-zone inventory management, safety thresholds, pull-line warnings,
and priority dispatching rules without double-counting between zones 3093, 3000, 3001, 1003, 3003.
Preserves leading zeroes for material codes and batches.
"""

import sqlite3
import datetime
import math
import json
import uuid
import os
import tempfile
from pathlib import Path
from typing import List, Dict, Any, Optional
from app.core.database import get_connection
from app.services.sap_connection import probe_connection
from app.services.source_import import parse_source

ZONE_MAPPINGS = {
    "1003": "Kho Body (Raw Materials / Stamping)",
    "3003": "Line Body (Sub-assembly / Welding)",
    "3093": "Kho MA (Main Assembly Warehouse)",
    "3000": "Line MA (Main Assembly Line Buffer)",
    "3001": "In-Transit (Vùng chuyển tiếp FG/OB)",
    "4004": "Kho EDC (Electrical / Electronics)",
    "4007": "In-Transit EDC STO"
}

SAFETY_LINE_THRESHOLD = 50.0  # Alert if line inventory < 50 PCS

ACTIVE_STOCK_FILTER = """snapshot_id = (
    SELECT MAX(sc.snapshot_id) FROM stock_snapshot_scopes sc
    WHERE sc.plant = stock_items.plant AND sc.sloc = stock_items.sloc
)"""


def stock_quantity(item):
    raw = item["unrestricted_qty"] if "unrestricted_qty" in item else item.get("qty", 0)
    if isinstance(raw, bool):
        raise ValueError("Số lượng tồn không được là true/false")
    qty = float(raw)
    if not math.isfinite(qty) or qty < 0:
        raise ValueError("Số lượng tồn phải hữu hạn và không âm")
    return qty

class StockService:
    def __init__(self, conn: Optional[sqlite3.Connection] = None):
        self._conn = conn

    def _get_conn(self) -> sqlite3.Connection:
        return self._conn or get_connection()

    def create_snapshot(
        self,
        source_type: str,
        plant: str,
        sloc: str,
        items: List[Dict[str, Any]],
        notes: Optional[str] = None
    ) -> Dict[str, Any]:
        """Ingest a stock snapshot with strict zone isolation and pull-line alerts."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            now = datetime.datetime.now().isoformat()
            total_qty = sum(stock_quantity(it) for it in items)
            unique_mats = len({str(it.get("material_code", "")).strip() for it in items if it.get("material_code")})

            # Validate complete input before committing the snapshot.
            if any(not str(it.get("material_code", "")).strip() for it in items):
                raise ValueError("Snapshot chứa dòng thiếu mã hàng")
            keys = [(str(it.get("plant") or plant), str(it.get("sloc") or sloc), str(it.get("material_code") or ""), str(it.get("batch") or ""), str(it.get("stock_type") or "UNRESTRICTED")) for it in items]
            if len(keys) != len(set(keys)):
                raise ValueError("Snapshot có dòng trùng Plant/SLoc/mã hàng/batch/stock type")
            with conn:
                cur = conn.execute("""
                    INSERT INTO stock_snapshots (source_type, snapshot_time, plant, sloc, total_rows, total_materials, total_qty, notes)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (source_type, now, plant, sloc, len(items), unique_mats, total_qty, notes or ""))
                snapshot_id = cur.lastrowid
                scopes = {(str(it.get("plant") or plant).strip(), str(it.get("sloc") or sloc).strip()) for it in items}
                if not scopes:
                    scopes = {(str(plant).strip(), str(sloc).strip())}
                conn.executemany("INSERT INTO stock_snapshot_scopes(snapshot_id, plant, sloc) VALUES (?, ?, ?)", [(snapshot_id, p, s) for p, s in scopes])

                # Process each item
                for idx, it in enumerate(items, 1):
                    mat = str(it.get("material_code", "")).strip()
                    batch = str(it.get("batch", "")).strip()
                    i_plant = str(it.get("plant") or plant).strip()
                    i_sloc = str(it.get("sloc") or sloc).strip()
                    stock_type = str(it.get("stock_type") or "UNRESTRICTED").strip()
                    qty = stock_quantity(it)
                    unit = str(it.get("unit") or "PCS").strip()

                    # Pull-line warning rule:
                    # If inventory in buffer line (3000 or 3003) is below threshold, flag pull_line_warning
                    pull_line_warn = None
                    if i_sloc in ("3000", "3003") and qty < SAFETY_LINE_THRESHOLD:
                        pull_line_warn = f"Cảnh báo kéo line: Tồn line {i_sloc} ({qty:.1f} {unit}) dưới ngưỡng an toàn {SAFETY_LINE_THRESHOLD} PCS"

                    # Priority rank calculation (FIFO based on batch date if available)
                    priority_rank = 100
                    if batch and len(batch) >= 6 and batch.isdigit():
                        priority_rank = int(batch[-4:])  # earlier batch gets lower priority number (higher dispatch priority)

                    priority_flag = "PRIORITY_OUT" if priority_rank < 50 else ("PULL_LINE" if pull_line_warn else "NORMAL")

                    conn.execute("""
                        INSERT INTO stock_items (
                            snapshot_id, plant, sloc, material_code, batch, stock_type,
                            unrestricted_qty, qty, unit, priority_rank, priority_flag, pull_line_warning
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        snapshot_id, i_plant, i_sloc, mat, batch, stock_type,
                        qty, qty, unit, priority_rank, priority_flag, pull_line_warn
                    ))

            return {
                "snapshot_id": snapshot_id,
                "total_rows": len(items),
                "total_materials": unique_mats,
                "total_qty": total_qty,
                "status": "success"
            }
        finally:
            if close_conn:
                conn.close()

    # -------------------------------------------------------------
    # LOW STOCK CONFIGURATION & THRESHOLDS
    # -------------------------------------------------------------
    def get_low_stock_config(self) -> Dict[str, Any]:
        """Get low stock thresholds and freshness parameters."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            cfg_row = conn.execute("SELECT value_json FROM settings WHERE key = 'low_stock_config'").fetchone()
            cfg = json.loads(cfg_row["value_json"]) if cfg_row and cfg_row["value_json"] else {}
            default_thresh = float(cfg.get("default_threshold", 50.0))
            max_age = float(cfg.get("max_age_hours", 24.0))

            rows = conn.execute("SELECT material_code, plant, threshold_qty, warning_message FROM stock_thresholds ORDER BY material_code, plant").fetchall()
            return {
                "default_threshold": default_thresh,
                "max_age_hours": max_age,
                "thresholds": [dict(r) for r in rows]
            }
        finally:
            if close_conn:
                conn.close()

    def update_low_stock_config(
        self,
        default_threshold: Optional[float] = None,
        max_age_hours: Optional[float] = None,
        thresholds: Optional[List[Dict[str, Any]]] = None
    ) -> Dict[str, Any]:
        """Update system-wide or per-material low stock thresholds."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            with conn:
                cfg_row = conn.execute("SELECT value_json FROM settings WHERE key = 'low_stock_config'").fetchone()
                cfg = json.loads(cfg_row["value_json"]) if cfg_row and cfg_row["value_json"] else {}
                if default_threshold is not None:
                    if isinstance(default_threshold, bool) or default_threshold < 0:
                        raise ValueError("Ngưỡng mặc định không hợp lệ")
                    cfg["default_threshold"] = float(default_threshold)
                if max_age_hours is not None:
                    if isinstance(max_age_hours, bool) or max_age_hours <= 0:
                        raise ValueError("Thời hạn nguồn (max_age_hours) phải > 0")
                    cfg["max_age_hours"] = float(max_age_hours)

                conn.execute("""
                    INSERT INTO settings (key, value_json, version, updated_at)
                    VALUES ('low_stock_config', ?, 1, CURRENT_TIMESTAMP)
                    ON CONFLICT(key) DO UPDATE SET value_json = excluded.value_json, updated_at = CURRENT_TIMESTAMP
                """, (json.dumps(cfg),))

                if thresholds is not None:
                    for t in thresholds:
                        mat = str(t.get("material_code") or "").strip()
                        if not mat:
                            continue
                        plt = str(t.get("plant") or "").strip() or None
                        t_qty = float(t.get("threshold_qty", 50.0))
                        msg = str(t.get("warning_message") or "").strip()
                        conn.execute("""
                            INSERT INTO stock_thresholds (material_code, plant, threshold_qty, warning_message, updated_at)
                            VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
                            ON CONFLICT(material_code, plant) DO UPDATE SET threshold_qty = excluded.threshold_qty, warning_message = excluded.warning_message, updated_at = CURRENT_TIMESTAMP
                        """, (mat, plt, t_qty, msg))

            return self.get_low_stock_config()
        finally:
            if close_conn:
                conn.close()

    def get_threshold(self, material_code: str, plant: Optional[str] = None) -> float:
        """Get safety threshold for a specific material and plant."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            mat = str(material_code).strip()
            plt = str(plant).strip() if plant else None
            if plt:
                row = conn.execute(
                    "SELECT threshold_qty FROM stock_thresholds WHERE material_code = ? AND plant = ?",
                    (mat, plt)
                ).fetchone()
                if row:
                    return float(row["threshold_qty"])

            row_any = conn.execute(
                "SELECT threshold_qty FROM stock_thresholds WHERE material_code = ? AND (plant IS NULL OR plant = '')",
                (mat,)
            ).fetchone()
            if row_any:
                return float(row_any["threshold_qty"])

            cfg_row = conn.execute("SELECT value_json FROM settings WHERE key = 'low_stock_config'").fetchone()
            if cfg_row and cfg_row["value_json"]:
                cfg = json.loads(cfg_row["value_json"])
                return float(cfg.get("default_threshold", 50.0))
            return 50.0
        finally:
            if close_conn:
                conn.close()

    # -------------------------------------------------------------
    # SOURCE FRESHNESS & COMPLETENESS CHECK
    # -------------------------------------------------------------
    def check_stock_freshness(
        self,
        plant: Optional[str] = None,
        sloc: Optional[str] = None,
        max_age_hours: Optional[float] = None
    ) -> Dict[str, Any]:
        """Check freshness and completeness of the latest stock snapshot."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            cfg = self.get_low_stock_config()
            limit_age = float(max_age_hours or cfg.get("max_age_hours", 24.0))

            sql = "SELECT * FROM stock_snapshots WHERE 1=1"
            params = []
            if plant:
                sql += " AND plant = ?"
                params.append(plant)
            if sloc:
                sql += " AND sloc = ?"
                params.append(sloc)
            sql += " ORDER BY id DESC LIMIT 1"

            snap = conn.execute(sql, params).fetchone()
            if not snap:
                return {
                    "is_fresh": False,
                    "is_stale": True,
                    "is_complete": False,
                    "snapshot_id": None,
                    "snapshot_time": None,
                    "age_hours": None,
                    "freshness_label": "Chưa có snapshot tồn",
                    "warning": f"Chưa có snapshot tồn cho Plant '{plant or 'all'}' / SLoc '{sloc or 'all'}'; cần tải snapshot mới",
                }

            snap_time_str = snap["snapshot_time"]
            age_hours = 0.0
            try:
                dt = datetime.datetime.fromisoformat(snap_time_str)
                age_hours = round(max(0.0, (datetime.datetime.now() - dt).total_seconds() / 3600.0), 2)
            except Exception:
                age_hours = 0.0

            is_stale = age_hours > limit_age
            is_complete = int(snap["total_rows"] or 0) > 0
            if is_stale:
                warning = f"Dữ liệu nguồn snapshot đã cũ ({age_hours} giờ, vượt ngưỡng {limit_age} giờ); cần cập nhật snapshot mới"
                fresh_label = f"Nguồn cũ ({age_hours}h)"
            elif not is_complete:
                warning = "Snapshot tồn không có dòng dữ liệu nào"
                fresh_label = "Nguồn rỗng"
            else:
                warning = None
                fresh_label = f"Nguồn mới ({age_hours}h)"

            return {
                "is_fresh": not is_stale and is_complete,
                "is_stale": is_stale,
                "is_complete": is_complete,
                "snapshot_id": snap["id"],
                "snapshot_time": snap_time_str,
                "age_hours": age_hours,
                "max_age_hours": limit_age,
                "freshness_label": fresh_label,
                "warning": warning,
                "plant": snap["plant"],
                "sloc": snap["sloc"],
                "total_rows": snap["total_rows"],
                "total_qty": snap["total_qty"]
            }
        finally:
            if close_conn:
                conn.close()

    # -------------------------------------------------------------
    # LOW STOCK ALERTS QUERY
    # -------------------------------------------------------------
    def list_low_stock_alerts(
        self,
        plant: Optional[str] = None,
        material_code: Optional[str] = None,
        scope_sloc: Optional[str] = None,
        shop: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """List materials currently below safety stock threshold without zone mixing."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            sql = f"""
                SELECT plant, material_code, SUM(unrestricted_qty) as total_unrestricted
                FROM stock_items
                WHERE stock_type = 'UNRESTRICTED' AND {ACTIVE_STOCK_FILTER}
            """
            params: List[Any] = []
            if plant:
                sql += " AND plant = ?"
                params.append(plant)
            if scope_sloc and scope_sloc.strip():
                sql += " AND sloc = ?"
                params.append(scope_sloc.strip())
            if material_code:
                sql += " AND material_code = ?"
                params.append(str(material_code).strip())
            sql += " GROUP BY plant, material_code ORDER BY plant, material_code"

            rows = conn.execute(sql, params).fetchall()
            alerts = []
            for r in rows:
                mat = r["material_code"]
                plt = r["plant"]
                tot = float(r["total_unrestricted"] or 0.0)
                thresh = self.get_threshold(mat, plt)
                if tot < thresh:
                    freshness = self.check_stock_freshness(plant=plt)
                    b_rows = conn.execute(f"""
                        SELECT batch, sloc, unrestricted_qty, priority_rank, priority_flag, stock_type
                        FROM stock_items
                        WHERE material_code = ? AND plant = ? AND {ACTIVE_STOCK_FILTER}
                        ORDER BY priority_rank ASC, batch ASC
                    """, (mat, plt)).fetchall()

                    alerts.append({
                        "material_code": mat,
                        "plant": plt,
                        "current_stock": tot,
                        "threshold_qty": thresh,
                        "shortage_to_threshold": round(thresh - tot, 4),
                        "is_low_stock": True,
                        "freshness": freshness,
                        "batches": [dict(b) for b in b_rows]
                    })
            return alerts
        finally:
            if close_conn:
                conn.close()

    # -------------------------------------------------------------
    # OPT-IN LOW STOCK MULTI-BATCH (2–3 BATCHES) PRIORITY INSPECTION
    # -------------------------------------------------------------
    def inspect_low_stock_priority(
        self,
        material_code: str,
        plant: str = "3001",
        required_qty: float = 0.0,
        scope_sloc: Optional[str] = None,
        shop: Optional[str] = None,
        max_batches: int = 3
    ) -> Dict[str, Any]:
        """Opt-in multi-batch priority inspection per USER_SCOPE_LOW_STOCK_20261013.

        Invariants:
        1. Only executed upon explicit user selection (opt-in).
        2. Detailed inspection per batch/SLoc (no blind aggregation).
        3. Proposes 2–3 available batches to fulfill need, without draining entire batch.
        4. Excludes or clearly labels locked/unverified batches.
        5. Does not exceed required qty or available qty.
        6. Accurately reports other shop impact: 'Chưa đánh giá được nhu cầu/dự trữ shop khác'.
        7. Results labeled 'Ghi nhận / Đề xuất' (PROPOSAL), never claimed as live SAP write.
        """
        mat = str(material_code).strip()
        plt = str(plant).strip()
        if not mat:
            raise ValueError("Mã hàng không được rỗng")
        if not plt:
            raise ValueError("Plant không được rỗng")
        if isinstance(required_qty, bool) or not isinstance(required_qty, (int, float)):
            raise ValueError("Số lượng yêu cầu phải là số")
        req_qty = float(required_qty)
        if not math.isfinite(req_qty) or req_qty <= 0:
            raise ValueError("Số lượng yêu cầu phải hữu hạn và > 0")

        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            freshness = self.check_stock_freshness(plant=plt, sloc=scope_sloc)
            threshold = self.get_threshold(mat, plt)

            sql = f"""
                SELECT * FROM stock_items
                WHERE material_code = ? AND plant = ?
            """
            params: List[Any] = [mat, plt]
            if scope_sloc and scope_sloc.strip():
                sql += " AND sloc = ?"
                params.append(scope_sloc.strip())
            sql += f" AND {ACTIVE_STOCK_FILTER} ORDER BY priority_rank ASC, batch ASC, id ASC"

            rows = conn.execute(sql, params).fetchall()

            available_batches: List[Dict[str, Any]] = []
            locked_batches: List[Dict[str, Any]] = []

            for r in rows:
                d = dict(r)
                stype = str(d.get("stock_type") or "UNRESTRICTED").strip().upper()
                pflag = str(d.get("priority_flag") or "").strip().upper()
                qty = float(d.get("unrestricted_qty") or 0.0)

                is_locked = (stype != "UNRESTRICTED") or ("LOCK" in pflag) or ("UNVERIFIED" in pflag)
                if is_locked or qty <= 0:
                    locked_batches.append({
                        "batch": d.get("batch") or "",
                        "sloc": d.get("sloc") or "",
                        "qty": qty,
                        "stock_type": stype,
                        "is_locked": True,
                        "reason": f"Batch {d.get('batch')} đang khóa hoặc chưa xác minh (stock_type: {stype})" if is_locked else "Tồn = 0"
                    })
                else:
                    available_batches.append(d)

            current_stock = sum(float(b["unrestricted_qty"]) for b in available_batches)
            is_low_stock = current_stock < threshold

            # Multi-batch proposal (select 2–3 candidate batches, up to max_batches)
            cand_limit = max(1, min(max_batches, 3))
            cands = available_batches[:cand_limit]

            remaining = req_qty
            proposed_batches: List[Dict[str, Any]] = []

            for c in cands:
                b_avail = float(c["unrestricted_qty"] or 0.0)
                alloc = min(b_avail, remaining)
                if alloc > 0:
                    proposed_batches.append({
                        "batch": c.get("batch") or "",
                        "sloc": c.get("sloc") or "",
                        "available_qty": b_avail,
                        "proposed_qty": round(alloc, 4),
                        "stock_type": c.get("stock_type") or "UNRESTRICTED",
                        "priority_rank": c.get("priority_rank") or 0,
                        "unit": c.get("unit") or "PCS",
                        "is_locked": False
                    })
                    remaining = max(0.0, remaining - alloc)

            total_proposed = sum(b["proposed_qty"] for b in proposed_batches)
            shortage_qty = round(max(0.0, req_qty - total_proposed), 4)

            scope_desc = f"Plant {plt}"
            if scope_sloc:
                scope_desc += f", SLoc {scope_sloc}"
            if shop:
                scope_desc += f", Shop {shop}"

            return {
                "mode": "LOW_STOCK_PRIORITY_OPT_IN",
                "material_code": mat,
                "plant": plt,
                "shop": shop or ("BODY" if "3001" in plt else "EDC"),
                "scope_sloc": scope_sloc,
                "scope_description": scope_desc,
                "current_stock": round(current_stock, 4),
                "threshold_qty": threshold,
                "is_low_stock": is_low_stock,
                "required_qty": req_qty,
                "total_proposed_qty": round(total_proposed, 4),
                "shortage_qty": shortage_qty,
                "candidate_batches_count": len(proposed_batches),
                "proposed_batches": proposed_batches,
                "locked_batches": locked_batches,
                "freshness": freshness,
                "other_shop_impact": "Chưa đánh giá được nhu cầu/dự trữ shop khác",
                "prevent_line_stop_guaranteed": False,
                "result_type": "PROPOSAL",
                "status_label": "Ghi nhận / Đề xuất",
                "is_sap_live": False,
                "notes": "Chế độ tùy chọn tồn thấp: Đề xuất phân bổ 2–3 batch. Người dùng có quyền điều chỉnh số lượng mỗi batch trước khi thực hiện."
            }
        finally:
            if close_conn:
                conn.close()

    # -------------------------------------------------------------
    # APPLY LOW STOCK MULTI-BATCH ALLOCATION (GHI NHẬN / ĐỀ XUẤT)
    # -------------------------------------------------------------
    def apply_low_stock_allocation(
        self,
        material_code: str,
        plant: str,
        required_qty: float,
        allocations: List[Dict[str, Any]],
        shop: Optional[str] = None,
        scope_sloc: Optional[str] = None,
        target: str = "PROPOSAL",
        ob_number: Optional[str] = None,
        branch_tab: Optional[str] = None,
        evidence_note: Optional[str] = None,
        other_shop_reserves: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """Apply / record multi-batch low stock allocation per USER_SCOPE_LOW_STOCK_20261013."""
        mat = str(material_code).strip()
        plt = str(plant).strip()
        if not mat or not plt:
            raise ValueError("Cần có mã hàng và plant")
        if not allocations:
            raise ValueError("Danh sách phân bổ batch không được rỗng")
        if len(allocations) > 3:
            raise ValueError("Chế độ tồn thấp giới hạn tối đa 3 batch (đề xuất 2–3 batch)")

        req_qty = float(required_qty)
        if not math.isfinite(req_qty) or req_qty <= 0:
            raise ValueError("Số lượng yêu cầu không hợp lệ")

        conn = self._get_conn()
        close_conn = self._conn is None
        today_str = datetime.date.today().strftime("%Y%m%d")
        alloc_code = f"LSA-{today_str}-{uuid.uuid4().hex[:6].upper()}"

        try:
            if close_conn:
                conn.execute("BEGIN IMMEDIATE")
            with conn:
                cleaned_allocations = []
                total_allocated = 0.0
                seen_batches = set()

                for idx, a in enumerate(allocations, 1):
                    batch = str(a.get("batch") or "").strip()
                    sloc = str(a.get("sloc") or "").strip()
                    if not batch:
                        raise ValueError(f"Dòng {idx}: Thiếu số batch")
                    if batch in seen_batches:
                        raise ValueError(f"Batch {batch} bị trùng trong cùng lượt phân bổ")
                    seen_batches.add(batch)

                    raw_alloc = a.get("allocated_qty") if "allocated_qty" in a else a.get("qty", 0.0)
                    if isinstance(raw_alloc, bool):
                        raise ValueError(f"Dòng {idx}: Số lượng phân bổ không được là boolean")
                    f_alloc = float(raw_alloc)
                    if not math.isfinite(f_alloc) or f_alloc <= 0:
                        raise ValueError(f"Dòng {idx}: Số lượng phân bổ của batch {batch} phải > 0")

                    st_row = conn.execute(f"""
                        SELECT SUM(unrestricted_qty) as avail
                        FROM stock_items
                        WHERE material_code = ? AND plant = ? AND batch = ?
                          AND stock_type = 'UNRESTRICTED' AND {ACTIVE_STOCK_FILTER}
                    """, (mat, plt, batch)).fetchone()
                    avail = float(st_row["avail"] or 0.0) if st_row else 0.0

                    if f_alloc > avail:
                        raise ValueError(f"Batch {batch} chỉ có tồn khả dụng {avail:g}, không thể phân bổ {f_alloc:g}")

                    total_allocated += f_alloc
                    cleaned_allocations.append({
                        "batch": batch,
                        "sloc": sloc,
                        "allocated_qty": f_alloc,
                        "available_qty": avail
                    })

                if total_allocated > req_qty:
                    raise ValueError(f"Tổng lượng phân bổ ({total_allocated:g}) vượt quá nhu cầu yêu cầu ({req_qty:g})")

                if other_shop_reserves and isinstance(other_shop_reserves, dict) and other_shop_reserves.get("has_verified_data"):
                    impact_text = f"Đã đối soát với dự trữ shop khác: {other_shop_reserves.get('note', 'Hợp lệ')}"
                else:
                    impact_text = "Chưa đánh giá được nhu cầu/dự trữ shop khác"

                conn.execute("""
                    INSERT INTO low_stock_allocations (
                        allocation_code, material_code, plant, shop, scope_sloc,
                        required_qty, total_allocated_qty, status, result_type,
                        sap_live_status, other_shop_impact, allocations_json, evidence_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, 'Ghi nhận / Đề xuất', 'PROPOSAL', 'NOT_APPLIED_SAP', ?, ?, ?)
                """, (
                    alloc_code, mat, plt, shop or ("BODY" if "3001" in plt else "EDC"), scope_sloc,
                    req_qty, total_allocated, impact_text,
                    json.dumps(cleaned_allocations, ensure_ascii=False),
                    json.dumps({"evidence_note": evidence_note or "", "prevent_line_stop_guaranteed": False}, ensure_ascii=False)
                ))

                created_entries = []
                if target == "WAITING_CF":
                    now_iso = datetime.datetime.now().isoformat()
                    for c in cleaned_allocations:
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
                                '', '', ?, '',
                                'PENDING_CF', 0, 0, ?,
                                'LOW_STOCK_PROPOSAL', ?, ?, ?
                            )
                        """, (
                            alloc_code,
                            ob_number or "", mat, plt, c["batch"], c["allocated_qty"],
                            shop or ("BODY" if "3001" in plt else "EDC"),
                            c["allocated_qty"],
                            f"{ob_number or ''}|{mat}|{plt}|{c['batch']}||",
                            now_iso, now_iso
                        ))
                        created_entries.append({"line_id": cur.lastrowid, **c})

                elif target == "OB_ITEMS":
                    branch = branch_tab or ("BODY_3001" if "3001" in plt else "EDC_4007")
                    for idx, c in enumerate(cleaned_allocations, 1):
                        uid = f"{branch}:{alloc_code}:{idx}"
                        cur = conn.execute("""
                            INSERT INTO ob_items (
                                item_uid, branch_tab, branch, row_order,
                                material_code, model, req_qty, required_qty,
                                batch, shop, ob_number, status, rc52_code, note, selected
                            ) VALUES (?, ?, ?, ?, ?, '', ?, ?, ?, ?, ?, 'Ghi nhận / Đề xuất', 'RC52_LOW_STOCK', ?, 1)
                        """, (
                            uid, branch, branch, idx, mat, c["allocated_qty"], c["allocated_qty"],
                            c["batch"], shop or ("BODY" if "3001" in plt else "EDC"),
                            ob_number or "", f"Phân bổ tồn thấp: {alloc_code}"
                        ))
                        created_entries.append({"ob_item_id": cur.lastrowid, **c})

                return {
                    "status": "success",
                    "allocation_code": alloc_code,
                    "material_code": mat,
                    "plant": plt,
                    "required_qty": req_qty,
                    "total_allocated_qty": total_allocated,
                    "allocations": cleaned_allocations,
                    "other_shop_impact": impact_text,
                    "prevent_line_stop_guaranteed": False,
                    "result_type": "PROPOSAL",
                    "status_label": "Ghi nhận / Đề xuất",
                    "is_sap_live": False,
                    "target": target,
                    "created_entries": created_entries,
                    "message": "Đã ghi nhận đề xuất phân bổ tồn thấp thành công (Chưa tạo OB thật trên SAP)."
                }
        finally:
            if close_conn:
                conn.close()

    def get_zone_summary(self, plant: Optional[str] = None) -> List[Dict[str, Any]]:
        """Get aggregated inventory by storage location without double-counting."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            query = """
                SELECT plant, sloc, COUNT(DISTINCT material_code) as mat_count,
                       SUM(unrestricted_qty) as total_qty,
                       SUM(CASE WHEN pull_line_warning IS NOT NULL THEN 1 ELSE 0 END) as warning_count
                FROM stock_items
                WHERE 1=1
            """
            params = []
            if plant:
                query += " AND plant = ?"
                params.append(plant)
            query += " AND " + ACTIVE_STOCK_FILTER + " GROUP BY plant, sloc ORDER BY plant, sloc"

            rows = conn.execute(query, params).fetchall()
            results = []
            for r in rows:
                results.append({
                    "plant": r["plant"],
                    "sloc": r["sloc"],
                    "zone_name": ZONE_MAPPINGS.get(r["sloc"], f"SLoc {r['sloc']}"),
                    "material_count": r["mat_count"],
                    "total_qty": r["total_qty"] or 0.0,
                    "warning_count": r["warning_count"]
                })
            return results
        finally:
            if close_conn:
                conn.close()

    def query_stock(
        self,
        material_code: Optional[str] = None,
        plant: Optional[str] = None,
        sloc: Optional[str] = None,
        pull_line_only: bool = False
    ) -> List[Dict[str, Any]]:
        """Query stock items with filtering and leading-zero preservation."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            sql = "SELECT * FROM stock_items WHERE 1=1"
            params = []
            if material_code:
                sql += " AND material_code = ?"
                params.append(str(material_code).strip())
            if plant:
                sql += " AND plant = ?"
                params.append(plant)
            if sloc:
                sql += " AND sloc = ?"
                params.append(sloc)
            if pull_line_only:
                sql += " AND pull_line_warning IS NOT NULL"
            sql += " AND " + ACTIVE_STOCK_FILTER + " ORDER BY priority_rank ASC, id ASC"

            rows = conn.execute(sql, params).fetchall()
            mappings = {r["material_code"]: r["model_code"] for r in conn.execute("SELECT material_code,model_code FROM model_mappings").fetchall()}
            prefixes = sorted(mappings, key=len, reverse=True)
            result = []
            for row in rows:
                item = dict(row)
                item["model"] = mappings.get(item["material_code"], next((mappings[p] for p in prefixes if item["material_code"].startswith(p)), ""))
                result.append(item)
            return result
        finally:
            if close_conn:
                conn.close()

    # -------------------------------------------------------------
    # SAP MB52 AUTOMATED DIRECT DOWNLOAD & SNAPSHOT SYNC
    # -------------------------------------------------------------
    def download_and_sync_sap_mb52(
        self,
        plants: Optional[List[str]] = None,
        sloc: Optional[str] = None,
        layout: str = "/KHOMA",
        test_bytes: Optional[bytes] = None,
        test_filename: Optional[str] = None,
        is_user_upload: bool = False
    ) -> Dict[str, Any]:
        """Tải tồn trực tiếp từ SAP MB52 và cập nhật DB Snapshot.
        
        Quy tắc:
        1. MB52 tải trực tiếp từ SAP và cập nhật DB sau mỗi lượt thành công.
           Snapshot mới thay thế tồn hiện hành theo scope, không cộng dồn lịch sử.
        2. File Excel trung gian do SAP xuất ra: tự xóa sau khi parse và kiểm tra hợp lệ.
        3. Nếu lỗi: giữ lại file để chẩn đoán, không xóa file người dùng.
        4. Nếu chưa có phiên SAP live và không có test input, ghi rõ trạng thái chưa kết nối.
        """
        temp_dir = Path(tempfile.gettempdir())
        temp_file: Optional[Path] = None

        if test_bytes is not None:
            fname = test_filename or f"SAP_MB52_{uuid.uuid4().hex[:8]}.xlsx"
            temp_file = temp_dir / fname
            temp_file.write_bytes(test_bytes)
        else:
            conn_info = probe_connection(timeout=3)
            if conn_info.get("state") != "CONNECTED":
                return {
                    "status": "SAP_NOT_CONNECTED",
                    "connected": False,
                    "message": "Chưa có phiên SAP GUI kết nối trên máy tính này; không thể tải tồn tự động. Hãy đăng nhập SAP GUI trước khi tải tồn.",
                    "sap_state": conn_info.get("state", "DISCONNECTED"),
                    "details": conn_info
                }
            try:
                from app.services.sap_automation import download_sap_mb52_live
                temp_file = download_sap_mb52_live(plants=plants, sloc=sloc, layout=layout, timeout=60)
            except Exception as exc:
                return {
                    "status": "SAP_EXECUTION_ERROR",
                    "connected": True,
                    "message": f"Lỗi khi thực thi lệnh tải tồn MB52 trên SAP GUI: {str(exc)}",
                    "error_detail": str(exc),
                    "sap_state": "CONNECTED"
                }

        try:
            content = temp_file.read_bytes()
            rows = parse_source("mb52", content, temp_file.name)
            if not rows:
                raise ValueError("Nguồn MB52 không có dòng dữ liệu chi tiết")

            p_list = plants or ["3001", "3093"]
            first_plant = rows[0].get("plant") or p_list[0]
            first_sloc = rows[0].get("sloc") or (sloc or "1003")

            snap_res = self.create_snapshot(
                source_type="SAP_MB52_DIRECT",
                plant=first_plant,
                sloc=first_sloc,
                items=rows,
                notes=f"Tải tồn trực tiếp từ SAP (MB52 {layout}) - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            )

            # Tự xóa file trung gian sau khi parse và kiểm tra hợp lệ
            if not is_user_upload and temp_file.exists():
                try:
                    temp_file.unlink()
                except OSError:
                    pass

            return {
                "status": "SUCCESS",
                "message": f"Tải tồn MB52 từ SAP thành công: Đã nạp {len(rows)} dòng vào Snapshot #{snap_res['snapshot_id']}. File trung gian đã tự động xóa.",
                "snapshot": snap_res,
                "intermediate_file_cleaned": True,
                "diagnostic_file_retained": False,
                "total_rows": len(rows),
                "total_qty": snap_res.get("total_qty", 0.0)
            }
        except Exception as exc:
            retained_path = str(temp_file) if (temp_file and temp_file.exists() and not is_user_upload) else None
            return {
                "status": "ERROR",
                "message": f"Lỗi parse/kiểm tra MB52: {str(exc)}. Đã giữ lại file trung gian để chẩn đoán tại: {retained_path or 'N/A'}",
                "diagnostic_file": retained_path,
                "diagnostic_file_retained": bool(retained_path),
                "intermediate_file_cleaned": False,
                "error_detail": str(exc)
            }

    # -------------------------------------------------------------
    # VISUAL STOCK MATRIX (LINE + KHO) BODY REPORT & DRILLDOWN
    # -------------------------------------------------------------
    def get_visual_stock_matrix(self) -> Dict[str, Any]:
        """Báo cáo tổng hợp trực quan (LINE + KHO) BODY theo mẫu ảnh tham chiếu."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            latest_snap = conn.execute(
                "SELECT * FROM stock_snapshots ORDER BY id DESC LIMIT 1"
            ).fetchone()
            if not latest_snap:
                return {
                    "title": "(LINE + KHO) BODY",
                    "has_data": False,
                    "updated_time": "--",
                    "updated_date": "--",
                    "snapshot_id": None,
                    "columns": BUCKET_DEFS,
                    "rows": [
                        {"stt": idx, "model": mod, "counts": {b["key"]: 0 for b in BUCKET_DEFS}, "total_materials": 0}
                        for idx, mod in enumerate(CANONICAL_MODEL_ROWS, 1)
                    ],
                    "totals": {
                        "counts": {b["key"]: 0 for b in BUCKET_DEFS},
                        "percentages": {b["key"]: "0%" for b in BUCKET_DEFS},
                        "grand_total": 0
                    },
                    "message": "Chưa có dữ liệu snapshot tồn kho"
                }

            s_time = latest_snap["snapshot_time"] or datetime.datetime.now().isoformat()
            try:
                dt = datetime.datetime.fromisoformat(s_time)
                up_time = dt.strftime("%Hh%M")
                up_date = dt.strftime("%d-%b")
            except Exception:
                up_time = "5h30"
                up_date = "13-Oct"

            # Query materials in Body scope (plants 3001, 3093) with active stock filter
            sql = f"""
                SELECT stock_items.material_code, stock_items.sloc, stock_items.plant, SUM(stock_items.unrestricted_qty) as qty,
                       COALESCE(m.us_rate, 1) as us_rate,
                       COALESCE(m.model, '') as mat_model
                FROM stock_items
                LEFT JOIN materials m ON stock_items.material_code = m.material_code
                WHERE (stock_items.plant = '3001' OR stock_items.plant = '3093')
                  AND {ACTIVE_STOCK_FILTER}
                GROUP BY stock_items.material_code, stock_items.sloc, stock_items.plant
            """
            rows = conn.execute(sql).fetchall()

            mappings = {r["material_code"]: r["model_code"] for r in conn.execute("SELECT material_code,model_code FROM model_mappings").fetchall()}

            by_mat: Dict[str, Dict[str, Any]] = {}
            for r in rows:
                mat = str(r["material_code"]).strip()
                sloc = str(r["sloc"]).strip()
                qty = float(r["qty"] or 0.0)
                if mat not in by_mat:
                    mod = r["mat_model"] or mappings.get(mat, "")
                    by_mat[mat] = {
                        "material_code": mat,
                        "model": mod,
                        "us_rate": int(r["us_rate"] or 1),
                        "sloc_3003": 0.0,
                        "sloc_1003": 0.0,
                        "sloc_3000": 0.0,
                        "sloc_3093": 0.0,
                        "sloc_3001": 0.0
                    }
                if sloc == "3003":
                    by_mat[mat]["sloc_3003"] += qty
                elif sloc == "1003":
                    by_mat[mat]["sloc_1003"] += qty
                elif sloc == "3000":
                    by_mat[mat]["sloc_3000"] += qty
                elif sloc == "3093":
                    by_mat[mat]["sloc_3093"] += qty
                elif sloc == "3001":
                    by_mat[mat]["sloc_3001"] += qty

            row_data: Dict[str, Dict[str, int]] = {
                mod: {b["key"]: 0 for b in BUCKET_DEFS} for mod in CANONICAL_MODEL_ROWS
            }
            row_totals: Dict[str, int] = {mod: 0 for mod in CANONICAL_MODEL_ROWS}

            for mat, d in by_mat.items():
                c_mod = map_model_to_canonical(d["model"])
                if not c_mod:
                    for cand in CANONICAL_MODEL_ROWS:
                        if cand.upper() in d["model"].upper():
                            c_mod = cand
                            break
                if not c_mod:
                    c_mod = "VF8"

                tong_body = d["sloc_3003"] + d["sloc_1003"]
                us = d["us_rate"] if d["us_rate"] > 0 else 1
                doh = int(tong_body / us)
                bucket = get_doh_bucket(doh)

                row_data[c_mod][bucket] += 1
                row_totals[c_mod] += 1

            table_rows = []
            col_totals: Dict[str, int] = {b["key"]: 0 for b in BUCKET_DEFS}
            grand_total = sum(row_totals.values())

            for idx, mod in enumerate(CANONICAL_MODEL_ROWS, 1):
                counts = row_data[mod]
                for b in BUCKET_DEFS:
                    col_totals[b["key"]] += counts[b["key"]]
                table_rows.append({
                    "stt": idx,
                    "model": mod,
                    "counts": counts,
                    "total_materials": row_totals[mod]
                })

            percentages: Dict[str, str] = {}
            for b in BUCKET_DEFS:
                c_cnt = col_totals[b["key"]]
                if grand_total > 0:
                    pct = round(c_cnt / grand_total * 100, 1)
                    percentages[b["key"]] = f"{pct:g}%"
                else:
                    percentages[b["key"]] = "0%"

            return {
                "title": "(LINE + KHO) BODY",
                "has_data": True,
                "updated_time": up_time,
                "updated_date": up_date,
                "snapshot_id": latest_snap["id"],
                "snapshot_time": s_time,
                "columns": BUCKET_DEFS,
                "rows": table_rows,
                "totals": {
                    "counts": col_totals,
                    "percentages": percentages,
                    "grand_total": grand_total
                }
            }
        finally:
            if close_conn:
                conn.close()

    def get_visual_matrix_materials(self, model_key: Optional[str] = None, bucket_key: Optional[str] = None) -> List[Dict[str, Any]]:
        """Lấy danh sách chi tiết các mã khi click vào ô số tổng trên ma trận (LINE + KHO) BODY."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            sql = f"""
                SELECT stock_items.material_code, stock_items.sloc, stock_items.plant, SUM(stock_items.unrestricted_qty) as qty,
                       COALESCE(m.us_rate, 1) as us_rate,
                       COALESCE(m.model, '') as mat_model
                FROM stock_items
                LEFT JOIN materials m ON stock_items.material_code = m.material_code
                WHERE (stock_items.plant = '3001' OR stock_items.plant = '3093')
                  AND {ACTIVE_STOCK_FILTER}
                GROUP BY stock_items.material_code, stock_items.sloc, stock_items.plant
            """
            rows = conn.execute(sql).fetchall()
            mappings = {r["material_code"]: r["model_code"] for r in conn.execute("SELECT material_code,model_code FROM model_mappings").fetchall()}

            by_mat: Dict[str, Dict[str, Any]] = {}
            for r in rows:
                mat = str(r["material_code"]).strip()
                sloc = str(r["sloc"]).strip()
                qty = float(r["qty"] or 0.0)
                if mat not in by_mat:
                    mod = r["mat_model"] or mappings.get(mat, "")
                    by_mat[mat] = {
                        "material_code": mat,
                        "model": mod,
                        "us_rate": int(r["us_rate"] or 1),
                        "line_3003": 0.0,
                        "kho_1003": 0.0,
                        "line_3000": 0.0,
                        "kho_3093": 0.0,
                        "in_transit_3001": 0.0
                    }
                if sloc == "3003":
                    by_mat[mat]["line_3003"] += qty
                elif sloc == "1003":
                    by_mat[mat]["kho_1003"] += qty
                elif sloc == "3000":
                    by_mat[mat]["line_3000"] += qty
                elif sloc == "3093":
                    by_mat[mat]["kho_3093"] += qty
                elif sloc == "3001":
                    by_mat[mat]["in_transit_3001"] += qty

            results = []
            for mat, d in by_mat.items():
                c_mod = map_model_to_canonical(d["model"]) or d["model"]
                if model_key and model_key.upper() not in c_mod.upper() and c_mod.upper() not in model_key.upper():
                    continue
                tong_body = d["line_3003"] + d["kho_1003"]
                us = d["us_rate"] if d["us_rate"] > 0 else 1
                doh = int(tong_body / us)
                b = get_doh_bucket(doh)
                if bucket_key and b != bucket_key:
                    continue

                d["tong_body"] = tong_body
                d["doh"] = doh
                d["bucket"] = b
                d["canonical_model"] = c_mod
                results.append(d)

            return sorted(results, key=lambda x: (x["doh"], x["material_code"]))
        finally:
            if close_conn:
                conn.close()

    # -------------------------------------------------------------
    # KHU GỢI Ý TỒN THẤP & CHUYỂN GỢI Ý SANG OB
    # -------------------------------------------------------------
    def get_low_stock_suggestions(
        self,
        plant: str = "3001",
        scope_sloc: Optional[str] = None,
        shop: str = "BODY"
    ) -> List[Dict[str, Any]]:
        """Khu gợi ý tồn thấp:
        
        Tính lượng bổ sung từ tồn hiện hành và DOH của model:
        DOH = INT(Tổng Body / US) theo baseline 261008_Stock_FG_MA_BODY_V3_6.xlsm.
        Không cố định 2-3 batch hay lấy hết batch. Tôn trọng plant/kho/shop và lượng khả dụng.
        """
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            sql = f"""
                SELECT stock_items.material_code, stock_items.sloc, stock_items.plant, SUM(stock_items.unrestricted_qty) as qty,
                       COALESCE(m.us_rate, 1) as us_rate,
                       COALESCE(m.model, '') as mat_model
                FROM stock_items
                LEFT JOIN materials m ON stock_items.material_code = m.material_code
                WHERE stock_items.plant = ? AND {ACTIVE_STOCK_FILTER}
            """
            params: List[Any] = [plant]
            if scope_sloc and scope_sloc.strip():
                sql += " AND stock_items.sloc = ?"
                params.append(scope_sloc.strip())
            sql += " GROUP BY stock_items.material_code, stock_items.sloc, stock_items.plant"

            rows = conn.execute(sql, params).fetchall()
            mappings = {r["material_code"]: r["model_code"] for r in conn.execute("SELECT material_code,model_code FROM model_mappings").fetchall()}

            by_mat: Dict[str, Dict[str, Any]] = {}
            for r in rows:
                mat = str(r["material_code"]).strip()
                sloc = str(r["sloc"]).strip()
                qty = float(r["qty"] or 0.0)
                if mat not in by_mat:
                    mod = r["mat_model"] or mappings.get(mat, "")
                    by_mat[mat] = {
                        "material_code": mat,
                        "model": mod,
                        "us_rate": int(r["us_rate"] or 1),
                        "line_3003": 0.0,
                        "kho_1003": 0.0,
                        "other_qty": 0.0
                    }
                if sloc == "3003":
                    by_mat[mat]["line_3003"] += qty
                elif sloc == "1003":
                    by_mat[mat]["kho_1003"] += qty
                else:
                    by_mat[mat]["other_qty"] += qty

            suggestions = []
            default_doh_threshold = 200

            for mat, d in by_mat.items():
                tong_body = d["line_3003"] + d["kho_1003"]
                us = d["us_rate"] if d["us_rate"] > 0 else 1
                doh = int(tong_body / us)

                t_row = conn.execute(
                    "SELECT threshold_qty FROM stock_thresholds WHERE material_code = ? AND plant = ?",
                    (mat, plant)
                ).fetchone()
                thresh_doh = float(t_row["threshold_qty"]) if t_row else default_doh_threshold

                if doh < thresh_doh:
                    target_qty = int(thresh_doh * us)
                    needed = max(0.0, float(target_qty - tong_body))

                    b_rows = conn.execute(f"""
                        SELECT batch, sloc, unrestricted_qty, priority_rank, stock_type
                        FROM stock_items
                        WHERE material_code = ? AND plant = ? AND stock_type = 'UNRESTRICTED'
                          AND {ACTIVE_STOCK_FILTER}
                        ORDER BY priority_rank ASC, batch ASC
                    """, (mat, plant)).fetchall()

                    batches = []
                    rem = needed
                    for b in b_rows:
                        b_qty = float(b["unrestricted_qty"] or 0.0)
                        alloc = min(b_qty, rem)
                        batches.append({
                            "batch": b["batch"],
                            "sloc": b["sloc"],
                            "available_qty": b_qty,
                            "proposed_qty": alloc,
                            "allocated_qty": alloc,
                            "priority_rank": b["priority_rank"]
                        })
                        rem = max(0.0, rem - alloc)

                    suggestions.append({
                        "material_code": mat,
                        "model": d["model"],
                        "us_rate": us,
                        "line_3003": d["line_3003"],
                        "kho_1003": d["kho_1003"],
                        "tong_body": tong_body,
                        "current_doh": doh,
                        "threshold_doh": thresh_doh,
                        "target_qty": target_qty,
                        "supplement_needed": needed,
                        "required_qty": needed,
                        "shop": shop,
                        "plant": plant,
                        "batches": batches
                    })

            return sorted(suggestions, key=lambda x: (x["current_doh"], x["material_code"]))
        finally:
            if close_conn:
                conn.close()

    def transfer_suggestions_to_ob(
        self,
        items: List[Dict[str, Any]],
        branch_tab: str = "BODY_3001"
    ) -> Dict[str, Any]:
        """Chuyển danh sách gợi ý tồn thấp vào bảng điền OB (ob_items)."""
        if not items:
            raise ValueError("Danh sách gợi ý chuyển sang OB không được rỗng")
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            with conn:
                max_order_row = conn.execute(
                    "SELECT COALESCE(MAX(row_order), 0) FROM ob_items WHERE branch_tab = ?",
                    (branch_tab,)
                ).fetchone()[0]
                cur_order = max_order_row + 1

                created = []
                for it in items:
                    mat = str(it.get("material_code", "")).strip()
                    if not mat:
                        continue
                    mod = str(it.get("model", "")).strip()
                    qty = float(it.get("required_qty") or it.get("qty") or 0.0)
                    batch = str(it.get("batch", "")).strip()
                    shop = str(it.get("shop", "BODY")).strip()
                    uid = f"{branch_tab}:SUGG:{uuid.uuid4().hex[:8]}"

                    cur = conn.execute("""
                        INSERT INTO ob_items (
                            item_uid, branch_tab, branch, row_order, selected,
                            material_code, model, req_qty, required_qty, batch, shop,
                            status, note
                        ) VALUES (?, ?, ?, ?, 1, ?, ?, ?, ?, ?, ?, 'Gợi ý tồn thấp', 'Chuyển từ khu gợi ý tồn thấp')
                    """, (
                        uid, branch_tab, branch_tab, cur_order,
                        mat, mod, qty, qty, batch, shop
                    ))
                    created.append({"id": cur.lastrowid, "item_uid": uid, "material_code": mat, "qty": qty})
                    cur_order += 1

            return {
                "status": "success",
                "transferred_count": len(created),
                "items": created,
                "branch_tab": branch_tab
            }
        finally:
            if close_conn:
                conn.close()

    def get_two_stock_tables(self, plant: str = "3001") -> Dict[str, Any]:
        """Lấy 2 bảng tồn trên màn tồn: Bảng ưu tiên và Bảng tổng tồn."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            sql = f"""
                SELECT stock_items.material_code, stock_items.sloc, stock_items.plant, stock_items.batch, stock_items.unrestricted_qty, stock_items.pull_line_warning, stock_items.priority_rank,
                       COALESCE(m.us_rate, 1) as us_rate,
                       COALESCE(m.model, '') as mat_model
                FROM stock_items
                LEFT JOIN materials m ON stock_items.material_code = m.material_code
                WHERE stock_items.plant = ? AND {ACTIVE_STOCK_FILTER}
                ORDER BY stock_items.material_code ASC, stock_items.sloc ASC
            """
            rows = conn.execute(sql, (plant,)).fetchall()
            mappings = {r["material_code"]: r["model_code"] for r in conn.execute("SELECT material_code,model_code FROM model_mappings").fetchall()}

            by_mat: Dict[str, Dict[str, Any]] = {}
            for r in rows:
                mat = str(r["material_code"]).strip()
                sloc = str(r["sloc"]).strip()
                qty = float(r["unrestricted_qty"] or 0.0)
                pull = r["pull_line_warning"]

                if mat not in by_mat:
                    mod = r["mat_model"] or mappings.get(mat, "")
                    by_mat[mat] = {
                        "material_code": mat,
                        "model": mod,
                        "us_rate": int(r["us_rate"] or 1),
                        "line_3003": 0.0,
                        "kho_1003": 0.0,
                        "line_3000": 0.0,
                        "kho_3093": 0.0,
                        "in_transit_3001": 0.0,
                        "pull_line_warning": pull,
                        "priority_rank": r["priority_rank"]
                    }
                if pull and not by_mat[mat]["pull_line_warning"]:
                    by_mat[mat]["pull_line_warning"] = pull

                if sloc == "3003":
                    by_mat[mat]["line_3003"] += qty
                elif sloc == "1003":
                    by_mat[mat]["kho_1003"] += qty
                elif sloc == "3000":
                    by_mat[mat]["line_3000"] += qty
                elif sloc == "3093":
                    by_mat[mat]["kho_3093"] += qty
                elif sloc == "3001":
                    by_mat[mat]["in_transit_3001"] += qty

            priority_table = []
            total_stock_table = []

            for mat, d in by_mat.items():
                tong_body = d["line_3003"] + d["kho_1003"]
                us = d["us_rate"] if d["us_rate"] > 0 else 1
                doh = int(tong_body / us)
                d["tong_body"] = tong_body
                d["doh"] = doh

                total_stock_table.append(d)

                is_priority = (doh <= 200) or (d["pull_line_warning"] is not None) or (d["priority_rank"] and d["priority_rank"] < 50)
                if is_priority:
                    warn_reason = []
                    if d["pull_line_warning"]:
                        warn_reason.append(d["pull_line_warning"])
                    if doh <= 200:
                        warn_reason.append(f"DOH thấp ({doh} XE <= 200 XE)")
                    p_entry = dict(d)
                    p_entry["reason"] = "; ".join(warn_reason) or "Ưu tiên xuất"
                    priority_table.append(p_entry)

            return {
                "priority_table": sorted(priority_table, key=lambda x: (x["doh"], x["material_code"])),
                "total_stock_table": sorted(total_stock_table, key=lambda x: (x["model"], x["material_code"])),
                "priority_count": len(priority_table),
                "total_count": len(total_stock_table)
            }
        finally:
            if close_conn:
                conn.close()


# Module-level helpers
CANONICAL_MODEL_ROWS = ["Limo", "VF5", "VF6", "VF6,7", "VF6 CIS", "VF7", "VF8", "VF8NR"]

def map_model_to_canonical(model_name: str) -> Optional[str]:
    m = str(model_name or "").strip().upper()
    if not m:
        return None
    if "CIS" in m:
        return "VF6 CIS"
    if "6,7" in m or "67" in m or "6/7" in m:
        return "VF6,7"
    if "VF8NP" in m or "VF8NR" in m:
        return "VF8NR"
    if "LIMO" in m:
        return "Limo"
    if "VF5" in m:
        return "VF5"
    if "VF6" in m:
        return "VF6"
    if "VF7" in m:
        return "VF7"
    if "VF8" in m:
        return "VF8"
    return None

BUCKET_DEFS = [
    {"key": "0XE", "label": "0XE", "min": 0, "max": 0, "desc": "0 XE", "is_alert": True},
    {"key": "<=50XE", "label": "<=50XE", "min": 1, "max": 50, "desc": "1–50 XE", "is_alert": True},
    {"key": "<=100XE", "label": "<=100XE", "min": 51, "max": 100, "desc": "51–100 XE", "is_alert": True},
    {"key": "<=150XE", "label": "<=150XE", "min": 101, "max": 150, "desc": "101–150 XE", "is_alert": True},
    {"key": "<=200XE", "label": "<=200XE", "min": 151, "max": 200, "desc": "151–200 XE", "is_alert": True},
    {"key": "<=300XE", "label": "<=300XE", "min": 201, "max": 300, "desc": "201–300 XE", "is_alert": False},
    {"key": "<=400XE", "label": "<=400XE", "min": 301, "max": 400, "desc": "301–400 XE", "is_alert": False},
    {"key": "<=500XE", "label": "<=500XE", "min": 401, "max": 500, "desc": "401–500 XE", "is_alert": False},
    {"key": ">500XE", "label": ">500XE", "min": 501, "max": None, "desc": ">500 XE", "is_alert": False},
]

def get_doh_bucket(doh: int) -> str:
    if doh <= 0:
        return "0XE"
    if doh <= 50:
        return "<=50XE"
    if doh <= 100:
        return "<=100XE"
    if doh <= 150:
        return "<=150XE"
    if doh <= 200:
        return "<=200XE"
    if doh <= 300:
        return "<=300XE"
    if doh <= 400:
        return "<=400XE"
    if doh <= 500:
        return "<=500XE"
    return ">500XE"
