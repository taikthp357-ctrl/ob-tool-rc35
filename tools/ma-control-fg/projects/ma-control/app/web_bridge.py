# -*- coding: utf-8 -*-
"""MA CONTROL FG — Web Bridge & In-Process Zero-Port Data Engine.

Inherited from MA-Logistic-Control architecture:
- 100% Zero-Port (Zero listening TCP/UDP sockets, corporate firewall compliant)
- In-process synchronization between SQLite (ma_control.db) and static data.js
- In-process action dispatcher for 5-step workflow (Check OB, Fill OB, Create Pick, Confirm CF, BBBG)
- Preserves all invariants: Zero Auto-CF, strict Stock Zone separation, RC52 & V2.2.14 parity.
"""

from __future__ import annotations

import argparse
import datetime
import json
import logging
import os
import sys
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

APP_ROOT = Path(__file__).resolve().parent.parent
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from app.core.config import APP_TITLE, APP_VERSION, BASE_DIR, DATA_DIR, DB_PATH
from app.core.database import get_connection, init_db
from app.services.stock_service import StockService
from app.services.ob_service import OBService
from app.services.pick_cf_service import PickCFService
from app.services.bbbg_service import BBBGService
from app.services.reconciliation_service import reconcile_direct_sap

logger = logging.getLogger("MAControlFG.WebBridge")

DEFAULT_DATA_JS_STATIC = BASE_DIR / "app" / "static" / "js" / "data.js"
DEFAULT_DATA_JS_TEMPLATES = BASE_DIR / "app" / "templates" / "data.js"
DEFAULT_DATA_JS_WEB = BASE_DIR / "app" / "web" / "data.js"
ACTION_QUEUE_FILE = DATA_DIR / "action_queue.json"


def get_system_snapshot(db_path: Optional[str] = None) -> Dict[str, Any]:
    """Extracts complete in-process system state from SQLite into a structured dictionary."""
    conn = get_connection(db_path)
    try:
        # 1. OB items Body 3001 & EDC 4007
        rows_body = conn.execute(
            "SELECT * FROM ob_items WHERE branch = 'BODY_3001' ORDER BY created_at ASC"
        ).fetchall()
        ob_body = [dict(r) for r in rows_body]

        rows_edc = conn.execute(
            "SELECT * FROM ob_items WHERE branch = 'EDC_4007' ORDER BY created_at ASC"
        ).fetchall()
        ob_edc = [dict(r) for r in rows_edc]

        # 2. Pick orders & items
        rows_pick = conn.execute(
            "SELECT * FROM pick_orders ORDER BY created_at DESC"
        ).fetchall()
        pick_orders = [dict(r) for r in rows_pick]

        # 3. BBBG records
        rows_bbbg = conn.execute(
            "SELECT * FROM bbbg_records ORDER BY created_at DESC"
        ).fetchall()
        bbbg_records = [dict(r) for r in rows_bbbg]

        # 4. Stock snapshots summary
        latest_snapshot = conn.execute(
            "SELECT * FROM stock_snapshots ORDER BY snapshot_time DESC LIMIT 1"
        ).fetchone()
        stock_summary = dict(latest_snapshot) if latest_snapshot else {
            "id": "NONE",
            "snapshot_time": datetime.datetime.now().isoformat(),
            "source": "INITIAL_BASELINE",
            "total_materials": 0,
            "total_qty": 0.0
        }

        # 5. Low stock alerts (< 200 cars / pull line)
        alerts = []
        rows_alerts = conn.execute("""
            SELECT material_code, plant, sloc, qty, priority_flag 
            FROM stock_items 
            WHERE priority_flag IN ('LOW_STOCK_200', 'PULL_LINE', 'PRIORITY_OUT')
            LIMIT 50
        """).fetchall()
        # 6. Recent system activities for Dashboard
        rows_events = conn.execute("""
            SELECT job_id, step_name, event_type, target_ref, status, created_at,
                   COALESCE(evidence, evidence_json, '') AS evidence
            FROM job_events ORDER BY created_at DESC LIMIT 20
        """).fetchall()
        recent_events = [dict(r) for r in rows_events]

        # 7. Load baseline dictionary if available
        dict_path = APP_ROOT / "app" / "core" / "baseline_dictionary.json"
        baseline_materials = {}
        baseline_obs = {}
        priority_warnings = []
        shifts = [
            {"id": "S1", "name": "Ca sáng", "time_range": "06:00 - 14:00", "hours": "8h"},
            {"id": "S2", "name": "Ca chiều", "time_range": "14:00 - 22:00", "hours": "8h"},
            {"id": "S3", "name": "Ca đêm", "time_range": "22:00 - 06:00", "hours": "8h"}
        ]
        if dict_path.exists():
            try:
                with open(dict_path, "r", encoding="utf-8") as f:
                    dict_data = json.load(f)
                    baseline_materials = dict_data.get("materials", {})
                    baseline_obs = dict_data.get("obs", {})
                    priority_warnings = dict_data.get("priority_warnings", [])
                    if dict_data.get("shifts"):
                        shifts = dict_data["shifts"]
            except Exception as e:
                logger.warning(f"Could not load baseline_dictionary.json: {e}")

        return {
            "version": APP_VERSION,
            "app_title": APP_TITLE,
            "zero_port_mode": True,
            "generated_at": datetime.datetime.now().isoformat(),
            "sap_status": "SAP_GUI_AVAILABLE",
            "stock_summary": stock_summary,
            "stock_alerts": alerts,
            "priority_warnings": priority_warnings,
            "shifts": shifts,
            "baseline_materials": baseline_materials,
            "baseline_obs": baseline_obs,
            "recent_events": recent_events,
            "ob_items_body_3001": ob_body,
            "ob_items_edc_4007": ob_edc,
            "pick_orders": pick_orders,
            "bbbg_records": bbbg_records,
            "waiting_cf_count": len([p for p in pick_orders if p.get("status") == "WAITING_USER_CF"]),
        }
    finally:
        conn.close()


def export_data_js(dataset: Dict[str, Any], output_path: Optional[Path] = None) -> List[Path]:
    """Serializes dataset into data.js for zero-port Edge App consumption."""
    out_paths = [output_path] if output_path else [DEFAULT_DATA_JS_STATIC, DEFAULT_DATA_JS_TEMPLATES, DEFAULT_DATA_JS_WEB]
    
    js_content = (
        f"// Auto-generated MA CONTROL FG Zero-Port Data Bridge\n"
        f"// Generated at: {datetime.datetime.now().isoformat()}\n"
        f"// 100% Zero-Port compliant (no open TCP sockets)\n"
        f"window.MA_CONTROL_DATA = {json.dumps(dataset, indent=2, ensure_ascii=False)};\n"
        f"window.OB_ITEMS_BODY = window.MA_CONTROL_DATA.ob_items_body_3001 || [];\n"
        f"window.OB_ITEMS_EDC = window.MA_CONTROL_DATA.ob_items_edc_4007 || [];\n"
        f"window.PICK_ORDERS = window.MA_CONTROL_DATA.pick_orders || [];\n"
        f"window.BBBG_RECORDS = window.MA_CONTROL_DATA.bbbg_records || [];\n"
        f"window.BASELINE_MATERIALS = window.MA_CONTROL_DATA.baseline_materials || {{}};\n"
        f"window.BASELINE_OBS = window.MA_CONTROL_DATA.baseline_obs || {{}};\n"
        f"window.PRIORITY_WARNINGS = window.MA_CONTROL_DATA.priority_warnings || [];\n"
        f"window.RECENT_EVENTS = window.MA_CONTROL_DATA.recent_events || [];\n"
        f"window.SHIFT_CONFIG = window.MA_CONTROL_DATA.shifts || [];\n"
    )

    written = []
    for p in out_paths:
        p.parent.mkdir(parents=True, exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            f.write(js_content)
        written.append(p)

    return written


def sync_web_dashboard(db_path: Optional[str] = None, output_path: Optional[Path] = None) -> Dict[str, Any]:
    """Runs a fast in-process sync from SQLite to data.js without opening any network ports."""
    init_db(db_path)
    dataset = get_system_snapshot(db_path)
    export_data_js(dataset, output_path)
    return dataset


class ZeroPortActionDispatcher:
    """Executes business workflow actions in-process without network latency or sockets."""

    @classmethod
    def check_ob_items(cls, branch: str = "BODY_3001", db_path: Optional[str] = None) -> Dict[str, Any]:
        """Step 2a: Check OB stock readiness."""
        conn = get_connection(db_path)
        try:
            service = OBService(conn=conn)
            updated = service.check_ob_items(branch_tab=branch)
        finally:
            conn.close()
        sync_web_dashboard(db_path)
        return {"status": "SUCCESS", "checked_count": len(updated), "items": updated}

    @classmethod
    def fill_ob_items(cls, branch: str = "BODY_3001", db_path: Optional[str] = None) -> Dict[str, Any]:
        """Step 2b: Fill OB according to RC52 logic."""
        conn = get_connection(db_path)
        try:
            service = OBService(conn=conn)
            filled = service.fill_ob_items(branch_tab=branch)
        finally:
            conn.close()
        sync_web_dashboard(db_path)
        return {"status": "SUCCESS", "filled_count": len(filled), "items": filled}

    @classmethod
    def create_pick_order(
        cls, shop: str, model: str, items: List[Dict[str, Any]], branch_tab: str = "BODY_3001", db_path: Optional[str] = None
    ) -> Dict[str, Any]:
        """Step 3: Generate Pick Order -> INVARIANT: MUST BE WAITING_USER_CF."""
        conn = get_connection(db_path)
        try:
            service = PickCFService(conn=conn)
            result = service.create_pick_order(shop=shop, model=model, items=items, branch_tab=branch_tab)
        finally:
            conn.close()
        sync_web_dashboard(db_path)
        return result

    @classmethod
    def confirm_pick_order(
        cls, pick_order_no: str, user_name: str, actual_quantities: Optional[Dict[str, float]] = None, db_path: Optional[str] = None
    ) -> Dict[str, Any]:
        """Step 4: Manual CF Gate Confirmation by user -> Unlocks Step 5."""
        conn = get_connection(db_path)
        try:
            if actual_quantities is None:
                if isinstance(pick_order_no, int) or (isinstance(pick_order_no, str) and str(pick_order_no).isdigit()):
                    po_row = conn.execute("SELECT id FROM pick_orders WHERE id = ?", (int(pick_order_no),)).fetchone()
                else:
                    po_row = conn.execute(
                        "SELECT id FROM pick_orders WHERE pick_order_code = ? OR pick_order_no = ?",
                        (str(pick_order_no), str(pick_order_no))
                    ).fetchone()
                if po_row:
                    po_id = po_row["id"]
                    lines = conn.execute(
                        "SELECT id, qty FROM pick_order_items WHERE pick_order_id = ? ORDER BY id",
                        (po_id,)
                    ).fetchall()
                    actual_quantities = {str(line["id"]): float(line["qty"]) for line in lines}

            service = PickCFService(conn=conn)
            result = service.verify_cf(
                pick_order_ref=pick_order_no,
                verified_by=user_name,
                actual_quantities=actual_quantities
            )
        finally:
            conn.close()
        sync_web_dashboard(db_path)
        return result

    @classmethod
    def generate_bbbg(
        cls,
        pick_order_no: str,
        sender_name: str,
        receiver_name: str = "",
        storage_bin: str = "",
        db_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Step 5: Handover record generation (Blocked if not USER_CONFIRMED)."""
        conn = get_connection(db_path)
        try:
            service = BBBGService(conn=conn)
            result = service.create_bbbg(
                pick_order_ref=pick_order_no,
                sender_name=sender_name,
                receiver_name=receiver_name,
                storage_bin=storage_bin,
            )
        finally:
            conn.close()
        sync_web_dashboard(db_path)
        return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="MA CONTROL FG - Zero-Port Web Bridge")
    parser.add_argument("--sync", action="store_true", help="Synchronize SQLite to data.js")
    parser.add_argument("--db", type=str, default=None, help="Custom database path")
    parser.add_argument("--action", type=str, default=None, help="In-process action to execute")
    parser.add_argument("--branch", type=str, default="BODY_3001", help="Branch: BODY_3001 or EDC_4007")
    parser.add_argument("--pick-no", type=str, default=None, help="Pick order number")
    parser.add_argument("--user", type=str, default="GL_BODY_01", help="User name for CF")
    args = parser.parse_args()

    if args.sync or not args.action:
        res = sync_web_dashboard(args.db)
        print(f"[OK] Synced {len(res['ob_items_body_3001'])} Body items, {len(res['ob_items_edc_4007'])} EDC items, {len(res['pick_orders'])} pick orders to data.js.")
        sys.exit(0)

    act = args.action.upper()
    if act == "CHECK_OB":
        out = ZeroPortActionDispatcher.check_ob_items(args.branch, args.db)
        print(json.dumps(out, indent=2, ensure_ascii=False))
    elif act == "FILL_OB":
        out = ZeroPortActionDispatcher.fill_ob_items(args.branch, args.db)
        print(json.dumps(out, indent=2, ensure_ascii=False))
    elif act == "CONFIRM_CF":
        if not args.pick_no:
            print("[ERROR] Missing --pick-no")
            sys.exit(1)
        out = ZeroPortActionDispatcher.confirm_pick_order(args.pick_no, args.user, args.db)
        print(json.dumps(out, indent=2, ensure_ascii=False))
    elif act == "GENERATE_BBBG":
        if not args.pick_no:
            print("[ERROR] Missing --pick-no")
            sys.exit(1)
        out = ZeroPortActionDispatcher.generate_bbbg(args.pick_no, args.user, db_path=args.db)
        print(json.dumps(out, indent=2, ensure_ascii=False))
    else:
        print(f"[ERROR] Unknown action: {act}")
        sys.exit(1)
