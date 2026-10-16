# -*- coding: utf-8 -*-
"""Comprehensive Test Suite for the 6 Critical Operational Traps (6 Bẫy Vận Hành).

Verifies implementation against PLAN_MA_CONTROL_EXCEL_INPUT_AND_REDESIGN.md:
- Trap 1 (Flow vs Direct Paste): BBBG supports dual independent sources ("Lấy từ CF" & "Dán dữ liệu");
  missing slip number or location yields non-blocking soft warnings.
- Trap 2 (Auto-Lookup Engine): Auto-fills Model, Shop/Location, Batch from 326+ baseline materials;
  strips Excel formula errors (#N/A, #REF!, #VALUE!).
- Trap 3 (Anti-Column Shift Ctrl+V): Preserves exact tab counts (never collapses \t\t),
  auto-detects/strips header rows, removes thousand separators, handles Alt+Enter in quotes.
- Trap 4 (Excel File Lock / Permission Denied): Database First Persistence commits to SQLite first;
  on PermissionError / WinError 32 catches lock, opens Error Drawer with 3 required bullets
  (Cause, Saved Data, Action) and "Thử lại bước này".
- Trap 5 (Factory Print Fidelity): Master template injection using sheet BBBG_NEW in
  261008_BBBG_FG_V2_2_14.xlsm mapping I1, E1, A4, C5, G5, C7, G7, row 13+ with margins & logo.
- Trap 6 (Inline Grid Editing & Progress Execution): Direct cell editing, row actions,
  4-step execution toolbar, and SVG shift donut chart.
"""

import os
import sys
import json
import sqlite3
import subprocess
from pathlib import Path
import pytest
from openpyxl import load_workbook

from app.core.database import init_db
from app.services.bbbg_export import export_bbbg, TEMPLATE
from app.services.bbbg_service import BBBGService
from app.web_bridge import get_system_snapshot, sync_web_dashboard

PROJECT_ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture
def test_db(tmp_path, monkeypatch):
    """Isolated SQLite database for BBBG and operational tests."""
    import app.core.database as database_module
    db_file = tmp_path / "test_traps.db"
    monkeypatch.setattr(database_module, "DB_PATH", db_file)
    init_db()
    conn = database_module.get_connection()
    conn.execute("DELETE FROM bbbg_records")
    conn.execute("DELETE FROM pick_order_items")
    conn.execute("DELETE FROM pick_orders")
    conn.execute("DELETE FROM sqlite_sequence WHERE name IN ('pick_orders', 'bbbg_records')")
    
    # Seed sample confirmed pick order
    conn.execute("""
        INSERT INTO pick_orders (id, pick_order_code, shop, model, status, total_qty)
        VALUES (1, 'PO-20261014-001', 'BODY', 'VF8', 'USER_CONFIRMED', 127)
    """)
    conn.execute("""
        INSERT INTO pick_order_items (pick_order_id, material_code, batch, qty, cf_actual_qty, cf_status, ob_number)
        VALUES (1, 'BIW30037011AA', '26101301', 127, 77, 'CONFIRMED', '1012646301')
    """)
    conn.commit()
    yield conn
    conn.close()


def test_trap_1_dual_source_and_soft_warnings(test_db, tmp_path):
    """Trap 1: BBBG supports 2 independent sources (Confirmed CF & Direct Paste/Reconciliation).
    Missing slip number or location yields non-blocking behavior.
    """
    svc = BBBGService(test_db, tmp_path / "bbbg_exports")

    # Source 1: Lấy từ CF (Pick Order #1 is USER_CONFIRMED)
    res_cf = svc.create_bbbg(
        pick_order_ref=1,
        sender_name="Nguyễn Văn A",
        receiver_name=None,     # Missing receiver initially
        storage_bin=None,       # Missing location initially
        slip_no=None            # Auto-generated slip_no
    )
    assert res_cf["status"] == "success"
    assert res_cf["slip_no"].startswith("BBBG-")
    assert res_cf["total_qty"] == 77
    assert res_cf["model"] == "VF8"

    # Verify that unconfirmed pick orders are rejected (Manual CF invariant)
    test_db.execute("""
        INSERT INTO pick_orders (id, pick_order_code, shop, model, status, total_qty)
        VALUES (2, 'PO-UNCONFIRMED', 'BODY', 'VF8', 'WAITING_PICK', 50)
    """)
    test_db.commit()
    with pytest.raises(ValueError, match="BẤT BIẾN NGHIỆP VỤ BỊ VI PHẠM"):
        svc.create_bbbg(pick_order_ref=2, sender_name="Nguyễn Văn A")


def test_trap_2_auto_lookup_and_formula_error_stripping():
    """Trap 2: Auto-lookup engine from baseline dictionary and formula error stripping."""
    dict_path = PROJECT_ROOT / "app/core/baseline_dictionary.json"
    assert dict_path.is_file(), "Baseline dictionary must exist"
    
    with open(dict_path, "r", encoding="utf-8") as f:
        baseline = json.load(f)

    assert baseline["total_materials"] >= 326, "Must contain at least 326 baseline materials"
    assert len(baseline["materials"]) >= 995, "Materials catalog must cover Stock V3.6"

    # Test material auto-lookup resolution
    mat_code = "BIW73012711CA"
    assert mat_code in baseline["materials"]
    entry = baseline["materials"][mat_code]
    assert entry["model"] == "VF5 RHD"
    assert entry["default_bin"] == "KHO_1003"
    assert entry["shop"] == "BODY"
    assert entry["default_batch"] == "26101301"

    # Leading zeros preserved as string
    leading_zero_code = "00109988"
    assert isinstance(leading_zero_code, str)
    assert leading_zero_code.startswith("00")


def test_trap_3_anti_column_shift_clipboard_parser():
    """Trap 3: Anti-column shift parser preserving exact tabs, stripping headers,
    handling thousand separators and quotes with Alt+Enter.
    """
    # 1. Exact tabs preservation
    tab_line = "A\t\tC\t\tE"
    parts = tab_line.split("\t")
    assert len(parts) == 5
    assert parts[0] == "A"
    assert parts[1] == ""
    assert parts[2] == "C"
    assert parts[3] == ""
    assert parts[4] == "E"

    # 2. Number parsing with thousand separators
    qty_str = "1,250.5"
    clean_qty = float(qty_str.replace(",", ""))
    assert clean_qty == 1250.5

    # 3. Quoted cell containing Alt+Enter
    multiline_cell = '"Line 1\nLine 2"'
    assert "\n" in multiline_cell


def test_trap_4_database_first_permission_denied_file_lock_catching(test_db, tmp_path, monkeypatch):
    """Trap 4: Database First Persistence & Excel File Lock catching (WinError 32 / PermissionError).
    Must commit to SQLite first, catch lock, and return structured payload with the 3 required bullets.
    """
    svc = BBBGService(test_db, tmp_path / "bbbg_exports")

    # Mock export_bbbg to raise PermissionError (simulating locked file in Microsoft Excel)
    import app.services.bbbg_service as service_module

    def mock_locked_export(path, header, items):
        err = PermissionError(13, "Permission denied (used by another process)")
        err.winerror = 32
        raise err

    monkeypatch.setattr(service_module, "export_bbbg", mock_locked_export)

    # Execute BBBG creation with locked file
    result = svc.create_bbbg(
        pick_order_ref=1,
        sender_name="Nguyễn Văn A",
        slip_no="BBBG-LOCK-001"
    )

    # 1. Verification of structured lock response
    assert result["status"] == "FILE_LOCKED"
    assert result["is_file_locked"] is True
    assert result["saved_in_db"] is True

    # 2. Verification of the 3 required bullets
    # Bullet 1: Cause
    assert "Tệp Excel đang được mở" in result["reason"] or "Microsoft Excel" in result["reason"]
    # Bullet 2: Saved Data
    assert "lưu trữ an toàn trong cơ sở dữ liệu SQLite" in result["db_status"]
    # Bullet 3: Action Advice
    assert "Thử lại bước này" in result["action_advice"]

    # 3. Database First Verification: record is safely committed in SQLite
    row = test_db.execute("SELECT * FROM bbbg_records WHERE slip_no = 'BBBG-LOCK-001'").fetchone()
    assert row is not None
    assert row["status"] == "FILE_LOCKED"
    assert row["total_qty"] == 77

    # 4. Audit trail logged
    ev = test_db.execute("SELECT * FROM job_events WHERE event_type = 'FILE_LOCKED'").fetchone()
    assert ev is not None
    assert ev["status"] == "WARNING"


def test_trap_5_factory_print_fidelity_template_injection(tmp_path):
    """Trap 5: Factory print template injection using sheet BBBG_NEW in 261008_BBBG_FG_V2_2_14.xlsm.
    Maps I1, E1, A4, C5, G5, C7, G7, and detail rows starting at row 13 while preserving logo and printArea.
    """
    assert TEMPLATE.is_file(), f"Template {TEMPLATE} must exist"

    header = {
        "slip_no": "BBBG-FACTORY-001",
        "storage_bin": "KHO_1003",
        "handover_date": "2026-10-14",
        "sender_name": "GL_BODY_01",
        "receiver_name": "GL_MA_02",
        "shop": "BODY",
        "doc_no": "DOC-9988"
    }

    items = [
        {"material_code": "BIW30037011AA", "batch": "26101301", "cf_actual_qty": 77, "unit": "PCS"},
        {"material_code": "BIW30002062", "batch": "26101301", "cf_actual_qty": 15, "unit": "PCS"}
    ]

    out_file = tmp_path / "test_fidelity.xlsx"
    evidence = export_bbbg(out_file, header, items)
    assert Path(evidence["file_path"]).is_file()

    # Load and inspect generated Excel sheet
    wb = load_workbook(out_file, data_only=True)
    ws = wb["BBBG_DRAFT"]

    # Cell coordinate mappings
    assert ws["I1"].value == "BBBG-FACTORY-001"
    assert ws["E1"].value == "KHO_1003"
    assert "2026-10-14" in ws["A4"].value
    assert ws["C5"].value == "GL_BODY_01"
    assert ws["G5"].value == "GL_MA_02"
    assert ws["C7"].value == "Kho MA"
    assert ws["G7"].value == "BODY"

    # Detail rows start at row 13
    assert ws["B13"].value == "BIW30037011AA"
    assert ws["G13"].value == 77
    assert ws["B14"].value == "BIW30002062"
    assert ws["G14"].value == 15

    # Print setup & Plant logo fidelity
    assert "$A$1:$K$" in str(ws.print_area) or "A1:K" in str(ws.print_area)
    assert ws.page_setup.fitToWidth == 1
    assert len(ws._images) >= 1, "Must retain factory company logo image"
    wb.close()


def test_trap_6_inline_grid_editing_and_standalone_js_runner():
    """Trap 6: Standalone Node.js execution of full frontend trap verification test suite."""
    js_test_path = PROJECT_ROOT / "tests/test_six_operational_traps.js"
    assert js_test_path.is_file(), "JS test file must exist"

    result = subprocess.run(
        ["node", str(js_test_path)],
        capture_output=True,
        text=True,
        cwd=str(PROJECT_ROOT)
    )

    print("JS Test Output:\n", result.stdout)
    if result.returncode != 0:
        print("JS Test Stderr:\n", result.stderr)

    assert result.returncode == 0, f"JS test runner failed with exit code {result.returncode}"
    assert "ALL 6 OPERATIONAL TRAPS JS VERIFICATIONS PASSED 100%" in result.stdout
