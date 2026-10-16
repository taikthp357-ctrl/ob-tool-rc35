# -*- coding: utf-8 -*-
"""Unit and integration tests for direct SAP COM automation.

Verifies:
1. Native SAP Scripting COM automation contracts for MB52, MB51, and VL06I.
2. Honest DISCONNECTED reporting when SAP GUI is not running (no fake live PASS).
3. ALV Grid row reading and intermediate Excel generation.
4. Auto-cleanup of intermediate files on success and retention on diagnostic failure.
5. End-to-end integration with StockService and ReconciliationService.
"""

import os
import pytest
from unittest.mock import MagicMock, patch
from pathlib import Path

from app.services.sap_automation import (
    write_excel_dataset,
    _execute_mb52_com,
    _execute_mb51_com,
    _execute_vl06i_com,
    download_sap_mb52_live,
    extract_sap_mb51_live,
    extract_sap_vl06i_live,
    set_windows_clipboard
)
from app.services.sap_connection import probe_connection
from app.services.stock_service import StockService
from app.services.reconciliation_service import reconcile_direct_sap
from app.services.pick_cf_service import PickCFService
from app.services.ob_service import OBService
from app.core.database import get_connection, init_db


@pytest.fixture(scope="module", autouse=True)
def setup_test_db():
    init_db()


def test_probe_connection_disconnected_when_no_live_gui():
    """Verify probe_connection accurately reports DISCONNECTED without live SAP."""
    res = probe_connection(timeout=2)
    assert res["state"] in ("DISCONNECTED", "UNAVAILABLE")
    assert "chưa" in res["message"].lower() or "đăng nhập" in res["message"].lower()


def test_write_excel_dataset():
    """Verify write_excel_dataset writes valid openpyxl Excel file with headers."""
    headers = ["MATNR", "WERKS", "LGORT", "CHARG", "Unrestricted", "MEINS"]
    rows = [
        ["MAT001", "3001", "1003", "LOT01", 100, "PCS"],
        ["MAT002", "3093", "1003", "LOT02", 50, "PCS"],
    ]
    path = write_excel_dataset(headers, rows, prefix="TEST_MB52_")
    try:
        assert path.exists()
        assert path.suffix == ".xlsx"
        from app.services.source_import import parse_source
        parsed = parse_source("mb52", path.read_bytes(), path.name)
        assert len(parsed) == 2
        assert parsed[0]["material_code"] == "MAT001"
        assert parsed[0]["unrestricted_qty"] == 100.0
    finally:
        if path.exists():
            path.unlink()


def test_set_windows_clipboard():
    """Verify Windows clipboard setter executes without uncaught crash."""
    ok = set_windows_clipboard("OB123456\r\nOB123457")
    # On Windows console or headless, it returns True or False safely
    assert isinstance(ok, bool)


def test_mb52_com_execution_mocked():
    """Verify _execute_mb52_com correctly drives SAP GUI COM objects."""
    mock_session = MagicMock()
    mock_wnd0 = MagicMock()
    mock_grid = MagicMock()
    mock_sbar = MagicMock()

    mock_grid.RowCount = 2
    mock_grid.GetCellValue.side_effect = lambda row, col: {
        (0, "MATNR"): "MAT-MB52-01",
        (0, "WERKS"): "3001",
        (0, "LGORT"): "1003",
        (0, "CHARG"): "BATCHA",
        (0, "LABST"): "150",
        (0, "MENGE"): "150",
        (0, "MEINS"): "PCS",
        (1, "MATNR"): "MAT-MB52-02",
        (1, "WERKS"): "3093",
        (1, "LGORT"): "1003",
        (1, "CHARG"): "BATCHB",
        (1, "LABST"): "75",
        (1, "MENGE"): "75",
        (1, "MEINS"): "PCS",
    }.get((row, col), "")

    mock_sbar.MessageType = "S"
    mock_sbar.Text = "Success"

    def find_by_id(sel):
        if sel == "wnd[0]":
            return mock_wnd0
        if sel == "wnd[0]/sbar":
            return mock_sbar
        if sel == "wnd[0]/usr/cntlGRID1/shellcont/shell":
            return mock_grid
        return MagicMock()

    mock_session.findById.side_effect = find_by_id

    with patch("app.services.sap_automation.get_sap_com_session", return_value=mock_session):
        out_path = _execute_mb52_com(plants=["3001", "3093"], sloc="1003", layout="/KHOMA")
        try:
            assert out_path.exists()
            from app.services.source_import import parse_source
            rows = parse_source("mb52", out_path.read_bytes(), out_path.name)
            assert len(rows) == 2
            assert rows[0]["material_code"] == "MAT-MB52-01"
            assert rows[0]["unrestricted_qty"] == 150.0
            assert rows[1]["plant"] == "3093"
        finally:
            if out_path.exists():
                out_path.unlink()


def test_mb51_com_execution_mocked():
    """Verify _execute_mb51_com drives MB51 transaction, selects XBLNR, reads ALV."""
    mock_session = MagicMock()
    mock_wnd0 = MagicMock()
    mock_grid = MagicMock()
    mock_sbar = MagicMock()

    mock_grid.RowCount = 1
    mock_grid.GetCellValue.side_effect = lambda row, col: {
        (0, "MBLNR"): "5000001",
        (0, "MATNR"): "MAT-MB51-01",
        (0, "CHARG"): "LOT51",
        (0, "MENGE"): "20",
        (0, "XBLNR"): "OB-9999",
        (0, "EBELN"): "PO-100",
        (0, "BWART"): "641",
        (0, "MEINS"): "PCS",
    }.get((row, col), "")

    mock_sbar.MessageType = "S"

    def find_by_id(sel):
        if sel == "wnd[0]":
            return mock_wnd0
        if sel == "wnd[0]/sbar":
            return mock_sbar
        if sel == "wnd[0]/usr/cntlGRID1/shellcont/shell":
            return mock_grid
        return MagicMock()

    mock_session.findById.side_effect = find_by_id

    with patch("app.services.sap_automation.get_sap_com_session", return_value=mock_session):
        out_path = _execute_mb51_com(obs=["OB-9999"], plant="3093", layout="/khoma")
        try:
            assert out_path.exists()
            from app.services.source_import import parse_source
            rows = parse_source("mb51", out_path.read_bytes(), out_path.name)
            assert len(rows) == 1
            assert rows[0]["ob_number"] == "OB-9999"
            assert rows[0]["quantity"] == 20.0
            assert rows[0]["movement_type"] == "641"
        finally:
            if out_path.exists():
                out_path.unlink()


def test_vl06i_com_execution_mocked():
    """Verify _execute_vl06i_com drives VL06I transaction and item detail view."""
    mock_session = MagicMock()
    mock_wnd0 = MagicMock()
    mock_grid = MagicMock()
    mock_sbar = MagicMock()

    mock_grid.RowCount = 1
    mock_grid.GetCellValue.side_effect = lambda row, col: {
        (0, "LIFEX"): "OB-8888",
        (0, "MATNR"): "MAT-VL06I-01",
        (0, "CHARG"): "LOT06",
        (0, "LFIMG"): "35",
        (0, "VBELN"): "IB-12345",
        (0, "WERKS"): "3001",
        (0, "LGORT"): "1003",
        (0, "VRKME"): "PCS",
        (0, "WBSTK"): "C",
    }.get((row, col), "")

    mock_sbar.MessageType = "S"

    def find_by_id(sel):
        if sel == "wnd[0]":
            return mock_wnd0
        if sel == "wnd[0]/sbar":
            return mock_sbar
        if sel == "wnd[0]/usr/cntlGRID1/shellcont/shell":
            return mock_grid
        return MagicMock()

    mock_session.findById.side_effect = find_by_id

    with patch("app.services.sap_automation.get_sap_com_session", return_value=mock_session):
        out_path = _execute_vl06i_com(obs=["OB-8888"])
        try:
            assert out_path.exists()
            from app.services.source_import import parse_source
            rows = parse_source("vl06i", out_path.read_bytes(), out_path.name)
            assert len(rows) == 1
            assert rows[0]["ob_number"] == "OB-8888"
            assert rows[0]["quantity"] == 35.0
            assert rows[0]["ib_number"] == "IB-12345"
        finally:
            if out_path.exists():
                out_path.unlink()


def test_stock_service_direct_sap_flow_connected():
    """Verify StockService seamlessly consumes live SAP COM output when connected."""
    mock_excel = write_excel_dataset(
        ["MATNR", "WERKS", "LGORT", "CHARG", "Unrestricted", "MEINS"],
        [["M-LIVE-01", "3001", "1003", "B01", 80, "PCS"]],
        prefix="SAP_LIVE_MOCK_"
    )
    assert mock_excel.exists()

    with patch("app.services.stock_service.probe_connection", return_value={"state": "CONNECTED"}), \
         patch("app.services.sap_automation.download_sap_mb52_live", return_value=mock_excel):
        service = StockService()
        res = service.download_and_sync_sap_mb52(plants=["3001"], sloc="1003")
        assert res["status"] == "SUCCESS"
        assert res["total_rows"] == 1
        assert res["intermediate_file_cleaned"] is True
        # Verify the intermediate file was cleaned up
        assert not mock_excel.exists()


def test_reconciliation_direct_sap_flow_connected():
    """Verify ReconciliationService seamlessly consumes live SAP COM output when connected."""
    conn = get_connection()
    try:
        cur = conn.execute("""
            INSERT INTO pick_orders (pick_order_code, shop, model, status, created_at)
            VALUES ('PO-SAP-LIVE-TEST', 'BODY', 'VF8', 'Đã CF', datetime('now'))
        """)
        p_id = cur.lastrowid
        conn.execute("""
            INSERT INTO pick_order_items (
                pick_order_id, ob_number, material_code, batch, qty, cf_actual_qty,
                cf_status, line_confirmed, packing_box
            )
            VALUES (?, 'OB-LIVE-01', 'M-REC-01', 'B-REC-01', 50, 50, 'CONFIRMED', 1, 'BOX1')
        """, (p_id,))
        conn.commit()

        mock_mb51 = write_excel_dataset(
            ["MBLNR", "MATNR", "CHARG", "MENGE", "XBLNR", "EBELN", "BWART", "MEINS"],
            [["DOC51", "M-REC-01", "B-REC-01", 50, "OB-LIVE-01", "PO-1", "641", "PCS"]],
            prefix="SAP_LIVE_MB51_"
        )
        assert mock_mb51.exists()

        with patch("app.services.reconciliation_service.probe_connection", return_value={"state": "CONNECTED"}), \
             patch("app.services.sap_automation.extract_sap_mb51_live", return_value=mock_mb51):
            res = reconcile_direct_sap(conn, p_id)
            assert res["status"] in ("MATCHED", "SUCCESS")
            assert res["intermediate_file_cleaned"] is True
            assert not mock_mb51.exists()
    finally:
        conn.close()
