# -*- coding: utf-8 -*-
"""Unit & Integration tests for SQLite 10-table Schema, WAL mode, Leading Zeroes & Stock V3.6 Service."""

import pytest
from app.core.database import init_db, get_connection
from app.services.stock_service import StockService

@pytest.fixture(scope="module", autouse=True)
def setup_schema():
    init_db()

def test_ten_tables_and_wal_mode():
    conn = get_connection()
    try:
        mode = conn.execute("PRAGMA journal_mode;").fetchone()[0]
        assert mode.lower() == "wal"

        rows = conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
        tables = {r["name"] for r in rows}
        expected = {
            "materials",
            "model_mappings",
            "stock_snapshots",
            "stock_items",
            "ob_items",
            "pick_orders",
            "pick_order_items",
            "bbbg_records",
            "job_events",
            "settings"
        }
        assert expected.issubset(tables)
    finally:
        conn.close()

def test_leading_zero_preservation_in_stock_and_ob():
    service = StockService()
    res = service.create_snapshot(
        source_type="TEST_ZERO",
        plant="03001",
        sloc="03000",
        items=[
            {
                "material_code": "000123456789",
                "batch": "0026101301",
                "plant": "03001",
                "sloc": "03000",
                "unrestricted_qty": 15.5,
                "unit": "PCS"
            }
        ]
    )
    assert res["status"] == "success"
    items = service.query_stock(material_code="000123456789")
    assert len(items) >= 1
    it = items[0]
    assert it["material_code"] == "000123456789"
    assert it["batch"] == "0026101301"
    assert it["plant"] == "03001"
    assert it["sloc"] == "03000"

def test_stock_v36_zone_isolation_and_pull_line_warning():
    service = StockService()
    service.create_snapshot(
        source_type="SAP_MB52",
        plant="3001",
        sloc="3093",
        items=[
            {"material_code": "MAT_ZONE_TEST", "batch": "26100101", "plant": "3001", "sloc": "3093", "unrestricted_qty": 500.0},
            {"material_code": "MAT_ZONE_TEST", "batch": "26100102", "plant": "3001", "sloc": "3000", "unrestricted_qty": 20.0},
            {"material_code": "MAT_ZONE_TEST", "batch": "26100103", "plant": "3001", "sloc": "3001", "unrestricted_qty": 100.0},
        ]
    )
    summary = service.get_zone_summary(plant="3001")
    by_sloc = {z["sloc"]: z for z in summary}
    assert "3093" in by_sloc
    assert "3000" in by_sloc
    assert "3001" in by_sloc

    # Verify line 3000 generated a pull-line warning because 20.0 < 50.0 threshold
    line_items = service.query_stock(material_code="MAT_ZONE_TEST", sloc="3000", pull_line_only=True)
    assert len(line_items) >= 1
    assert "Cảnh báo kéo line" in line_items[0]["pull_line_warning"]
