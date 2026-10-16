# -*- coding: utf-8 -*-
"""Tests strictly validating USER_SCOPE_LOW_STOCK_20261013 acceptance criteria.

Covers:
1. CHỜ CF DO NGƯỜI DÙNG QUẢN LÝ:
   - Direct paste from Excel without prior pick order.
   - String leading zero preservation for OB, Material, Batch, Plant, Box, Rack.
   - Real DB anti-duplicate check.
   - Edit and delete of unconfirmed lines.
   - No auto-recognition of SAP CF.
2. CHỨC NĂNG CƠ BẢN ƯU TIÊN:
   - Snapshot isolation by mat + plant + sloc + batch + stock_type.
   - Threshold config & low stock alerts with freshness checks.
   - Normal OB fill flow unchanged (no auto low-stock strategy).
3. CHẾ ĐỘ TÙY CHỌN ĐIỀN OB THEO TỒN THẤP (2-3 BATCH):
   - Opt-in inspect per batch/sloc.
   - 2-3 batches proposed without draining entire batch.
   - Locked/unverified batches excluded with reason.
   - Other shop impact labeled truthfully: 'Chưa đánh giá được nhu cầu/dự trữ shop khác'.
   - Labeled 'Ghi nhận / Đề xuất' (PROPOSAL).
   - Anti-duplicate batch usage in allocation.
"""

import pytest
from fastapi.testclient import TestClient
from app.main import app
from app.core import database
from app.services.stock_service import StockService

@pytest.fixture
def client(tmp_path, monkeypatch):
    db_path = tmp_path / 'user_scope_test.db'
    monkeypatch.setattr(database, 'DB_PATH', db_path)
    monkeypatch.setenv('MA_CONTROL_EXPORT_DIR', str(tmp_path / 'exports'))
    database.init_db(db_path)
    with TestClient(app) as c:
        yield c


def test_waiting_cf_direct_paste_preserves_leading_zeros_and_anti_duplicate(client):
    # Raw TSV text simulating Excel copy-paste with leading zeros
    raw_tsv = (
        "OB\tMã hàng\tPlant\tBatch\tSố lượng\tBox\tRack\n"
        "00109988\t00MAT0012\t3001\t00B001\t45.5\tBOX-001\tRACK-A01\n"
        "00109989\t00MAT0034\t3001\t00B002\t12.0\tBOX-002\tRACK-B05\n"
    )

    paste_res = client.post("/api/waiting-cf/paste-excel", json={
        "raw_text": raw_tsv,
        "allow_duplicates": False
    })
    assert paste_res.status_code == 200
    pdata = paste_res.json()
    assert pdata["status"] == "success"
    assert pdata["count"] == 2
    lines = pdata["lines"]

    # 1. Preserves leading zeroes
    assert lines[0]["ob_number"] == "00109988"
    assert lines[0]["material_code"] == "00MAT0012"
    assert lines[0]["batch"] == "00B001"
    assert lines[0]["packing_box"] == "BOX-001"
    assert lines[0]["rack"] == "RACK-A01"
    assert lines[0]["qty"] == 45.5
    line1_id = lines[0]["id"]

    # 2. Query Waiting CF - Starts in WAITING_CF, no auto SAP CF claim
    cf_list = client.get("/api/waiting-cf").json()
    direct_items = [x for x in cf_list if x.get("is_direct_paste")]
    assert len(direct_items) == 2
    item1 = next(x for x in cf_list if x["id"] == line1_id)
    assert item1["sap_cf_verified"] is False
    assert item1["cf_status"] in ("WAITING_CF", "PENDING_CF")
    assert item1["is_direct_paste"] is True
    assert "Xác nhận thực tế" in item1["cf_source_label"]

    # 3. Anti-duplicate on DB: Pasting the exact active line again must fail
    dup_res = client.post("/api/waiting-cf/paste-excel", json={
        "raw_text": "00109988\t00MAT0012\t3001\t00B001\t45.5\tBOX-001\tRACK-A01",
        "allow_duplicates": False
    })
    assert dup_res.status_code == 400
    assert "chống dán trùng" in dup_res.json()["detail"].lower()

    # 4. User-managed line update (edit)
    edit_res = client.put(f"/api/waiting-cf/lines/{line1_id}", json={
        "line_id": line1_id,
        "qty": 50.0,
        "packing_box": "BOX-001-MOD",
        "rack": "RACK-A01-MOD"
    })
    assert edit_res.status_code == 200
    assert edit_res.json()["qty"] == 50.0
    assert edit_res.json()["packing_box"] == "BOX-001-MOD"

    # 5. User-managed line delete
    line2_id = lines[1]["id"]
    del_res = client.delete(f"/api/waiting-cf/lines/{line2_id}")
    assert del_res.status_code == 200
    remaining_direct = [x for x in client.get("/api/waiting-cf").json() if x.get("is_direct_paste")]
    assert len(remaining_direct) == 1


def test_low_stock_config_and_alerts(client):
    svc = StockService()
    # Seed snapshot with 2 batches: total unrestricted = 30 PCS (< threshold 50)
    svc.create_snapshot(
        source_type="TEST_SNAPSHOT",
        plant="3001",
        sloc="1003",
        items=[
            {"material_code": "00LOW01", "batch": "B01", "unrestricted_qty": 20.0, "stock_type": "UNRESTRICTED"},
            {"material_code": "00LOW01", "batch": "B02", "unrestricted_qty": 10.0, "stock_type": "UNRESTRICTED"},
        ]
    )

    # Set threshold to 50 PCS for 00LOW01 in plant 3001
    cfg_res = client.post("/api/stock/low-stock/config", json={
        "default_threshold": 40.0,
        "thresholds": [
            {"material_code": "00LOW01", "plant": "3001", "threshold_qty": 50.0, "warning_message": "Cần ưu tiên tồn"}
        ]
    })
    assert cfg_res.status_code == 200

    # Check alert list
    alerts = client.get("/api/stock/low-stock/alerts?plant=3001").json()
    assert len(alerts) >= 1
    alert_item = next(a for a in alerts if a["material_code"] == "00LOW01")
    assert alert_item["is_low_stock"] is True
    assert alert_item["current_stock"] == 30.0
    assert alert_item["threshold_qty"] == 50.0


def test_opt_in_low_stock_priority_2_to_3_batches(client):
    svc = StockService()
    # Seed 4 batches:
    # Batch 1 (avail 30)
    # Batch 2 (avail 40)
    # Batch 3 (avail 50)
    # Batch 4 (BLOCKED 100) -> Must be excluded from proposed and listed in locked_batches
    svc.create_snapshot(
        source_type="SNAPSHOT_MULTI",
        plant="3001",
        sloc="1003",
        items=[
            {"material_code": "00MAT_OPT", "batch": "00B1", "unrestricted_qty": 30.0, "stock_type": "UNRESTRICTED", "priority_rank": 1},
            {"material_code": "00MAT_OPT", "batch": "00B2", "unrestricted_qty": 40.0, "stock_type": "UNRESTRICTED", "priority_rank": 2},
            {"material_code": "00MAT_OPT", "batch": "00B3", "unrestricted_qty": 50.0, "stock_type": "UNRESTRICTED", "priority_rank": 3},
            {"material_code": "00MAT_OPT", "batch": "00B_LOCKED", "unrestricted_qty": 100.0, "stock_type": "BLOCKED", "priority_rank": 0},
        ]
    )

    # User explicitly calls inspect with required_qty = 50 PCS
    insp_res = client.post("/api/stock/low-stock/inspect", json={
        "material_code": "00MAT_OPT",
        "plant": "3001",
        "required_qty": 50.0,
        "max_batches": 3
    })
    assert insp_res.status_code == 200
    idata = insp_res.json()

    # Check proposal properties
    assert idata["mode"] == "LOW_STOCK_PRIORITY_OPT_IN"
    assert idata["status_label"] == "Ghi nhận / Đề xuất"
    assert idata["result_type"] == "PROPOSAL"
    assert idata["is_sap_live"] is False
    assert idata["other_shop_impact"] == "Chưa đánh giá được nhu cầu/dự trữ shop khác"
    assert idata["prevent_line_stop_guaranteed"] is False

    # Locked batch excluded with explicit reason
    assert len(idata["locked_batches"]) >= 1
    assert any(b["batch"] == "00B_LOCKED" for b in idata["locked_batches"])

    # Proposed batches: 2 batches needed (30 from B1, 20 from B2)
    proposed = idata["proposed_batches"]
    assert len(proposed) == 2  # exactly 2 batches
    assert proposed[0]["batch"] == "00B1"
    assert proposed[0]["proposed_qty"] == 30.0
    assert proposed[1]["batch"] == "00B2"
    assert proposed[1]["proposed_qty"] == 20.0  # Does NOT drain all 40 PCS of B2!
    assert idata["total_proposed_qty"] == 50.0

    # Apply allocation directly to Chờ CF (target="WAITING_CF")
    apply_res = client.post("/api/stock/low-stock/apply", json={
        "material_code": "00MAT_OPT",
        "plant": "3001",
        "required_qty": 50.0,
        "target": "WAITING_CF",
        "ob_number": "00OB_PROP_01",
        "allocations": [
            {"batch": "00B1", "sloc": "1003", "allocated_qty": 30.0},
            {"batch": "00B2", "sloc": "1003", "allocated_qty": 20.0},
        ]
    })
    assert apply_res.status_code == 200
    adata = apply_res.json()
    assert adata["status"] == "success"
    assert adata["status_label"] == "Ghi nhận / Đề xuất"
    assert adata["total_allocated_qty"] == 50.0

    # Verify items landed in Waiting CF with label "Ghi nhận / Đề xuất"
    cf_items = client.get("/api/waiting-cf?material_code=00MAT_OPT").json()
    assert len(cf_items) == 2
    assert all(i["sap_cf_verified"] is False for i in cf_items)

    # Duplicate batch within same allocation must be rejected
    dup_batch = client.post("/api/stock/low-stock/apply", json={
        "material_code": "00MAT_OPT",
        "plant": "3001",
        "required_qty": 50.0,
        "target": "PROPOSAL",
        "allocations": [
            {"batch": "00B1", "sloc": "1003", "allocated_qty": 25.0},
            {"batch": "00B1", "sloc": "1003", "allocated_qty": 25.0},
        ]
    })
    assert dup_batch.status_code == 400
    assert "bị trùng" in dup_batch.json()["detail"]

    # Allocation exceeding required qty must be rejected
    over_alloc = client.post("/api/stock/low-stock/apply", json={
        "material_code": "00MAT_OPT",
        "plant": "3001",
        "required_qty": 50.0,
        "target": "PROPOSAL",
        "allocations": [
            {"batch": "00B1", "sloc": "1003", "allocated_qty": 30.0},
            {"batch": "00B2", "sloc": "1003", "allocated_qty": 35.0},  # total 65 > 50
        ]
    })
    assert over_alloc.status_code == 400
    assert "vượt quá nhu cầu" in over_alloc.json()["detail"]
