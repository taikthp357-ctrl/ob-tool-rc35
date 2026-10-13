# -*- coding: utf-8 -*-
"""Test Suite for MA CONTROL 5-Step Flow Invariants & Manual CF Gate."""

import pytest
from fastapi.testclient import TestClient
from app.main import app
from app.models.database import init_db, get_connection

@pytest.fixture(scope="module", autouse=True)
def setup_db():
    init_db()

def test_health_check(monkeypatch):
    monkeypatch.setattr("app.api.routes_config.probe_connection", lambda: {"state": "DISCONNECTED", "message": "No test SAP session"})
    client = TestClient(app)
    resp = client.get("/api/health")
    assert resp.status_code == 200
    data = resp.json()
    assert data["status"] == "healthy"
    assert data["sap_connection"] == "DISCONNECTED"

def test_fetch_ob_items():
    client = TestClient(app)
    resp = client.get("/api/ob/items?branch=BODY_3001")
    assert resp.status_code == 200
    items = resp.json()
    assert isinstance(items, list)
    assert len(items) >= 1

def test_pick_order_creation_invariant_waiting_cf():
    """BẤT BIẾN: Đơn pick sinh ra PHẢI ở trạng thái 'WAITING_USER_CF', không được tự động CF."""
    client = TestClient(app)
    payload = {
        "shop": "BODY",
        "model": "VF8",
        "items": [
            {
                "id": "test-item-101",
                "material_code": "BIW71001001AA",
                "batch": "26101301",
                "required_qty": 100,
                "ob_number": "1012649001"
            }
        ]
    }
    resp = client.post("/api/pick/create", json=payload)
    assert resp.status_code == 200
    data = resp.json()
    assert data["status"] == "success"
    assert data["pick_status"] == "WAITING_USER_CF"
    assert data["status_code"] == "WAITING_USER_CF"
    assert data["status_label"] == "Chờ người dùng CF"
    pick_no = data["pick_order_no"]

    # Verify directly in SQLite
    conn = get_connection()
    row = conn.execute("SELECT status FROM pick_orders WHERE pick_order_no = ?", (pick_no,)).fetchone()
    conn.close()
    assert row["status"] in ("Chờ người dùng CF", "WAITING_USER_CF")

def test_bbbg_blocked_when_not_confirmed():
    """BẤT BIẾN: Tuyệt đối không cho phép tạo BBBG khi đơn pick chưa được người dùng CF."""
    client = TestClient(app)
    # Use the pick order just created
    conn = get_connection()
    row = conn.execute("SELECT pick_order_no FROM pick_orders WHERE status IN ('Chờ người dùng CF', 'WAITING_USER_CF')").fetchone()
    conn.close()
    pick_no = row["pick_order_no"]

    resp = client.post("/api/bbbg/generate", json={
        "pick_order_no": pick_no,
        "sender_name": "GL_BODY_01",
        "receiver_name": "TL_MA_01",
        "storage_bin": "BIN_A1"
    })
    # Must fail with 400 Bad Request
    assert resp.status_code == 400
    err_msg = resp.json()["detail"]
    assert "BẤT BIẾN NGHIỆP VỤ BỊ VI PHẠM" in err_msg

def test_manual_cf_and_bbbg_unlock():
    """Xác nhận thủ công mở khóa tạo BBBG thành công."""
    client = TestClient(app)
    conn = get_connection()
    row = conn.execute("SELECT pick_order_no FROM pick_orders WHERE status IN ('Chờ người dùng CF', 'WAITING_USER_CF')").fetchone()
    conn.close()
    pick_no = row["pick_order_no"]

    # Step 4: Người dùng CF thủ công
    cf_resp = client.post(f"/api/pick/{pick_no}/confirm", json={"confirmed_by": "GL_BODY_01"})
    assert cf_resp.status_code == 400  # Legacy endpoint cannot silently assume quantities
    detail = client.get(f"/api/pick-orders/{pick_no}").json()
    cf_resp = client.post(f"/api/pick-orders/{pick_no}/verify-cf", json={"verified_by": "GL_BODY_01", "actual_quantities": {str(line["id"]): line["qty"] for line in detail["items"]}})
    assert cf_resp.status_code == 200
    assert cf_resp.json()["status"] == "success"

    # Step 5: Bây giờ mới được phép tạo BBBG
    bbbg_resp = client.post("/api/bbbg/generate", json={
        "pick_order_no": pick_no,
        "sender_name": "GL_BODY_01",
        "receiver_name": "TL_MA_01",
        "storage_bin": "BIN_A1"
    })
    assert bbbg_resp.status_code == 200
    bbbg_data = bbbg_resp.json()
    assert bbbg_data["status"] == "success"
    assert bbbg_data["bbbg_no"].startswith("BBBG-")
