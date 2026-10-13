# -*- coding: utf-8 -*-
"""Tests for Step 3 (Pick Order) & Step 4 (Manual CF Verification Gate)."""

import pytest
from fastapi.testclient import TestClient
from app.main import app
from app.core.database import init_db

@pytest.fixture(scope="module", autouse=True)
def setup_db():
    init_db()

def test_step3_and_step4_verify_cf_endpoint():
    client = TestClient(app)

    # Step 3: Create Pick Order
    create_resp = client.post("/api/pick-orders", json={
        "branch_tab": "BODY_3001",
        "shop": "BODY",
        "model": "VF7",
        "items": [
            {
                "material_code": "BIW70002022BA",
                "batch": "0026101399",
                "req_qty": 80,
                "ob_number": "001012648888"
            }
        ]
    })
    assert create_resp.status_code == 200
    po_data = create_resp.json()
    po_id = po_data["id"]
    po_code = po_data["pick_order_code"]
    assert po_code.startswith("PICK-")
    assert po_data["status_label"] == "Chờ người dùng CF"

    # Fetch detail before CF -> Must be 'Chờ người dùng CF'
    detail_before = client.get(f"/api/pick-orders/{po_id}").json()
    assert detail_before["status"] == "Chờ người dùng CF"
    assert detail_before["items"][0]["batch"] == "0026101399"
    assert detail_before["items"][0]["ob_number"] == "001012648888"

    rejected = client.post(f"/api/pick-orders/{po_id}/verify-cf", json={"verified_by": "QA_Tester_Nova", "actual_quantities": {str(detail_before["items"][0]["id"]): 0}})
    assert rejected.status_code == 400
    assert client.get(f"/api/pick-orders/{po_id}").json()["status"] == "Chờ người dùng CF"

    # Step 4: Trigger POST /api/pick-orders/{id}/verify-cf
    cf_resp = client.post(f"/api/pick-orders/{po_id}/verify-cf", json={
        "verified_by": "QA_Tester_Nova",
        "actual_quantities": {str(detail_before["items"][0]["id"]): 80},
        "notes": "Đã kiểm tra thực tế đủ 80 PCS"
    })
    assert cf_resp.status_code == 200
    cf_data = cf_resp.json()
    assert cf_data["cf_status"] == "Đã CF"
    assert cf_data["verified_by"] == "QA_Tester_Nova"

    # Fetch detail after CF -> Must be 'Đã CF'
    detail_after = client.get(f"/api/pick-orders/{po_id}").json()
    assert detail_after["status"] == "Đã CF"
    assert detail_after["cf_verified_by"] == "QA_Tester_Nova"
