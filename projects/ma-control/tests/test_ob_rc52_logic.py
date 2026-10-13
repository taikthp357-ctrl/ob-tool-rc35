# -*- coding: utf-8 -*-
"""Unit tests for OB Tool RC52 inheritance (Body 3001 & EDC 4007)."""

import pytest
from app.core.database import init_db
from app.services.ob_service import OBService
from app.services.stock_service import StockService

@pytest.fixture(scope="module", autouse=True)
def setup_db():
    init_db()

def test_delete_draft_3001_verification_codes():
    ob_svc = OBService()

    # Invariant: No speculation from suffix (99, 88, 00)
    res_no_guess = ob_svc.delete_draft_3001("1012646399")
    assert res_no_guess["code"] == "DELETE_VERIFIED"
    assert res_no_guess["verified"] is True

    res_ok = ob_svc.delete_draft_3001("1012646321")
    assert res_ok["code"] == "DELETE_VERIFIED"
    assert res_ok["verified"] is True

    res_sto_lock = ob_svc.delete_draft_3001("1012646399", simulated_condition="WAIT_STO_LOCK")
    assert res_sto_lock["code"] == "WAIT_STO_LOCK"
    assert "TUYỆT ĐỐI KHÔNG HỦY 3001" in res_sto_lock["message"]

    res_ob_lock = ob_svc.delete_draft_3001("1012646388", simulated_condition="WAIT_OB_LOCK")
    assert res_ob_lock["code"] == "WAIT_OB_LOCK"

    res_fail = ob_svc.delete_draft_3001("1012646300", simulated_condition="DELETE_FAILED")
    assert res_fail["code"] == "DELETE_FAILED"

def test_dominant_sto_clustering_no_popup():
    ob_svc = OBService()
    materials = ["EDC_MAT_1", "EDC_MAT_2", "EDC_MAT_3"]
    candidates = {
        "EDC_MAT_1": [
            {"sto_number": "4500000001", "due_qty": 50},
            {"sto_number": "4500000002", "due_qty": 100},
        ],
        "EDC_MAT_2": [
            {"sto_number": "4500000002", "due_qty": 80},
        ],
        "EDC_MAT_3": [
            {"sto_number": "4500000001", "due_qty": 30},
            {"sto_number": "4500000002", "due_qty": 90},
        ]
    }
    dominant = ob_svc.calculate_dominant_sto(materials, candidates)
    # 4500000002 satisfies all 3 materials (score=3, due=270) vs 4500000001 (score=2, due=80)
    assert dominant == "4500000002"

def test_replan_deleted_ob():
    ob_svc = OBService()
    res = ob_svc.replan_deleted_ob("EDC88001122AA", "1012647001")
    assert res["action"] == "OB_DELETED_REPLAN"
    assert res["status"] == "REPLANNED"

def test_check_and_fill_ob_items():
    stock_svc = StockService()
    stock_svc.create_snapshot(
        source_type="TEST_OB_CHECK",
        plant="3001",
        sloc="1003",
        items=[
            {"material_code": "BIW73012711CA", "batch": "26101301", "unrestricted_qty": 1000.0},
            {"material_code": "BIW73012036CA", "batch": "26101301", "unrestricted_qty": 1000.0},
        ]
    )
    ob_svc = OBService()
    checked = ob_svc.check_ob_items(branch_tab="BODY_3001", item_ids=[1, 2])
    assert len(checked) == 2
    assert all(it["status"] == "Sẵn sàng điền" for it in checked)

    filled = ob_svc.fill_ob_items(branch_tab="BODY_3001", item_ids=[1, 2])
    assert len(filled) == 2
    assert all(it["status"] == "Chờ pick / CF" and it["ob_number"] for it in filled)
