# -*- coding: utf-8 -*-
"""Tests for Step 5 (BBBG V2.2.14) and SAP Worker Idempotency / Recovery."""

import uuid
import pytest
from app.core.database import init_db, get_connection
from app.services.pick_cf_service import PickCFService
from app.services.bbbg_service import BBBGService
from app.services.sap_worker import SAPWorkerEngine, sanitize_log_text

@pytest.fixture(scope="module", autouse=True)
def setup_db():
    init_db()

def actuals_for(order_id):
    conn = get_connection()
    try:
        return {str(row["id"]): row["qty"] for row in conn.execute("SELECT id, qty FROM pick_order_items WHERE pick_order_id=?", (order_id,))}
    finally:
        conn.close()

def test_bbbg_limo_vf7_mapping_and_dedup_append_only():
    pick_svc = PickCFService()
    bbbg_svc = BBBGService()

    # Create pick order with BIW71001001AA (should map to VF8 LIMO)
    po = pick_svc.create_pick_order(
        shop="BODY",
        model="UNKNOWN",
        items=[{"material_code": "BIW71001001AA", "batch": "26101305", "req_qty": 60, "ob_number": "1012649111"}]
    )
    po_id = po["id"]

    # Attempting BBBG before CF must raise ValueError
    with pytest.raises(ValueError, match="BẤT BIẾN NGHIỆP VỤ BỊ VI PHẠM"):
        bbbg_svc.create_bbbg(pick_order_ref=po_id, sender_name="GL_01")

    # Verify CF manually
    pick_svc.verify_cf(po_id, verified_by="GL_01", actual_quantities=actuals_for(po_id))

    # Create BBBG with explicit slip_no and existing duplicate VAT attachment
    year_str = po["pick_order_code"].split("-")[1][:4]
    dup_vat_name = f"VAT_{po['pick_order_code']}_{year_str}.pdf"
    unique_slip = f"SLIP-UNIQUE-{uuid.uuid4().hex[:8].upper()}"
    res = bbbg_svc.create_bbbg(
        pick_order_ref=po_id,
        sender_name="GL_01",
        receiver_name="TL_02",
        storage_bin="LINE_A",
        slip_no=unique_slip,
        existing_vat_attachments=[dup_vat_name]
    )
    assert res["status"] == "success"
    assert res["model"] == "VF8 LIMO"
    assert res["gos_bbbg_attached"] == 0
    assert res["general_log_saved"] == 0
    assert res["is_draft"] is True
    # Duplicate VAT skipped
    assert res["gos_vat_attached"] == 0

    # Append-only dedup check: re-creating with same slip_no must be rejected
    po2 = pick_svc.create_pick_order(
        shop="BODY",
        model="VF8 LIMO",
        items=[{"material_code": "BIW71001001AA", "batch": "26101305", "req_qty": 30, "ob_number": "1012649112"}]
    )
    pick_svc.verify_cf(po2["id"], verified_by="GL_01", actual_quantities=actuals_for(po2["id"]))
    with pytest.raises(ValueError, match="Khóa sổ tổng đã tồn tại"):
        bbbg_svc.create_bbbg(pick_order_ref=po2["id"], sender_name="GL_01", slip_no=unique_slip)

def test_sap_worker_idempotent_lock_resume_and_password_sanitization():
    worker = SAPWorkerEngine()

    # 1. Password scrubbing
    dirty = 'Connecting user=ADMIN password=SuperSecret123! token="abcdef999"'
    clean = sanitize_log_text(dirty)
    assert "SuperSecret123!" not in clean
    assert "abcdef999" not in clean
    assert "***REDACTED***" in clean

    # 2. Idempotent lock prevents duplicate concurrent execution
    bkey = "RUN_ALL_BATCH_20261013_01"
    worker.acquire_idempotent_lock(bkey)
    with pytest.raises(RuntimeError, match="IDEMPOTENT_LOCK_CONFLICT"):
        worker.acquire_idempotent_lock(bkey)
    worker.release_idempotent_lock(bkey)

    # 3. Pipeline execution & resume from failed step
    call_counter = {"step1": 0, "step2": 0}

    def step1_fn():
        call_counter["step1"] += 1
        return {"sap_status_msg": "Step 1 OK"}

    def step2_fail_once():
        call_counter["step2"] += 1
        if call_counter["step2"] == 1:
            raise RuntimeError("SAP Session Timeout")
        return {"sap_status_msg": "Step 2 Recovered"}

    steps = [
        {"step_name": "READ_MB51", "action": step1_fn},
        {"step_name": "POST_VL02N", "action": step2_fail_once}
    ]

    pipeline_key = f"PIPELINE_RESUME_KEY_{uuid.uuid4().hex[:8]}"
    first_run = worker.execute_pipeline(pipeline_key, "OB_BATCH", steps, resume=False)
    assert first_run["status"] == "FAILED"
    assert first_run["failed_step"] == "POST_VL02N"
    assert call_counter["step1"] == 1
    assert call_counter["step2"] == 1

    # Resume: Step 1 must be skipped, Step 2 executed and succeeded
    second_run = worker.execute_pipeline(pipeline_key, "OB_BATCH", steps, resume=True)
    assert second_run["status"] == "SUCCEEDED"
    assert call_counter["step1"] == 1  # Not called again!
    assert call_counter["step2"] == 2
