"""Regression coverage for the operator CF gate, on isolated databases only."""
import math

import pytest

from app.core.database import init_db, get_connection
from app.services.pick_cf_service import PickCFService
from app.models.schemas import PickVerifyCFRequest


@pytest.fixture
def pick(tmp_path):
    db = tmp_path / "cf.sqlite"
    init_db(db)
    conn = get_connection(db)
    service = PickCFService(conn)
    order = service.create_pick_order("BODY", "VF8", [
        {"material_code": "000123", "batch": "00007", "ob_number": "000001", "req_qty": 10},
        {"material_code": "000123", "batch": "00007", "ob_number": "000002", "req_qty": 20},
    ])
    lines = conn.execute("SELECT * FROM pick_order_items WHERE pick_order_id=? ORDER BY id", (order["id"],)).fetchall()
    yield service, conn, order, lines
    conn.close()


@pytest.mark.parametrize("bad", [None, {}, {"000123": 10}, 0, -1, 9, 11, math.nan, math.inf, True, "10"])
def test_invalid_actuals_leave_order_unchanged(pick, bad):
    service, conn, order, lines = pick
    actuals = bad if bad is None or isinstance(bad, dict) else {str(lines[0]["id"]): bad, str(lines[1]["id"]): 20}
    before = conn.execute("SELECT COUNT(*) FROM job_events").fetchone()[0]
    with pytest.raises(ValueError):
        service.verify_cf(order["id"], "Operator", actuals)
    assert conn.execute("SELECT status FROM pick_orders WHERE id=?", (order["id"],)).fetchone()[0] == "Chờ người dùng CF"
    assert conn.execute("SELECT SUM(line_confirmed) FROM pick_order_items WHERE pick_order_id=?", (order["id"],)).fetchone()[0] == 0
    assert conn.execute("SELECT COUNT(*) FROM job_events").fetchone()[0] == before


def test_explicit_line_actuals_are_saved_and_retry_is_idempotent(pick):
    service, conn, order, lines = pick
    actuals = {str(line["id"]): line["qty"] for line in lines}
    assert service.verify_cf(order["id"], "Operator", actuals)["status"] == "success"
    saved = conn.execute("SELECT * FROM pick_order_items WHERE pick_order_id=? ORDER BY id", (order["id"],)).fetchall()
    assert [line["cf_actual_qty"] for line in saved] == [10, 20]
    assert all(line["material_code"] == "000123" and line["batch"] == "00007" for line in saved)
    before = conn.execute("SELECT COUNT(*) FROM job_events").fetchone()[0]
    assert service.verify_cf(order["id"], "Operator", actuals)["status"] == "already_confirmed"
    assert conn.execute("SELECT COUNT(*) FROM job_events").fetchone()[0] == before


def test_missing_operator_rejected(pick):
    service, _, order, lines = pick
    with pytest.raises(ValueError):
        service.verify_cf(order["id"], " ", {str(line["id"]): line["qty"] for line in lines})


def test_late_line_mismatch_does_not_partially_confirm(pick):
    service, conn, order, lines = pick
    with pytest.raises(ValueError):
        service.verify_cf(order["id"], "Operator", {str(lines[0]["id"]): 10, str(lines[1]["id"]): 19})
    assert not any(line["line_confirmed"] for line in conn.execute(
        "SELECT line_confirmed FROM pick_order_items WHERE pick_order_id=?", (order["id"],)))


def test_extra_line_and_rejected_order_blocked(pick):
    service, conn, order, lines = pick
    actuals = {str(line["id"]): line["qty"] for line in lines}
    with pytest.raises(ValueError):
        service.verify_cf(order["id"], "Operator", {**actuals, "999999": 1})
    with conn:
        conn.execute("UPDATE pick_orders SET status='REJECTED' WHERE id=?", (order["id"],))
    with pytest.raises(ValueError):
        service.verify_cf(order["id"], "Operator", actuals)


def test_request_rejects_boolean_quantity():
    with pytest.raises(ValueError):
        PickVerifyCFRequest(verified_by="Operator", actual_quantities={"1": True})


@pytest.mark.parametrize("qty", [0, -1, math.nan, math.inf, True])
def test_create_rejects_invalid_quantity_without_fallback(pick, qty):
    service, conn, _, _ = pick
    before = conn.execute("SELECT COUNT(*) FROM pick_orders").fetchone()[0]
    with pytest.raises(ValueError):
        service.create_pick_order("BODY", "VF8", [{"material_code": "000123", "ob_number": "000001", "req_qty": qty, "required_qty": 10}])
    assert conn.execute("SELECT COUNT(*) FROM pick_orders").fetchone()[0] == before


def test_create_rejects_repick_and_source_mismatch(pick):
    service, conn, _, lines = pick
    source_id = lines[0]["ob_item_id"]
    payload = {"id": source_id, "material_code": "000123", "batch": "00007", "ob_number": "000001", "req_qty": 10}
    with pytest.raises(ValueError, match="đã có đơn pick"):
        service.create_pick_order("BODY", "VF8", [payload])
    with pytest.raises(ValueError, match="không khớp"):
        service.create_pick_order("BODY", "VF8", [{**payload, "material_code": "OTHER"}])


def test_create_rejects_duplicate_source(pick):
    service, _, _, _ = pick
    payload = {"material_code": "000123", "batch": "00007", "ob_number": "000001", "req_qty": 10}
    with pytest.raises(ValueError, match="bị lặp"):
        service.create_pick_order("BODY", "VF8", [payload, payload])


def test_identical_unlinked_pick_retry_returns_original(pick):
    service, conn, _, _ = pick
    if not any(column["name"] == "source_key" for column in conn.execute("PRAGMA table_info(pick_orders)")):
        conn.execute("ALTER TABLE pick_orders ADD COLUMN source_key TEXT")
        conn.execute("CREATE UNIQUE INDEX idx_pick_source_key ON pick_orders(source_key)")
        conn.commit()
    payload = [{"material_code": "000991", "batch": "00001", "ob_number": "000099", "req_qty": 7}]
    first = service.create_pick_order("BODY", "VF8", payload)
    before = conn.execute("SELECT COUNT(*) FROM job_events").fetchone()[0]
    second = service.create_pick_order("BODY", "VF8", payload)
    assert second["id"] == first["id"]
    assert second["idempotent"] is True
    assert conn.execute("SELECT COUNT(*) FROM job_events").fetchone()[0] == before
    with pytest.raises(ValueError):
        service.create_pick_order("BODY", "VF8", [{**payload[0], "req_qty": 0, "qty": 7}])


def test_simultaneous_unlinked_retries_create_one_pick(tmp_path, monkeypatch):
    from concurrent.futures import ThreadPoolExecutor
    db = tmp_path / "concurrent.sqlite"
    init_db(db)
    conn = get_connection(db)
    with conn:
        if not any(column["name"] == "source_key" for column in conn.execute("PRAGMA table_info(pick_orders)")):
            conn.execute("ALTER TABLE pick_orders ADD COLUMN source_key TEXT")
            conn.execute("CREATE UNIQUE INDEX idx_pick_source_key ON pick_orders(source_key)")
        before = conn.execute("SELECT COUNT(*) FROM pick_orders").fetchone()[0]
    conn.close()
    monkeypatch.setattr("app.services.pick_cf_service.get_connection", lambda: get_connection(db))
    payload = [{"material_code": "000992", "batch": "00001", "ob_number": "000098", "req_qty": 7}]
    def create():
        return PickCFService().create_pick_order("BODY", "VF8", payload)
    with ThreadPoolExecutor(max_workers=2) as pool:
        results = list(pool.map(lambda _: create(), range(2)))
    assert results[0]["id"] == results[1]["id"]
    conn = get_connection(db)
    try:
        assert conn.execute("SELECT COUNT(*) FROM pick_orders").fetchone()[0] == before + 1
    finally:
        conn.close()
