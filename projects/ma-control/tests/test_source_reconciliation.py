"""Independent regression cases derived from embedded V2.2.14 VBA."""
import copy
import datetime as dt

import pytest

from app.core.database import get_connection, init_db
from app.services.pick_cf_service import PickCFService
from app.services.reconciliation_service import (
    reconcile_rows, save_reconciliation, get_verified_reconciliation,
)


def sources(quantity=10):
    common = {"ob_number": "000001", "material_code": "000123", "batch": "00007", "unit": "EA", "quantity": quantity}
    return [dict(common, movement_type="641", doc_no="000900")], [
        dict(common, ib_number="000800", plant="3001", sloc="1003", gr_status="C")
    ]


def payload(quantity=10):
    mb, vl = sources(quantity)
    meta = {"complete": True, "requested_obs": ["000001"], "captured_at": dt.datetime.now().isoformat(), "filename": "real-export.csv"}
    return {"mb51": mb, "vl06i": vl, "metadata": {"mb51": copy.deepcopy(meta), "vl06i": copy.deepcopy(meta)}}


@pytest.fixture
def confirmed(tmp_path):
    path = tmp_path / "source.sqlite"
    init_db(path)
    conn = get_connection(path)
    service = PickCFService(conn)
    order = service.create_pick_order("BODY", "VF8", [{"material_code": "000123", "batch": "00007", "ob_number": "000001", "req_qty": 10}])
    lines = conn.execute("SELECT * FROM pick_order_items WHERE pick_order_id=?", (order["id"],)).fetchall()
    service.verify_cf(order["id"], "Pilot Operator", {str(x["id"]): x["qty"] for x in lines})
    yield conn, order["id"]
    conn.close()


def test_match_preserves_codes_and_provenance():
    mb, vl = sources()
    result = reconcile_rows(mb, vl, {"000001"})
    assert result[0]["ob_number"] == "000001"
    assert result[0]["batch"] == "00007"
    assert result[0]["quantity"] == 10
    assert result[0]["doc_no"] == "000900"
    assert result[0]["ib_number"] == "000800"


@pytest.mark.parametrize("movement,reverse", [("641", "642"), ("643", "644"), ("645", "646"), ("647", "648"), ("601", "602"), ("351", "352")])
def test_supported_reversal_net_matches(movement, reverse):
    mb, vl = sources()
    mb[0].update(quantity=15, movement_type=movement)
    mb.append(dict(mb[0], quantity=5, movement_type=reverse, doc_no="000901"))
    result = reconcile_rows(mb, vl, {"000001"})
    assert result[0]["quantity"] == 10
    assert result[0]["doc_no"] == "000900, 000901"


@pytest.mark.parametrize("movement", ["101", "102", "999", ""])
def test_unknown_movements_never_become_outbound(movement):
    mb, vl = sources()
    mb[0]["movement_type"] = movement
    with pytest.raises(ValueError):
        reconcile_rows(mb, vl, {"000001"})


def test_unit_mismatch_rejected():
    mb, vl = sources()
    vl[0]["unit"] = "PCS"
    with pytest.raises(ValueError):
        reconcile_rows(mb, vl, {"000001"})


@pytest.mark.parametrize("missing", ["mb51", "vl06i"])
def test_missing_bidirectional_key_rejected(missing):
    mb, vl = sources()
    with pytest.raises(ValueError):
        reconcile_rows([] if missing == "mb51" else mb, [] if missing == "vl06i" else vl, {"000001"})


def test_net_zero_and_no_history_do_not_create_fake_output():
    mb, _ = sources()
    mb.append(dict(mb[0], movement_type="642", doc_no="000901"))
    with pytest.raises(ValueError):
        reconcile_rows(mb, [], {"000001"})
    with pytest.raises(ValueError):
        reconcile_rows([], [], {"000001"})


def test_exact_duplicate_rejected():
    mb, vl = sources()
    with pytest.raises(ValueError):
        reconcile_rows(mb + copy.deepcopy(mb), vl, {"000001"})


def test_mixed_shop_rejected():
    mb, vl = sources()
    mb.append(dict(mb[0], material_code="000124", doc_no="000901"))
    vl.append(dict(vl[0], material_code="000124", plant="4007", sloc="4004", ib_number="000801"))
    with pytest.raises(ValueError):
        reconcile_rows(mb, vl, {"000001"})


@pytest.mark.parametrize("bad_status", [None, "", "D"])
def test_vl_goods_movement_status_required(bad_status):
    mb, vl = sources()
    vl[0]["gr_status"] = bad_status
    with pytest.raises(ValueError):
        reconcile_rows(mb, vl, {"000001"})


def test_vl_only_zero_key_cannot_be_silently_discarded():
    mb, vl = sources()
    vl.append(dict(vl[0], material_code="000124", quantity=0, ib_number="000801"))
    with pytest.raises(ValueError):
        reconcile_rows(mb, vl, {"000001"})


@pytest.mark.parametrize("kind", ["mb51", "vl06i"])
def test_missing_quantity_is_invalid_even_for_extra_zero_row(kind):
    mb, vl = sources()
    rows = mb if kind == "mb51" else vl
    extra = dict(rows[0], material_code="000124")
    extra.pop("quantity")
    rows.append(extra)
    with pytest.raises(ValueError):
        reconcile_rows(mb, vl, {"000001"})


@pytest.mark.parametrize("bad", [True, False, float("inf"), float("nan")])
def test_non_numeric_quantity_is_never_physical_stock(bad):
    mb, vl = sources(1)
    mb[0]["quantity"] = bad
    vl[0]["quantity"] = bad
    with pytest.raises(ValueError):
        reconcile_rows(mb, vl, {"000001"})


@pytest.mark.parametrize("field,value", [("complete", False), ("complete", "true"), ("filename", ""), ("captured_at", ""), ("captured_at", "not a date"), ("requested_obs", ["another-ob"])])
def test_source_metadata_required(confirmed, field, value):
    conn, pick_id = confirmed
    data = payload()
    data["metadata"]["mb51"][field] = value
    with pytest.raises(ValueError):
        save_reconciliation(conn, pick_id, data)
    assert conn.execute("SELECT COUNT(*) FROM source_reconciliations").fetchone()[0] == 0


def test_cf_and_source_quantities_must_match(confirmed):
    conn, pick_id = confirmed
    with pytest.raises(ValueError):
        save_reconciliation(conn, pick_id, payload(9))


def test_revalidation_blocks_changed_pick(confirmed):
    conn, pick_id = confirmed
    first = save_reconciliation(conn, pick_id, payload())
    assert get_verified_reconciliation(conn, pick_id)["source_hash"] == first["source_hash"]
    conn.execute("UPDATE pick_order_items SET qty=11,cf_actual_qty=11 WHERE pick_order_id=?", (pick_id,))
    with pytest.raises(ValueError):
        get_verified_reconciliation(conn, pick_id)


def test_revalidation_blocks_stale_source(confirmed):
    import json
    conn, pick_id = confirmed
    evidence = save_reconciliation(conn, pick_id, payload())
    evidence["sources_metadata"]["mb51"]["captured_at"] = (dt.datetime.now() - dt.timedelta(days=2)).isoformat()
    conn.execute("UPDATE source_reconciliations SET evidence_json=? WHERE pick_order_id=?", (json.dumps(evidence), pick_id))
    with pytest.raises(ValueError):
        get_verified_reconciliation(conn, pick_id)
