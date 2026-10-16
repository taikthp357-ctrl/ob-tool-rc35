import datetime
import json
import sqlite3
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from app.main import app
from app.core import database
from app.services.stock_service import StockService
from app.services.ob_service import OBService
from app.services.source_import import parse_source
from app.models.schemas import normalize_ob_status


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, 'DB_PATH', tmp_path / 'pilot.db')
    monkeypatch.setenv('MA_CONTROL_EXPORT_DIR', str(tmp_path / 'exports'))
    database.init_db(tmp_path / 'pilot.db')
    with TestClient(app) as client:
        yield client


def test_snapshot_replacement_preserves_zero_and_empty_scope(client):
    svc = StockService()
    item = {'material_code': '00001', 'batch': '00002', 'unrestricted_qty': 10}
    svc.create_snapshot('first', '3001', '1003', [item])
    svc.create_snapshot('second', '3001', '1003', [{**item, 'unrestricted_qty': 0, 'qty': 99}])
    assert svc.query_stock(material_code='00001')[0]['unrestricted_qty'] == 0
    svc.create_snapshot('third', '3001', '1003', [{**item, 'unrestricted_qty': 4}])
    assert svc.get_zone_summary(plant='3001')[0]['total_qty'] == 4
    svc.create_snapshot('empty', '3001', '1003', [])
    assert svc.query_stock(material_code='00001') == []
    conn = database.get_connection()
    assert conn.execute('SELECT COUNT(*) FROM stock_snapshots').fetchone()[0] == 4
    conn.close()


def test_ob_stock_checks_isolate_zone_batch_and_reserve_selected_rows(client):
    svc = StockService()
    for plant, sloc, batch, qty in [('3001','1003','00002',10), ('3001','3000','00002',100), ('3001','1003','OTHER',999)]:
        svc.create_snapshot('source', plant, sloc, [{'material_code':'00001','batch':batch,'qty':qty}])
    # Latest 1003 snapshot must include all batches of that scope, not sum history.
    svc.create_snapshot('complete', '3001', '1003', [{'material_code':'00001','batch':'00002','qty':10}, {'material_code':'00001','batch':'OTHER','qty':999}])
    saved = client.post('/api/ob/batch-save', json={'items':[{'item_uid':str(i),'material_code':'00001','batch':'00002','req_qty':7,'ob_number':f'00000{i}'} for i in range(2)]})
    ids = saved.json()['ids']
    result = client.post('/api/ob/check', json={'branch_tab':'BODY_3001','item_ids':ids}).json()['items']
    assert result[0]['avail_qty'] == 10
    assert result[0]['status'] == 'Sẵn sàng điền'
    assert result[1]['status'] == 'Thiếu SL OB' and result[1]['avail_qty'] == 3


def test_batch_save_prevents_cross_tab_overwrite_and_client_cf_claim(client):
    row = {'item_uid':'same','material_code':'00001','batch':'00002','required_qty':1,'status':'Đã CF'}
    body = client.post('/api/ob/batch-save', json={'branch_tab':'BODY_3001','items':[row]}).json()['ids'][0]
    edc = client.post('/api/ob/batch-save', json={'branch_tab':'EDC_4007','items':[row]}).json()['ids'][0]
    assert body != edc
    assert next(i for i in client.get('/api/ob/items?branch_tab=BODY_3001').json() if i['id'] == body)['status'] == 'Chưa kiểm tra'
    assert client.post('/api/ob/batch-save', json={'branch_tab':'EDC_4007','items':[{**row,'id':body}]}).status_code == 400


def test_no_synthetic_ob_or_draft_delete_in_manual_mode(client, monkeypatch):
    import app.services.ob_service as module
    monkeypatch.setattr(module, 'SAP_MODE', 'MANUAL')
    saved = client.post('/api/ob/batch-save', json={'items':[{'material_code':'00001','required_qty':1}]}).json()
    assert client.post('/api/ob/fill', json={'item_ids':saved['ids']}).status_code == 400
    result = client.post('/api/ob/delete-draft-3001?ob_number=00000123').json()
    assert result['verified'] is False and result['code'] == 'LIVE_WRITE_NOT_IMPLEMENTED'


def test_source_reader_preserves_leading_zeroes_and_rejects_bad_source():
    mb = b'OB,Material,Batch,Quantity,Unit,Movement Type,Material Document\n0000123,00001,00002,10,EA,641,0000009\n'
    row = parse_source('mb51', mb, 'mb51.csv')[0]
    assert (row['ob_number'],row['material_code'],row['batch']) == ('0000123','00001','00002')
    with pytest.raises(ValueError):
        parse_source('mb51', b'OB,Quantity\n123,10', 'bad.csv')


def test_actual_file_import_cf_reconcile_draft_upgrade_and_download(client):
    pick = client.post('/api/pick-orders', json={'shop':'BODY','model':'VF7','items':[{'material_code':'00001','batch':'00002','req_qty':10,'ob_number':'0000123'}]}).json()
    pick_id = pick['id']
    details = client.get(f'/api/pick-orders/{pick_id}').json()
    line = str(details['items'][0]['id'])
    assert client.post(f'/api/pick-orders/{pick_id}/verify-cf', json={'verified_by':'Pilot','actual_quantities':{line:0}}).status_code == 400
    assert client.post(f'/api/pick-orders/{pick_id}/verify-cf', json={'verified_by':'Pilot','actual_quantities':{line:10}}).status_code == 200
    draft = client.post('/api/bbbg/create', json={'pick_order_id':pick_id,'sender_name':'Pilot','receiver_name':'Receiver'}).json()
    assert draft['is_draft'] is True and draft['gos_bbbg_attached'] == 0
    assert client.get(draft['download_url']).status_code == 200
    captured = datetime.datetime.now().isoformat()
    sources = {
        'mb51':b'OB,Material,Batch,Quantity,Unit,Movement Type,Material Document\n0000123,00001,00002,10,EA,641,0000009\n',
        'vl06i':b'OB,Material,Batch,Quantity,Unit,IB,Plant,SLoc,GR Status\n0000123,00001,00002,10,EA,0000033,3001,1003,C\n',
    }
    for kind, body in sources.items():
        result = client.post(f'/api/pick-orders/{pick_id}/sources/{kind}', params={'filename':kind+'.csv','captured_at':captured,'complete':'true'}, content=body)
        assert result.status_code == 200, result.text
    result = client.post(f'/api/pick-orders/{pick_id}/reconcile-uploaded')
    assert result.status_code == 200, result.text
    final = client.post('/api/bbbg/create', json={'pick_order_id':pick_id,'sender_name':'Pilot','receiver_name':'Receiver'})
    assert final.status_code == 200, final.text
    final = final.json()
    assert final['id'] == draft['id'] and final['bbbg_no'] == draft['bbbg_no']
    assert final['is_draft'] is False and final['reconciliation_status'] == 'IMPORTED_RECONCILED'
    assert final['gos_bbbg_attached'] == 0 and final['gos_vat_attached'] == 0
    assert client.get(final['download_url']).status_code == 200
    assert client.get(f'/api/pick-orders/{pick_id}').json()['status'] == 'Đã CF'
    assert client.post(f'/api/pick-orders/{pick_id}/sources/mb51', params={'filename':'mb51.csv','captured_at':captured,'complete':'true'}, content=sources['mb51']).status_code == 409


def test_legacy_confirm_accepts_explicit_actuals_but_not_implicit_cf(client):
    pick = client.post('/api/pick-orders', json={'shop':'BODY','model':'VF7','items':[{'material_code':'00022','batch':'0033','req_qty':4,'ob_number':'000099'}]}).json()
    line_id = client.get(f"/api/pick-orders/{pick['id']}").json()['items'][0]['id']
    endpoint = f"/api/pick/{pick['pick_order_code']}/confirm"
    assert client.post(endpoint, json={'confirmed_by':'Operator'}).status_code == 400
    assert client.post(endpoint, json={'confirmed_by':'Operator','actual_quantities':{str(line_id):4}}).status_code == 200


def test_ob_status_standardization_and_pick_blocking_invariants(client):
    # Criterion 1: Standardized 8 OB statuses
    for code, expected_code, expected_group in [
        ('UNVERIFIED', 'UNVERIFIED', 'Chưa kiểm tra'),
        ('INVALID', 'INVALID_CANCELLED', 'Cần xử lý'),
        ('CANCELLED', 'INVALID_CANCELLED', 'Cần xử lý'),
        ('LOCKED', 'LOCKED_SHORTAGE', 'Cần xử lý'),
        ('SHORTAGE', 'LOCKED_SHORTAGE', 'Cần xử lý'),
        ('READY_FOR_PICK', 'READY_FOR_PICK', 'Đủ điều kiện'),
        ('WAITING_CF', 'WAITING_CF', 'Chờ CF'),
        ('PARTIAL_CF', 'PARTIAL_CF', 'Chờ CF'),
        ('CF_COMPLETED', 'CF_COMPLETED', 'Đã hoàn thành'),
        ('READY_FOR_BBBG', 'READY_FOR_BBBG', 'Đã hoàn thành'),
    ]:
        norm = normalize_ob_status(code)
        assert norm['status_code'] == expected_code
        assert norm['status_group'] == expected_group

    # Criterion 2: Having an OB number string does NOT make an unverified row WAITING_CF,
    # and create_pick_order blocks unverified / shortage linked rows without changing their status.
    svc = StockService()
    svc.create_snapshot('stock', '3001', '1003', [{'material_code': 'MAT-READY', 'batch': 'B01', 'qty': 50}])

    saved = client.post('/api/ob/batch-save', json={
        'branch_tab': 'BODY_3001',
        'items': [
            {'item_uid': 'u-unverified', 'material_code': 'MAT-READY', 'batch': 'B01', 'req_qty': 20, 'ob_number': '1012640001'},
            {'item_uid': 'u-shortage', 'material_code': 'MAT-SHORT', 'batch': 'B01', 'req_qty': 100, 'ob_number': '1012640002'},
        ]
    }).json()
    id_unverified, id_shortage = saved['ids']

    # Verify u-unverified stays UNVERIFIED despite having ob_number
    items_before = {i['id']: i for i in client.get('/api/ob/items?branch_tab=BODY_3001').json()}
    assert items_before[id_unverified]['status_code'] == 'UNVERIFIED'

    # Try creating pick on unverified linked row -> MUST FAIL and NOT convert to WAITING_CF
    bad_pick1 = client.post('/api/pick-orders', json={
        'branch_tab': 'BODY_3001', 'shop': 'BODY', 'model': 'VF7',
        'items': [{'id': id_unverified, 'material_code': 'MAT-READY', 'batch': 'B01', 'req_qty': 20, 'ob_number': '1012640001'}]
    })
    assert bad_pick1.status_code == 400
    assert next(i for i in client.get('/api/ob/items?branch_tab=BODY_3001').json() if i['id'] == id_unverified)['status_code'] == 'UNVERIFIED'

    # Check stock on shortage row -> becomes LOCKED_SHORTAGE
    client.post('/api/ob/check', json={'branch_tab': 'BODY_3001', 'item_ids': [id_shortage]})
    # Try creating pick on shortage linked row -> MUST FAIL
    bad_pick2 = client.post('/api/pick-orders', json={
        'branch_tab': 'BODY_3001', 'shop': 'BODY', 'model': 'VF7',
        'items': [{'id': id_shortage, 'material_code': 'MAT-SHORT', 'batch': 'B01', 'req_qty': 100, 'ob_number': '1012640002'}]
    })
    assert bad_pick2.status_code == 400

    # Block unverified manual OB input
    bad_pick3 = client.post('/api/pick-orders', json={
        'branch_tab': 'BODY_3001', 'shop': 'BODY', 'model': 'VF7',
        'items': [{'material_code': 'MAT-MANUAL', 'batch': 'B01', 'req_qty': 10, 'ob_number': '1012640003', 'unverified': True}]
    })
    assert bad_pick3.status_code == 400

    # Empty MB51 must not speculate OB deletion without verification
    ob_svc = OBService()
    with pytest.raises(ValueError, match="MB51 rỗng"):
        ob_svc.replan_deleted_ob("MAT-READY", "1012640001", mb51_rows=[], verified_deleted=False)


def test_waiting_cf_partial_progress_filtering_history_and_bbbg_gate(client):
    # Create a pick order with 2 lines
    pick = client.post('/api/pick-orders', json={
        'branch_tab': 'BODY_3001',
        'shop': 'BODY',
        'model': 'VF8',
        'items': [
            {'material_code': '00MAT1', 'batch': '00B1', 'req_qty': 50, 'ob_number': '00OB1001'},
            {'material_code': '00MAT2', 'batch': '00B2', 'req_qty': 30, 'ob_number': '00OB1002'},
        ]
    }).json()
    po_id = pick['id']
    po_code = pick['pick_order_code']

    # Query Waiting CF list & filters
    waiting_po = client.get(f'/api/waiting-cf?pick_order_code={po_code}').json()
    assert len(waiting_po) == 2
    assert all(l['sap_cf_verified'] is False and 'Xác nhận thực tế thủ công' in l['cf_source_label'] for l in waiting_po)
    assert len(client.get('/api/waiting-cf?ob_number=00OB1001').json()) == 1
    assert len(client.get('/api/waiting-cf?material_code=00MAT2').json()) == 1

    line1_id = waiting_po[0]['id']
    line2_id = waiting_po[1]['id']

    # Partial CF without reason when missing > 0 must fail
    no_reason = client.post('/api/waiting-cf/save-line', json={
        'line_id': line1_id, 'actual_qty': 35, 'verified_by': 'Tester_K'
    })
    assert no_reason.status_code == 400

    # Save partial CF on line 1 with reason & evidence
    part_res = client.post('/api/waiting-cf/save-line', json={
        'line_id': line1_id,
        'actual_qty': 35,
        'verified_by': 'Tester_K',
        'reason': 'Đang lấy thêm 15 PCS từ dãy B',
        'evidence': 'Kiểm đếm pallet 1'
    })
    assert part_res.status_code == 200
    pdata = part_res.json()
    assert pdata['cf_status'] == 'PARTIAL_CF'
    assert pdata['missing_qty'] == 15
    assert pdata['order_status'] == 'PARTIAL_CF'
    assert pdata['bbbg_unlocked'] is False

    # BBBG MUST remain blocked when order is PARTIAL_CF
    bbbg_blocked = client.post('/api/bbbg/create', json={'pick_order_id': po_id, 'sender_name': 'Tester_K'})
    assert bbbg_blocked.status_code == 400

    # Complete line 2 (30/30), order still PARTIAL_CF because line 1 is 35/50
    l2_res = client.post('/api/waiting-cf/save-line', json={
        'line_id': line2_id, 'actual_qty': 30, 'verified_by': 'Tester_K', 'evidence': 'Đủ pallet 2'
    }).json()
    assert l2_res['cf_status'] == 'CF_COMPLETED'
    assert l2_res['order_status'] == 'PARTIAL_CF'
    assert l2_res['bbbg_unlocked'] is False

    # Complete line 1 (50/50) -> 100% CF completed -> leaves Waiting CF & unlocks BBBG!
    l1_full = client.post('/api/waiting-cf/save-line', json={
        'line_id': line1_id, 'actual_qty': 50, 'verified_by': 'Tester_K', 'evidence': 'Đã gom đủ 50 PCS'
    }).json()
    assert l1_full['cf_status'] == 'CF_COMPLETED'
    assert l1_full['order_status'] == 'Đã CF'
    assert l1_full['bbbg_unlocked'] is True
    assert l1_full['left_waiting_cf'] is True

    # Order has left the active Waiting CF list
    assert client.get(f'/api/waiting-cf?pick_order_code={po_code}').json() == []

    # History recorded with operator, time, and evidence
    hist = client.get(f'/api/waiting-cf/history?pick_order_id={po_id}').json()
    assert len(hist) >= 3
    assert any(h['verified_by'] == 'Tester_K' and h['status'] == 'PARTIAL_CF' for h in hist)

    # BBBG is now unlocked
    bbbg_ok = client.post('/api/bbbg/create', json={'pick_order_id': po_id, 'sender_name': 'Tester_K', 'receiver_name': 'Recv_A'})
    assert bbbg_ok.status_code == 200
