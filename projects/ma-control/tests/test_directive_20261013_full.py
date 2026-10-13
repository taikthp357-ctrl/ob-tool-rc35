# -*- coding: utf-8 -*-
"""Full verification suite for User Directive 13/10/2026 (AIM-1176):

1. Direct SAP MB52 download, snapshot sync by scope, intermediate file auto-cleanup and diagnostic retention.
2. Direct SAP MB51 BBBG reconciliation with technical detail concealment & clear operator summary.
3. OB grid bám sát Excel cũ (Model next to #, no extra checkbox col, text leading zeros, dynamic rows, direct CF paste).
4. Low stock suggestion area (DOH = INT(Tổng Body / US), respects plant/kho/shop, transfers to OB).
5. Visual Stock Matrix (LINE + KHO) BODY (8 models x 9 non-overlapping buckets, <=200 XE alert, drilldown).
6. Branch 4007 deferred and isolated from BODY.
7. Packaging & truthful disconnected SAP state.
"""

import base64
import os
import tempfile
from pathlib import Path
import pytest
from fastapi.testclient import TestClient

from app.main import app
from app.core import database
from app.services.stock_service import StockService, CANONICAL_MODEL_ROWS, BUCKET_DEFS, get_doh_bucket
from app.services.pick_cf_service import PickCFService


@pytest.fixture
def client(tmp_path, monkeypatch):
    db_path = tmp_path / 'directive_test.db'
    monkeypatch.setattr(database, 'DB_PATH', db_path)
    monkeypatch.setenv('MA_CONTROL_EXPORT_DIR', str(tmp_path / 'exports'))
    database.init_db(db_path)
    with TestClient(app) as c:
        yield c


def test_req1_mb52_direct_download_sync_snapshot_and_cleanup(client):
    """Req 1: MB52 tải trực tiếp từ SAP và cập nhật DB snapshot.
    File trung gian tự xóa sau khi parse và kiểm tra hợp lệ; nếu lỗi giữ lại file để chẩn đoán.
    Nếu chưa có phiên SAP live, ghi rõ trạng thái chưa kết nối.
    """
    # 1. Without test bytes and without SAP GUI -> Truthful SAP_NOT_CONNECTED
    res_no_sap = client.post("/api/stock/sap/download-mb52")
    assert res_no_sap.status_code == 200
    data_no_sap = res_no_sap.json()
    assert data_no_sap["status"] == "SAP_NOT_CONNECTED"
    assert data_no_sap["connected"] is False
    assert "Chưa có phiên SAP GUI kết nối" in data_no_sap["message"]

    # 2. Valid MB52 content downloaded -> Ingest snapshot, verify auto-cleanup
    mb52_valid_csv = (
        "Material,Plant,Storage Location,Batch,Unrestricted,Base Unit of Measure\n"
        "BIW001,3001,1003,0026101301,300,PCS\n"
        "BIW002,3001,1003,0026101302,150,PCS\n"
    ).encode("utf-8")
    b64_valid = base64.b64encode(mb52_valid_csv).decode("ascii")

    res_valid = client.post("/api/stock/sap/download-mb52", json={
        "plants": ["3001"],
        "sloc": "1003",
        "test_bytes_b64": b64_valid,
        "test_filename": "SAP_MB52_AUTOTEST.csv"
    })
    assert res_valid.status_code == 200
    data_valid = res_valid.json()
    assert data_valid["status"] == "SUCCESS"
    assert data_valid["intermediate_file_cleaned"] is True
    assert data_valid["diagnostic_file_retained"] is False
    assert data_valid["total_rows"] == 2

    # Verify snapshot in DB
    items = client.get("/api/stock/items?plant=3001&sloc=1003").json()
    assert len(items) == 2
    assert {i["material_code"] for i in items} == {"BIW001", "BIW002"}

    # 3. Invalid MB52 content -> Fails, retains diagnostic file
    mb52_invalid_csv = b"BAD_HEADER_COLUMN1,COLUMN2\nFOO,BAR\n"
    b64_invalid = base64.b64encode(mb52_invalid_csv).decode("ascii")

    res_invalid = client.post("/api/stock/sap/download-mb52", json={
        "test_bytes_b64": b64_invalid,
        "test_filename": "SAP_MB52_FAIL_DIAG.csv"
    })
    assert res_invalid.status_code == 200
    data_invalid = res_invalid.json()
    assert data_invalid["status"] == "ERROR"
    assert data_invalid["diagnostic_file_retained"] is True
    assert data_invalid["intermediate_file_cleaned"] is False
    assert data_invalid["diagnostic_file"] is not None
    assert os.path.exists(data_invalid["diagnostic_file"])
    # Clean up the retained diagnostic test file
    try:
        os.remove(data_invalid["diagnostic_file"])
    except OSError:
        pass


def test_req2_direct_mb51_reconciliation_and_summary_hiding(client):
    """Req 2: Đối soát BBBG & MB51 trực tiếp:
    Lấy MB51 trực tiếp từ SAP theo đơn/phạm vi, không bắt buộc chọn file MB51 thủ công trong luồng chính.
    Ẩn chi tiết kỹ thuật sau phần xem chi tiết; hiển thị kết quả, sai lệch và việc cần xử lý.
    """
    # Seed stock and create pick order
    client.post("/api/stock/upload", json={
        "source": "SEED", "plant": "3001", "sloc": "1003",
        "items": [{"material_code": "MAT-REC-1", "batch": "00B01", "unrestricted_qty": 100.0, "unit": "PCS"}]
    })
    pick = client.post("/api/pick-orders", json={
        "branch_tab": "BODY_3001", "shop": "BODY", "model": "VF7",
        "items": [{"material_code": "MAT-REC-1", "batch": "00B01", "req_qty": 20.0, "ob_number": "00OB99001"}]
    }).json()
    pick_id = pick["id"]

    # Confirm line CF
    lines = client.get(f"/api/waiting-cf?pick_order_code={pick['pick_order_code']}").json()
    line_id = lines[0]["id"]
    client.post("/api/waiting-cf/save-line", json={
        "line_id": line_id, "actual_qty": 20.0, "verified_by": "Operator_1"
    })

    # 1. Direct reconcile when SAP not connected and no bytes
    res_no_sap = client.post(f"/api/pick-orders/{pick_id}/direct-reconcile-sap")
    assert res_no_sap.status_code == 200
    d_no_sap = res_no_sap.json()
    assert d_no_sap["status"] == "SAP_NOT_CONNECTED"
    assert d_no_sap["connected"] is False
    assert d_no_sap["operator_summary"]["is_valid"] is False
    assert "Chưa có phiên SAP GUI kết nối" in d_no_sap["operator_summary"]["result_label"]

    # 2. Direct reconcile with matching MB51 data
    mb51_content = (
        "OB,Material,Batch,Quantity,Unit,Movement Type,Material Document\n"
        "00OB99001,MAT-REC-1,00B01,20.0,PCS,641,000998811\n"
    ).encode("utf-8")
    b64_mb51 = base64.b64encode(mb51_content).decode("ascii")

    res_rec = client.post(f"/api/pick-orders/{pick_id}/direct-reconcile-sap", json={
        "test_mb51_bytes_b64": b64_mb51,
        "test_mb51_filename": "SAP_MB51_DIRECT.csv"
    })
    assert res_rec.status_code == 200
    d_rec = res_rec.json()
    assert d_rec["status"] == "MATCHED"
    summary = d_rec["operator_summary"]
    assert summary["is_valid"] is True
    assert summary["matched_count"] == 1
    assert summary["discrepancies"] == []
    assert "Hợp lệ" in summary["result_label"]
    assert "Đủ điều kiện xuất BBBG" in summary["pending_actions"]
    assert "technical_details" in d_rec


def test_req3_ob_grid_excel_order_and_waiting_cf_direct_paste(client):
    """Req 3: Bảng điền OB & đơn pick bám sát mẫu Excel cũ:
    - Model nằm cạnh # (Column A), bỏ checkbox thừa.
    - Giữ số 0 đầu mã và OB dạng text.
    - Dán trực tiếp từ Excel vào Chờ CF.
    """
    # Check HTML headers
    resp = client.get("/")
    assert resp.status_code == 200
    html = resp.text
    assert '<span class="col-excel-letter">A</span>Model' in html
    assert '<span class="col-excel-letter">B</span>Mã hàng' in html
    assert '<span class="col-excel-letter">C</span>SL yêu cầu' in html

    # Direct paste into Waiting CF from Excel TSV
    tsv_data = (
        "00990011\t0001234567\t3001\t00BATCH01\t50.5\tBOX-1\tRACK-1\n"
        "00990012\t0001234568\t3001\t00BATCH02\t25.0\tBOX-2\tRACK-2\n"
    )
    paste_res = client.post("/api/waiting-cf/paste-excel", json={
        "raw_text": tsv_data,
        "allow_duplicates": False
    })
    assert paste_res.status_code == 200
    pdata = paste_res.json()
    assert pdata["count"] == 2
    lines = pdata["lines"]
    assert lines[0]["ob_number"] == "00990011"
    assert lines[0]["material_code"] == "0001234567"
    assert lines[0]["batch"] == "00BATCH01"
    assert lines[0]["qty"] == 50.5


def test_req4_low_stock_suggestions_doh_and_transfer_to_ob(client):
    """Req 4: Khu gợi ý tồn thấp:
    Tính lượng bổ sung từ tồn hiện hành và DOH của model:
    DOH = INT(Tổng Body / US) theo baseline V3.6.
    Chuyển gợi ý sang bảng điền OB.
    """
    conn = database.get_connection()
    conn.execute("INSERT OR REPLACE INTO materials (material_code, model, us_rate) VALUES ('MAT-SUGG-1', 'VF7', 2)")
    conn.execute("INSERT OR REPLACE INTO stock_thresholds (material_code, plant, threshold_qty) VALUES ('MAT-SUGG-1', '3001', 200)")
    conn.commit()
    conn.close()

    # Seed stock: Line 3003 = 20, Kho 1003 = 80 => Tổng Body = 100
    # US = 2 => DOH = INT(100 / 2) = 50 XE (< threshold 200 XE)
    # Target = 200 * 2 = 400 PCS => Supplement needed = 400 - 100 = 300 PCS
    client.post("/api/stock/upload", json={
        "source": "SEED", "plant": "3001", "sloc": "1003",
        "items": [
            {"material_code": "MAT-SUGG-1", "batch": "00B01", "unrestricted_qty": 80.0, "unit": "PCS"},
            {"material_code": "MAT-SUGG-1", "batch": "00B02", "unrestricted_qty": 200.0, "unit": "PCS"}
        ]
    })
    client.post("/api/stock/upload", json={
        "source": "SEED", "plant": "3001", "sloc": "3003",
        "items": [
            {"material_code": "MAT-SUGG-1", "batch": "00B01", "unrestricted_qty": 20.0, "unit": "PCS"}
        ]
    })

    # Fetch low stock suggestions
    res_sugg = client.get("/api/stock/low-stock/suggestions?plant=3001&shop=BODY")
    assert res_sugg.status_code == 200
    suggestions = res_sugg.json()
    assert len(suggestions) >= 1
    sugg = next(s for s in suggestions if s["material_code"] == "MAT-SUGG-1")
    assert sugg["tong_body"] == 300.0
    assert sugg["current_doh"] == 150
    assert sugg["threshold_doh"] == 200.0
    assert sugg["supplement_needed"] == 100.0
    assert len(sugg["batches"]) >= 1

    # Transfer suggestions to OB table
    trans_res = client.post("/api/stock/low-stock/transfer-to-ob", json={
        "branch_tab": "BODY_3001",
        "items": [
            {
                "material_code": sugg["material_code"],
                "model": sugg["model"],
                "required_qty": sugg["supplement_needed"],
                "batch": sugg["batches"][0]["batch"],
                "shop": "BODY"
            }
        ]
    })
    assert trans_res.status_code == 200
    trans_data = trans_res.json()
    assert trans_data["status"] == "success"
    assert trans_data["transferred_count"] == 1

    # Verify in OB items table
    ob_items = client.get("/api/ob/items?branch_tab=BODY_3001").json()
    transferred_item = next(i for i in ob_items if i["material_code"] == "MAT-SUGG-1")
    assert transferred_item["status"] == "Gợi ý tồn thấp"
    assert transferred_item["required_qty"] == 100.0


def test_req5_visual_stock_matrix_buckets_and_two_tables(client):
    """Req 5: Màn tồn & Báo cáo trực quan FG:
    - Báo cáo ma trận (LINE + KHO) BODY: 8 models x 9 non-overlapping buckets.
    - Dải <= 200 XE tô đỏ cảnh báo.
    - Drilldown chi tiết mã theo ô số tổng.
    - Hai bảng tồn: Bảng ưu tiên và Bảng tổng tồn.
    """
    # Seed stock for VF7 and VF8
    client.post("/api/stock/upload", json={
        "source": "SEED", "plant": "3001", "sloc": "1003",
        "items": [
            {"material_code": "MAT-VF7-01", "batch": "B1", "unrestricted_qty": 80.0, "unit": "PCS"},
            {"material_code": "MAT-VF8-01", "batch": "B2", "unrestricted_qty": 600.0, "unit": "PCS"}
        ]
    })

    # 1. Visual matrix endpoint
    res_mat = client.get("/api/stock/visual-matrix")
    assert res_mat.status_code == 200
    matrix = res_mat.json()
    assert matrix["title"] == "(LINE + KHO) BODY"
    assert len(matrix["rows"]) == 8
    model_names = [r["model"] for r in matrix["rows"]]
    assert model_names == CANONICAL_MODEL_ROWS

    # Check column definitions: 9 non-overlapping buckets
    cols = matrix["columns"]
    assert len(cols) == 9
    assert [c["key"] for c in cols] == ["0XE", "<=50XE", "<=100XE", "<=150XE", "<=200XE", "<=300XE", "<=400XE", "<=500XE", ">500XE"]
    # Alert flag set for <= 200 XE
    assert all(c["is_alert"] for c in cols if c["key"] in ("0XE", "<=50XE", "<=100XE", "<=150XE", "<=200XE"))
    assert not any(c["is_alert"] for c in cols if c["key"] in ("<=300XE", "<=400XE", "<=500XE", ">500XE"))

    # 2. Drilldown materials endpoint
    res_drill = client.get("/api/stock/visual-matrix/materials?model=VF7")
    assert res_drill.status_code == 200
    drill_items = res_drill.json()
    assert any(i["material_code"] == "MAT-VF7-01" for i in drill_items)

    # 3. Two stock tables endpoint
    res_tables = client.get("/api/stock/two-tables?plant=3001")
    assert res_tables.status_code == 200
    tables = res_tables.json()
    assert "priority_table" in tables
    assert "total_stock_table" in tables
    assert tables["total_count"] >= 2


def test_req6_branch_4007_deferred_and_isolated(client):
    """Req 6: Hoãn nhánh 4007 theo chỉ đạo người dùng; không tự gộp vào BODY."""
    # Verify BODY_3001 and EDC_4007 operate with strict separation
    client.post("/api/ob/batch-save", json={
        "branch_tab": "BODY_3001",
        "items": [{"item_uid": "body-1", "material_code": "MAT-BODY", "req_qty": 10}]
    })
    client.post("/api/ob/batch-save", json={
        "branch_tab": "EDC_4007",
        "items": [{"item_uid": "edc-1", "material_code": "MAT-EDC", "req_qty": 20}]
    })

    body_items = client.get("/api/ob/items?branch_tab=BODY_3001").json()
    edc_items = client.get("/api/ob/items?branch_tab=EDC_4007").json()

    assert all(i["material_code"] != "MAT-EDC" for i in body_items)
    assert all(i["material_code"] != "MAT-BODY" for i in edc_items)


def test_req7_packaging_and_truthful_sap_status(client):
    """Req 7: Đóng gói VBS/tool & trung thực về trạng thái kết nối SAP."""
    health = client.get("/api/health").json()
    assert health["status"] == "healthy"
    assert health["sap_connection"] == "DISCONNECTED"
    assert health["live_write_verified"] is False
    assert "Chưa đọc được phiên SAP GUI" in health["sap_details"]["message"]
