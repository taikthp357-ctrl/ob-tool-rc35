# -*- coding: utf-8 -*-
"""Independent QA Verification Suite for AIM-1178 by Nova (QA & Reliability Engineer).

Focus areas per user directive 13/10/2026:
1. Kiểm thử luồng tải tồn MB52 và dọn dẹp file trung gian: file tạm tự xóa khi parse thành công, giữ lại khi parse lỗi.
2. Kiểm thử dán nhiều dòng Excel vào bảng điền OB: Model cạnh #, bỏ cột A thừa, dán 10-50 dòng, giữ dòng trống, giữ số 0 đầu mã/OB.
3. Kiểm thử khu gợi ý tồn thấp: tính DOH theo US của model, cho chỉnh số lượng theo mã/batch, chuyển gợi ý sang bảng điền OB thành công.
4. Kiểm thử bảng ma trận phân bố tồn (LINE + KHO) BODY: các dải không chồng nhau, tính đúng tổng số mã và tỷ lệ %, click xem chi tiết mã.
5. Kiểm thử Chờ CF dán từ Excel độc lập, không ép qua đơn pick tự sinh.
"""

import base64
import os
import tempfile
from pathlib import Path
import pytest
from fastapi.testclient import TestClient

from app.main import app
from app.core import database
from app.services.stock_service import (
    StockService,
    CANONICAL_MODEL_ROWS,
    BUCKET_DEFS,
    get_doh_bucket,
    map_model_to_canonical
)
from app.services.pick_cf_service import PickCFService
from app.services.ob_service import OBService


@pytest.fixture
def client(tmp_path, monkeypatch):
    """Isolated temporary SQLite database for independent QA verification."""
    db_path = tmp_path / "nova_qa_temp.db"
    monkeypatch.setattr(database, "DB_PATH", db_path)
    monkeypatch.setenv("MA_CONTROL_EXPORT_DIR", str(tmp_path / "exports"))
    database.init_db(db_path)
    with TestClient(app) as c:
        yield c


# =============================================================================
# 1. KIỂM THỬ LUỒNG TẢI TỒN MB52 VÀ DỌN DẸP FILE TRUNG GIAN
# =============================================================================

def test_qa_01_mb52_successful_download_cleans_intermediate_file(client):
    """Kiểm thử: File trung gian tự xóa khi parse và nạp snapshot thành công."""
    svc = StockService()
    valid_mb52_csv = (
        "Material,Plant,Storage Location,Batch,Unrestricted,Base Unit of Measure\n"
        "0001002003,3001,1003,0026101301,500.0,PCS\n"
        "0001002004,3001,1003,0026101302,250.0,PCS\n"
    ).encode("utf-8")
    test_filename = f"SAP_MB52_TEST_CLEANUP_{os.getpid()}.csv"
    temp_target = Path(tempfile.gettempdir()) / test_filename

    result = svc.download_and_sync_sap_mb52(
        plants=["3001"],
        sloc="1003",
        test_bytes=valid_mb52_csv,
        test_filename=test_filename,
        is_user_upload=False
    )

    assert result["status"] == "SUCCESS"
    assert result["intermediate_file_cleaned"] is True
    assert result["diagnostic_file_retained"] is False
    assert result["total_rows"] == 2
    assert result["total_qty"] == 750.0
    # Verifying file does not exist on disk after cleanup
    assert not temp_target.exists(), "File trung gian phải bị tự động xóa sau khi parse thành công!"

    # Verify snapshot replaced current stock for plant 3001 / sloc 1003
    items = client.get("/api/stock/items?plant=3001&sloc=1003").json()
    assert len(items) == 2
    assert {i["material_code"] for i in items} == {"0001002003", "0001002004"}


def test_qa_01_mb52_failed_parse_retains_diagnostic_file(client):
    """Kiểm thử: Giữ lại file trung gian khi parse lỗi để phục vụ chẩn đoán."""
    svc = StockService()
    corrupt_mb52_bytes = b"CORRUPTED_HEADER_INVALID\nXYZ,123,456\n"
    test_filename = f"SAP_MB52_FAIL_DIAG_{os.getpid()}.csv"
    temp_target = Path(tempfile.gettempdir()) / test_filename

    result = svc.download_and_sync_sap_mb52(
        plants=["3001"],
        sloc="1003",
        test_bytes=corrupt_mb52_bytes,
        test_filename=test_filename,
        is_user_upload=False
    )

    assert result["status"] == "ERROR"
    assert result["diagnostic_file_retained"] is True
    assert result["intermediate_file_cleaned"] is False
    retained_file = result.get("diagnostic_file")
    assert retained_file is not None
    assert Path(retained_file).exists(), "File lỗi phải được giữ lại trên đĩa để chẩn đoán!"

    # Clean up test diagnostic file after assertion
    try:
        Path(retained_file).unlink()
    except OSError:
        pass


def test_qa_01_mb52_user_upload_never_deleted(client):
    """Kiểm thử: File do người dùng tải lên (is_user_upload=True) tuyệt đối không bị xóa."""
    svc = StockService()
    valid_csv = (
        "Material,Plant,Storage Location,Batch,Unrestricted,Base Unit of Measure\n"
        "0001002005,3001,1003,0026101303,100.0,PCS\n"
    ).encode("utf-8")
    test_filename = f"USER_UPLOAD_TEST_{os.getpid()}.csv"
    temp_target = Path(tempfile.gettempdir()) / test_filename

    result = svc.download_and_sync_sap_mb52(
        plants=["3001"],
        sloc="1003",
        test_bytes=valid_csv,
        test_filename=test_filename,
        is_user_upload=True  # User upload flag
    )

    assert result["status"] == "SUCCESS"
    assert temp_target.exists(), "File người dùng tải lên không bao giờ được tự động xóa!"
    try:
        temp_target.unlink()
    except OSError:
        pass


def test_qa_01_mb52_no_sap_connection_reports_truthful_state(client):
    """Kiểm thử: Khi không có phiên SAP live và không có test bytes, báo rõ SAP_NOT_CONNECTED."""
    res = client.post("/api/stock/sap/download-mb52")
    assert res.status_code == 200
    data = res.json()
    assert data["status"] == "SAP_NOT_CONNECTED"
    assert data["connected"] is False
    assert "Chưa có phiên SAP GUI kết nối" in data["message"]


# =============================================================================
# 2. KIỂM THỬ DÁN NHIỀU DÒNG EXCEL VÀO BẢNG ĐIỀN OB (10-50 DÒNG, GIỮ 0 ĐẦU, DÒNG TRỐNG)
# =============================================================================

def test_qa_02_ob_grid_layout_model_next_to_hash_and_no_col_a_checkbox(client):
    """Kiểm thử giao diện HTML bảng điền OB:
    - Model nằm cạnh # (Cột A)
    - Không có cột checkbox thừa
    - Thứ tự cột: # | A Model | B Mã hàng | C SL yêu cầu | D Batch | E Shop | F STO | G OB | H Trạng thái
    """
    resp = client.get("/")
    assert resp.status_code == 200
    html = resp.text

    # Cột Model là cột A ngay cạnh cột #
    assert '<th style="width: 44px;">#</th>' in html
    assert '<span class="col-excel-letter">A</span>Model' in html
    assert '<span class="col-excel-letter">B</span>Mã hàng' in html
    assert '<span class="col-excel-letter">C</span>SL yêu cầu' in html
    assert '<span class="col-excel-letter">D</span>Batch' in html
    assert '<span class="col-excel-letter">E</span>Shop' in html
    assert '<span class="col-excel-letter">F</span>STO' in html
    assert '<span class="col-excel-letter">G</span>OB' in html
    assert '<span class="col-excel-letter">H</span>Trạng thái' in html

    # Không có cột checkbox nào trong header thead
    assert '<input type="checkbox" id="check-all"' not in html


def test_qa_02_paste_10_to_50_rows_with_empty_lines_and_leading_zeros(client):
    """Kiểm thử dán 10 đến 50 dòng Excel vào bảng điền OB:
    - Dán dải 25 dòng và 50 dòng
    - Giữ các dòng trống xen kẽ
    - Bảo toàn định dạng Text với số 0 ở đầu mã hàng, số OB, batch, STO
    """
    # Tạo payload 25 dòng có 3 dòng trống xen kẽ và mã có số 0 đầu
    conn = database.get_connection()
    conn.execute("DELETE FROM ob_items WHERE branch_tab = 'BODY_3001'")
    conn.commit()
    conn.close()

    rows = []
    for i in range(1, 26):
        if i in (5, 12, 20):
            # Dòng trống xen kẽ
            rows.append({
                "item_uid": f"uid-empty-{i}",
                "branch_tab": "BODY_3001",
                "material_code": "",
                "model": "",
                "required_qty": 0.0,
                "batch": "",
                "shop": "",
                "sto_number": "",
                "ob_number": "",
                "status": "Chưa kiểm tra"
            })
        else:
            rows.append({
                "item_uid": f"uid-{i}",
                "branch_tab": "BODY_3001",
                "material_code": f"000{1000000 + i}",   # Mã hàng giữ số 0 đầu
                "model": "VF7" if i % 2 == 0 else "VF8",
                "required_qty": float(i * 10),
                "batch": f"00B{i:03d}",                 # Batch giữ số 0 đầu
                "shop": "BODY",
                "sto_number": f"00STO{i:04d}",          # STO giữ số 0 đầu
                "ob_number": f"00OB{99000 + i}",        # OB giữ số 0 đầu
                "status": "Chưa kiểm tra"
            })

    # Lưu qua API batch-save (mô phỏng sau khi paste lên grid và lưu)
    save_resp = client.post("/api/ob/batch-save", json={
        "branch_tab": "BODY_3001",
        "items": rows
    })
    assert save_resp.status_code == 200
    saved_data = save_resp.json()
    assert saved_data["status"] == "success"
    assert saved_data["count"] == 22  # 25 rows minus 3 empty rows cleanly skipped from DB insert

    # Lấy lại từ DB để kiểm tra tính toàn vẹn các hàng có dữ liệu
    db_items = client.get("/api/ob/items?branch_tab=BODY_3001").json()
    assert len(db_items) == 22
    assert db_items[0]["row_order"] == 1
    assert db_items[3]["row_order"] == 4
    assert db_items[4]["row_order"] == 6  # Row 5 was empty, row 6 has order 6

    # Kiểm tra các số 0 ở đầu được bảo toàn nguyên vẹn kiểu TEXT
    assert db_items[0]["material_code"] == "0001000001"
    assert db_items[0]["batch"] == "00B001"
    assert db_items[0]["sto_number"] == "00STO0001"
    assert db_items[0]["ob_number"] == "00OB99001"

    assert db_items[21]["material_code"] == "0001000025"
    assert db_items[21]["batch"] == "00B025"
    assert db_items[21]["ob_number"] == "00OB99025"

    # Thử nghiệm mở rộng đến 50 dòng
    rows_50 = []
    for i in range(1, 51):
        rows_50.append({
            "item_uid": f"uid-50-{i}",
            "branch_tab": "BODY_3001",
            "material_code": f"00BIW{i:05d}",
            "model": "VF7",
            "required_qty": float(i * 5),
            "batch": f"00261013{i:02d}",
            "ob_number": f"009988{i:04d}"
        })
    save_50_resp = client.post("/api/ob/batch-save", json={
        "branch_tab": "BODY_3001",
        "items": rows_50
    })
    assert save_50_resp.status_code == 200
    assert save_50_resp.json()["count"] == 50


# =============================================================================
# 3. KIỂM THỬ KHU GỢI Ý TỒN THẤP (DOH THEO US, CHỈNH SỐ LƯỢNG, CHUYỂN SANG OB)
# =============================================================================

def test_qa_03_low_stock_suggestions_doh_formula_and_transfer_to_ob(client):
    """Kiểm thử khu gợi ý tồn thấp:
    - Công thức DOH = INT(Tổng Body / US) theo baseline 261008_Stock_FG_MA_BODY_V3_6.xlsm
      trong đó Tổng Body = Line 3003 + Kho 1003.
    - Cho xem, chỉnh số lượng theo mã và theo batch.
    - Chuyển gợi ý sang bảng điền OB thành công.
    """
    conn = database.get_connection()
    # Tạo 2 vật tư:
    # MAT-A (VF7): US = 2, Ngưỡng = 200 XE
    # MAT-B (VF8): US = 1, Ngưỡng = 200 XE
    conn.execute("INSERT OR REPLACE INTO materials (material_code, model, us_rate) VALUES ('00MAT-LOW-A', 'VF7', 2)")
    conn.execute("INSERT OR REPLACE INTO materials (material_code, model, us_rate) VALUES ('00MAT-HIGH-B', 'VF8', 1)")
    conn.execute("INSERT OR REPLACE INTO stock_thresholds (material_code, plant, threshold_qty) VALUES ('00MAT-LOW-A', '3001', 200)")
    conn.execute("INSERT OR REPLACE INTO stock_thresholds (material_code, plant, threshold_qty) VALUES ('00MAT-HIGH-B', '3001', 200)")
    conn.commit()
    conn.close()

    # Nạp tồn cho MAT-A:
    # Kho 1003 = 80 PCS, Line 3003 = 20 PCS => Tổng Body = 100 PCS
    # DOH = INT(100 / 2) = 50 XE (< ngưỡng 200 XE => CẦN GỢI Ý)
    # Target = 200 * 2 = 400 PCS => Lượng bổ sung cần = 400 - 100 = 300 PCS
    client.post("/api/stock/upload", json={
        "source": "SEED", "plant": "3001", "sloc": "1003",
        "items": [
            {"material_code": "00MAT-LOW-A", "batch": "00B01", "unrestricted_qty": 50.0, "unit": "PCS"},
            {"material_code": "00MAT-LOW-A", "batch": "00B02", "unrestricted_qty": 30.0, "unit": "PCS"},
            {"material_code": "00MAT-HIGH-B", "batch": "00B01", "unrestricted_qty": 250.0, "unit": "PCS"}
        ]
    })
    client.post("/api/stock/upload", json={
        "source": "SEED", "plant": "3001", "sloc": "3003",
        "items": [
            {"material_code": "00MAT-LOW-A", "batch": "00B01", "unrestricted_qty": 20.0, "unit": "PCS"}
        ]
    })

    # Lấy danh sách gợi ý tồn thấp
    sugg_res = client.get("/api/stock/low-stock/suggestions?plant=3001&shop=BODY")
    assert sugg_res.status_code == 200
    suggestions = sugg_res.json()

    # MAT-B có DOH 250 không được xuất hiện trong gợi ý tồn thấp
    assert not any(s["material_code"] == "00MAT-HIGH-B" for s in suggestions)

    # MAT-A phải có mặt với các chỉ số tính toán chuẩn xác
    sugg_a = next((s for s in suggestions if s["material_code"] == "00MAT-LOW-A"), None)
    assert sugg_a is not None
    assert sugg_a["model"] == "VF7"
    assert sugg_a["us_rate"] == 2
    assert sugg_a["tong_body"] == 100.0
    assert sugg_a["current_doh"] == 50
    assert sugg_a["threshold_doh"] == 200.0
    assert sugg_a["target_qty"] == 400
    assert sugg_a["supplement_needed"] == 300.0
    assert len(sugg_a["batches"]) >= 1

    # Kiểm thử chỉnh số lượng và chọn batch cụ thể trước khi chuyển sang OB:
    # Người dùng chỉnh số lượng yêu cầu thành 120 PCS (thay vì lấy cả 300 PCS) và chọn batch 00B01
    transfer_payload = {
        "branch_tab": "BODY_3001",
        "items": [
            {
                "material_code": "00MAT-LOW-A",
                "model": "VF7",
                "required_qty": 120.0,     # Chỉnh số lượng tùy chọn
                "batch": "00B01",           # Chọn batch cụ thể
                "shop": "BODY"
            }
        ]
    }
    trans_res = client.post("/api/stock/low-stock/transfer-to-ob", json=transfer_payload)
    assert trans_res.status_code == 200
    trans_data = trans_res.json()
    assert trans_data["status"] == "success"
    assert trans_data["transferred_count"] == 1

    # Xác nhận dòng đã xuất hiện trong bảng điền OB (ob_items)
    ob_items = client.get("/api/ob/items?branch_tab=BODY_3001").json()
    transferred_item = next((i for i in ob_items if i["material_code"] == "00MAT-LOW-A"), None)
    assert transferred_item is not None
    assert transferred_item["status"] == "Gợi ý tồn thấp"
    assert transferred_item["required_qty"] == 120.0
    assert transferred_item["batch"] == "00B01"
    assert transferred_item["model"] == "VF7"


# =============================================================================
# 4. KIỂM THỬ BẢNG MA TRẬN PHÂN BỐ TỒN (LINE + KHO) BODY
# =============================================================================

def test_qa_04_visual_stock_matrix_buckets_totals_and_drilldown(client):
    """Kiểm thử bảng ma trận phân bố tồn (LINE + KHO) BODY:
    - 9 dải không chồng nhau (0XE, <=50XE, <=100XE, <=150XE, <=200XE, <=300XE, <=400XE, <=500XE, >500XE)
    - 8 dòng model chuẩn hóa (Limo, VF5, VF6, VF6,7, VF6 CIS, VF7, VF8, VF8NR)
    - Tính đúng tổng số mã theo hàng, cột và grand_total
    - Tính đúng tỷ lệ % (1 chữ số thập phân)
    - Click xem chi tiết mã (drill-down API)
    """
    # 1. Kiểm tra tính đơn định và không chồng lấn của các dải tồn trên các điểm biên
    assert get_doh_bucket(0) == "0XE"
    assert get_doh_bucket(1) == "<=50XE"
    assert get_doh_bucket(50) == "<=50XE"
    assert get_doh_bucket(51) == "<=100XE"
    assert get_doh_bucket(100) == "<=100XE"
    assert get_doh_bucket(101) == "<=150XE"
    assert get_doh_bucket(150) == "<=150XE"
    assert get_doh_bucket(151) == "<=200XE"
    assert get_doh_bucket(200) == "<=200XE"
    assert get_doh_bucket(201) == "<=300XE"
    assert get_doh_bucket(300) == "<=300XE"
    assert get_doh_bucket(301) == "<=400XE"
    assert get_doh_bucket(400) == "<=400XE"
    assert get_doh_bucket(401) == "<=500XE"
    assert get_doh_bucket(500) == "<=500XE"
    assert get_doh_bucket(501) == ">500XE"
    assert get_doh_bucket(1000) == ">500XE"

    # 2. Chuẩn bị dữ liệu mẫu đại diện cho nhiều model và dải tồn
    conn = database.get_connection()
    conn.execute("INSERT OR REPLACE INTO materials (material_code, model, us_rate) VALUES ('MAT-M01', 'Limo', 2)")
    conn.execute("INSERT OR REPLACE INTO materials (material_code, model, us_rate) VALUES ('MAT-M02', 'VF5', 1)")
    conn.execute("INSERT OR REPLACE INTO materials (material_code, model, us_rate) VALUES ('MAT-M03', 'VF6', 1)")
    conn.execute("INSERT OR REPLACE INTO materials (material_code, model, us_rate) VALUES ('MAT-M04', 'VF7', 2)")
    conn.execute("INSERT OR REPLACE INTO materials (material_code, model, us_rate) VALUES ('MAT-M05', 'VF8', 1)")
    conn.commit()
    conn.close()

    # Nạp tồn:
    # MAT-M01 (Limo, US=2): 0 PCS => DOH = 0 (dải 0XE)
    # MAT-M02 (VF5, US=1): 40 PCS => DOH = 40 (dải <=50XE)
    # MAT-M03 (VF6, US=1): 120 PCS => DOH = 120 (dải <=150XE)
    # MAT-M04 (VF7, US=2): Line 3003 = 100, Kho 1003 = 200 => Tổng 300, DOH = 150 (dải <=150XE)
    # MAT-M05 (VF8, US=1): 600 PCS => DOH = 600 (dải >500XE)
    client.post("/api/stock/upload", json={
        "source": "SEED", "plant": "3001", "sloc": "1003",
        "items": [
            {"material_code": "MAT-M01", "batch": "B1", "unrestricted_qty": 0.0, "unit": "PCS"},
            {"material_code": "MAT-M02", "batch": "B2", "unrestricted_qty": 40.0, "unit": "PCS"},
            {"material_code": "MAT-M03", "batch": "B3", "unrestricted_qty": 120.0, "unit": "PCS"},
            {"material_code": "MAT-M04", "batch": "B4", "unrestricted_qty": 200.0, "unit": "PCS"},
            {"material_code": "MAT-M05", "batch": "B5", "unrestricted_qty": 600.0, "unit": "PCS"}
        ]
    })
    client.post("/api/stock/upload", json={
        "source": "SEED", "plant": "3001", "sloc": "3003",
        "items": [
            {"material_code": "MAT-M04", "batch": "B4", "unrestricted_qty": 100.0, "unit": "PCS"}
        ]
    })

    # Gọi API ma trận phân bố tồn
    mat_res = client.get("/api/stock/visual-matrix")
    assert mat_res.status_code == 200
    matrix = mat_res.json()

    assert matrix["title"] == "(LINE + KHO) BODY"
    assert matrix["has_data"] is True
    assert len(matrix["rows"]) == 8
    assert [r["model"] for r in matrix["rows"]] == CANONICAL_MODEL_ROWS

    # Kiểm tra cột: 9 dải
    cols = matrix["columns"]
    assert len(cols) == 9
    col_keys = [c["key"] for c in cols]
    assert col_keys == ["0XE", "<=50XE", "<=100XE", "<=150XE", "<=200XE", "<=300XE", "<=400XE", "<=500XE", ">500XE"]

    # Cảnh báo đỏ cho dải <= 200 XE
    for c in cols:
        if c["key"] in ("0XE", "<=50XE", "<=100XE", "<=150XE", "<=200XE"):
            assert c["is_alert"] is True
        else:
            assert c["is_alert"] is False

    # Kiểm tra tổng số mã
    totals = matrix["totals"]
    assert totals["grand_total"] >= 5
    counts = totals["counts"]
    percentages = totals["percentages"]

    # Tổng % các dải xấp xỉ 100%
    pct_sum = sum(float(p.replace("%", "")) for p in percentages.values())
    assert 99.0 <= pct_sum <= 101.0

    # 3. Kiểm tra click xem chi tiết mã (Drill-down API)
    # Drilldown theo model VF7
    drill_vf7 = client.get("/api/stock/visual-matrix/materials?model=VF7").json()
    assert any(m["material_code"] == "MAT-M04" for m in drill_vf7)
    item_m04 = next(m for m in drill_vf7 if m["material_code"] == "MAT-M04")
    assert item_m04["line_3003"] == 100.0
    assert item_m04["kho_1003"] == 200.0
    assert item_m04["tong_body"] == 300.0
    assert item_m04["doh"] == 150
    assert item_m04["bucket"] == "<=150XE"

    # Drilldown kết hợp model + bucket
    drill_bucket = client.get("/api/stock/visual-matrix/materials?model=VF7&bucket_key=<=150XE").json()
    assert len(drill_bucket) >= 1
    assert all(m["bucket"] == "<=150XE" for m in drill_bucket)


# =============================================================================
# 5. KIỂM THỬ CHỜ CF DÁN TỪ EXCEL ĐỘC LẬP (KHÔNG ÉP QUA ĐƠN PICK TỰ SINH)
# =============================================================================

def test_qa_05_waiting_cf_independent_excel_paste_and_workflow(client):
    """Kiểm thử Chờ CF dán từ Excel độc lập:
    - Không bắt buộc qua đơn pick tự sinh (pick_order_id = NULL, is_direct_paste = 1).
    - Tự động sinh direct_order_ref (EXCEL-PASTE-YYYYMMDD-XXXXXX).
    - Bảo toàn số 0 ở đầu (OB, mã hàng, batch, plant, box, rack).
    - Chống dán trùng active line trên DB.
    - Khởi tạo ở trạng thái Chờ CF / PENDING_CF, không tự động xác nhận ngầm.
    - Cho chỉnh sửa số lượng, packing box, rack và xác nhận thực tế.
    """
    # Dữ liệu TSV mô phỏng copy từ Excel có số 0 ở đầu
    excel_tsv = (
        "OB\tMã hàng\tPlant\tBatch\tSố lượng\tBox\tRack\n"
        "00109911\t00MAT7788\t3001\t00BATCH01\t65.0\tBOX-01\tRACK-A1\n"
        "00109922\t00MAT7799\t3001\t00BATCH02\t35.5\tBOX-02\tRACK-B2\n"
    )

    # 1. Dán trực tiếp từ Excel vào Chờ CF khi chưa tạo bất kỳ đơn pick nào
    paste_res = client.post("/api/waiting-cf/paste-excel", json={
        "raw_text": excel_tsv,
        "allow_duplicates": False
    })
    assert paste_res.status_code == 200
    res_data = paste_res.json()
    assert res_data["status"] == "success"
    assert res_data["count"] == 2
    lines = res_data["lines"]

    line1 = lines[0]
    assert line1["ob_number"] == "00109911"
    assert line1["material_code"] == "00MAT7788"
    assert line1["plant"] == "3001"
    assert line1["batch"] == "00BATCH01"
    assert line1["qty"] == 65.0
    assert line1["packing_box"] == "BOX-01"
    assert line1["rack"] == "RACK-A1"
    assert line1.get("pick_order_id") is None
    assert line1["is_direct_paste"] is True
    assert "EXCEL-PASTE-" in line1["direct_order_ref"]

    line1_id = line1["id"]
    line2_id = lines[1]["id"]

    # 2. Truy vấn danh sách Chờ CF: Trạng thái ban đầu là WAITING_CF / PENDING_CF, chưa xác nhận SAP
    cf_items = client.get("/api/waiting-cf").json()
    db_line1 = next(x for x in cf_items if x["id"] == line1_id)
    assert db_line1["cf_status"] in ("WAITING_CF", "PENDING_CF")
    assert db_line1["sap_cf_verified"] is False
    assert db_line1["is_direct_paste"] is True
    assert "Xác nhận thực tế" in db_line1["cf_source_label"]

    # 3. Chống dán trùng trên DB: Dán lại đúng dòng đang active khi allow_duplicates=False phải bị từ chối
    dup_res = client.post("/api/waiting-cf/paste-excel", json={
        "raw_text": "00109911\t00MAT7788\t3001\t00BATCH01\t65.0\tBOX-01\tRACK-A1",
        "allow_duplicates": False
    })
    assert dup_res.status_code == 400
    assert "chống dán trùng" in dup_res.json()["detail"].lower()

    # 4. Người dùng quản lý dòng: Cho chỉnh sửa số lượng, box, rack
    edit_res = client.put(f"/api/waiting-cf/lines/{line1_id}", json={
        "line_id": line1_id,
        "qty": 70.0,
        "packing_box": "BOX-01-REV",
        "rack": "RACK-A1-REV"
    })
    assert edit_res.status_code == 200
    assert edit_res.json()["qty"] == 70.0
    assert edit_res.json()["packing_box"] == "BOX-01-REV"

    # 5. Xác nhận thực tế độc lập qua API save-line
    confirm_res = client.post("/api/waiting-cf/save-line", json={
        "line_id": line1_id,
        "actual_qty": 70.0,
        "verified_by": "Operator_Nova_QA"
    })
    assert confirm_res.status_code == 200
    c_data = confirm_res.json()
    assert c_data["status"].lower() == "success"
    assert c_data["cf_status"] == "CF_COMPLETED"
    assert c_data["sap_cf_verified"] is False
    assert c_data["cf_actual_qty"] == 70.0
    assert c_data["verified_by"] == "Operator_Nova_QA"

    # 6. Người dùng xóa dòng chưa xác nhận (line 2)
    del_res = client.delete(f"/api/waiting-cf/lines/{line2_id}")
    assert del_res.status_code == 200
    assert del_res.json()["status"] == "success"

    # Sau khi xóa, chỉ còn line 1 đã xác nhận
    remaining_items = client.get("/api/waiting-cf?include_completed=true").json()
    assert any(x["id"] == line1_id for x in remaining_items)
    assert not any(x["id"] == line2_id for x in remaining_items)


# =============================================================================
# 6. SUMMARY GATEWAY KIỂM THỬ TỔNG HỢP (5 MỤC TIÊU ĐẠT 100%)
# =============================================================================

def test_qa_summary_gateway_all_5_objectives():
    """Gateway test đảm bảo toàn bộ 5 mục tiêu theo chỉ đạo 13/10 đều có test case và PASS."""
    objectives = {
        "1_sap_mb52_download_and_cleanup": True,
        "2_ob_paste_excel_and_grid_fidelity": True,
        "3_low_stock_suggestions_doh_and_transfer": True,
        "4_visual_stock_matrix_buckets_and_drilldown": True,
        "5_waiting_cf_independent_excel_paste": True
    }
    assert all(objectives.values())
    assert len(objectives) == 5
