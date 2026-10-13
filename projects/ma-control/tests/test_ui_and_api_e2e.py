# -*- coding: utf-8 -*-
"""End-to-End API and Excel-like Grid UI Verification Suite."""

import re
import json
import subprocess
from pathlib import Path
import pytest
from fastapi.testclient import TestClient
from app.main import app
from app.core.database import init_db
from app.services.ob_service import OBService

@pytest.fixture(scope="module", autouse=True)
def setup_db():
    init_db()

def test_ui_html_mockup_elements():
    client = TestClient(app)
    resp = client.get("/")
    assert resp.status_code == 200
    html = resp.text

    # Header & Badges
    assert "MA CONTROL — Điều hành FG" in html
    assert "sap-status-label" in html
    assert "SAP đã kết nối" not in html
    assert "RC52 • FG" in html

    # 5-Step Stepper
    assert "1. Tải tồn" in html
    assert "2. Điền OB" in html
    assert "3. Ra đơn pick" in html
    assert "4. Người dùng CF" in html
    assert "5. BBBG" in html

    # Tabs & Toolbar Buttons
    assert "Body • 3001" in html
    assert "EDC • 4007" in html
    assert "Lấy từ tồn" in html
    assert "Dán từ Excel" in html
    assert "+ Thêm dòng" in html
    assert "Kiểm tra" in html
    assert "Điền OB" in html
    assert "Ra đơn pick" in html

    # Excel Columns A-I
    for letter in ["A", "B", "C", "D", "E", "F", "G", "H"]:
        assert f'<span class="col-excel-letter">{letter}</span>' in html

    # Pick Orders Section
    assert "Đơn pick đã tạo" in html
    assert "Mở danh sách chờ CF" in html

def test_full_5_step_e2e_api_workflow():
    client = TestClient(app)

    # Step 1: Upload Stock
    s_resp = client.post("/api/stock/upload", json={
        "source": "E2E_TEST",
        "plant": "3001",
        "sloc": "1003",
        "items": [
            {"material_code": "BIW70002022BA", "batch": "0026101388", "unrestricted_qty": 500.0, "unit": "PCS"}
        ]
    })
    assert s_resp.status_code == 200

    # Step 2: Save & Check & Fill OB
    save_resp = client.post("/api/ob/batch-save", json={
        "branch_tab": "BODY_3001",
        "items": [
            {
                "item_uid": "e2e-uid-1",
                "material_code": "BIW70002022BA",
                "model": "VF7",
                "req_qty": 150.0,
                "batch": "0026101388",
                "shop": "BODY"
            }
        ]
    })
    assert save_resp.status_code == 200

    check_resp = client.post("/api/ob/check", json={"branch_tab": "BODY_3001"})
    assert check_resp.status_code == 200

    fill_resp = client.post("/api/ob/fill", json={"branch_tab": "BODY_3001"})
    assert fill_resp.status_code == 200

    # Step 3: Create Pick Order
    pick_resp = client.post("/api/pick-orders", json={
        "branch_tab": "BODY_3001",
        "shop": "BODY",
        "model": "VF7",
        "items": [
            {
                "material_code": "BIW70002022BA",
                "model": "VF7",
                "req_qty": 150.0,
                "batch": "0026101388",
                "ob_number": "1012649555"
            }
        ]
    })
    assert pick_resp.status_code == 200
    po_id = pick_resp.json()["id"]

    # Verify BBBG blocked before Step 4
    blocked_resp = client.post("/api/bbbg/create", json={
        "pick_order_id": po_id,
        "sender_name": "E2E_Sender"
    })
    assert blocked_resp.status_code == 400

    # Step 4: Manual CF Verification
    cf_resp = client.post(f"/api/pick-orders/{po_id}/verify-cf", json={
        "verified_by": "E2E_Approver",
        "actual_quantities": {str(line["id"]): line["qty"] for line in client.get(f"/api/pick-orders/{po_id}").json()["items"]}
    })
    assert cf_resp.status_code == 200
    assert cf_resp.json()["cf_status"] == "Đã CF"

    # Step 5: Generate BBBG
    bbbg_resp = client.post("/api/bbbg/create", json={
        "pick_order_id": po_id,
        "sender_name": "E2E_Sender",
        "receiver_name": "E2E_Receiver",
        "storage_bin": "BIN_VF7_01"
    })
    assert bbbg_resp.status_code == 200
    b_data = bbbg_resp.json()
    assert b_data["status"] == "success"
    assert b_data["model"] == "VF7"

def test_full_5_step_edc_4007_and_body_3001_lifecycle():
    """Comprehensive QA verification of 5-step flow across BODY_3001 & EDC_4007."""
    client = TestClient(app)
    ob_svc = OBService()

    # Step 1: Upload Stock across zones 3093, 3000, 3001 without double counting
    s_resp = client.post("/api/stock/upload", json={
        "source": "NOVA_QA_SNAPSHOT",
        "plant": "3001",
        "sloc": "3093",
        "items": [
            {"material_code": "00BIW71001001", "batch": "00260015", "plant": "3001", "sloc": "3093", "unrestricted_qty": 300.0, "unit": "PCS"},
            {"material_code": "00BIW71001001", "batch": "00261013", "plant": "3001", "sloc": "3000", "unrestricted_qty": 25.0, "unit": "PCS"},
            {"material_code": "00BIW71001001", "batch": "00261014", "plant": "3001", "sloc": "3001", "unrestricted_qty": 80.0, "unit": "PCS"},
            {"material_code": "EDC88001122AA", "batch": "00261015", "plant": "4007", "sloc": "4004", "unrestricted_qty": 200.0, "unit": "PCS"},
        ]
    })
    assert s_resp.status_code == 200

    summary = client.get("/api/stock/summary?plant=3001").json()
    sloc_map = {z["sloc"]: z for z in summary}
    assert "3093" in sloc_map and "3000" in sloc_map and "3001" in sloc_map

    stock_items = client.get("/api/stock/items?material_code=00BIW71001001").json()
    assert any(it["priority_flag"] == "PRIORITY_OUT" and it["batch"] == "00260015" for it in stock_items)
    assert any(it["sloc"] == "3000" and it["pull_line_warning"] and "Cảnh báo kéo line" in it["pull_line_warning"] for it in stock_items)

    # Step 2A: Body 3001 Draft Delete API & Shortage check
    del_ok = client.post("/api/ob/delete-draft-3001?ob_number=1012646321").json()
    assert del_ok["code"] == "DELETE_VERIFIED"
    # Verify no guessing from suffix 99 without simulated_condition
    del_no_guess = client.post("/api/ob/delete-draft-3001?ob_number=1012646399").json()
    assert del_no_guess["code"] == "DELETE_VERIFIED"
    del_sto = client.post("/api/ob/delete-draft-3001?ob_number=1012646399&simulated_condition=WAIT_STO_LOCK").json()
    assert del_sto["code"] == "WAIT_STO_LOCK"

    # Step 2B: EDC 4007 Dominant STO, WAIT_STO_LOCK, OB_DELETED_REPLAN, Check & Fill
    dom_sto = ob_svc.calculate_dominant_sto(
        ["EDC88001122AA", "EDC88001199BB"],
        {
            "EDC88001122AA": [{"sto_number": "4500999001", "due_qty": 120}],
            "EDC88001199BB": [{"sto_number": "4500999001", "due_qty": 80}],
        }
    )
    assert dom_sto == "4500999001"
    replan = ob_svc.replan_deleted_ob("EDC88001122AA", "1012647999")
    assert replan["action"] == "OB_DELETED_REPLAN"

    edc_save = client.post("/api/ob/batch-save", json={
        "branch_tab": "EDC_4007",
        "items": [
            {
                "item_uid": "EDC_4007:nova-edc-ok",
                "material_code": "EDC88001122AA",
                "model": "VF9",
                "req_qty": 50.0,
                "batch": "00261015",
                "shop": "EDC",
                "sto_number": dom_sto,
                "status": "Chưa kiểm tra"
            },
            {
                "item_uid": "EDC_4007:nova-edc-shortage",
                "material_code": "EDC_MISSING_MAT",
                "model": "VF9",
                "req_qty": 999.0,
                "batch": "00261015",
                "shop": "EDC",
                "sto_number": dom_sto,
                "status": "Chưa kiểm tra"
            }
        ]
    })
    assert edc_save.status_code == 200

    before_items = client.get("/api/ob/items?branch_tab=EDC_4007").json()
    uid_map_before = {it["item_uid"]: it for it in before_items if it.get("item_uid") in ("EDC_4007:nova-edc-ok", "EDC_4007:nova-edc-shortage")}
    assert uid_map_before["EDC_4007:nova-edc-ok"]["status"] == "Chưa kiểm tra"

    edc_check = client.post("/api/ob/check", json={
        "branch_tab": "EDC_4007",
        "item_ids": [uid_map_before["EDC_4007:nova-edc-ok"]["id"], uid_map_before["EDC_4007:nova-edc-shortage"]["id"]]
    }).json()
    checked_map = {it["item_uid"]: it for it in edc_check["items"]}
    assert checked_map["EDC_4007:nova-edc-ok"]["status"] == "Sẵn sàng điền"
    assert checked_map["EDC_4007:nova-edc-shortage"]["status"] == "Thiếu SL OB"
    assert checked_map["EDC_4007:nova-edc-shortage"]["missing_qty"] == 999.0

    edc_fill = client.post("/api/ob/fill", json={
        "branch_tab": "EDC_4007",
        "item_ids": [uid_map_before["EDC_4007:nova-edc-ok"]["id"], uid_map_before["EDC_4007:nova-edc-shortage"]["id"]]
    }).json()
    assert edc_fill["filled_count"] == 1
    filled_item = edc_fill["items"][0]
    assert filled_item["item_uid"] == "EDC_4007:nova-edc-ok"
    assert filled_item["status"] == "Chờ pick / CF"
    assert filled_item["ob_number"] != ""

    # Step 3: Create Pick Order & verify PICK-YYYYMMDD-XXX and 'Chờ người dùng CF'
    pick_resp = client.post("/api/pick-orders", json={
        "branch_tab": "EDC_4007",
        "shop": "EDC",
        "model": "VF9",
        "items": [
            {
                "id": filled_item["id"],
                "material_code": filled_item["material_code"],
                "batch": "00261015",
                "req_qty": 50.0,
                "sto_number": dom_sto,
                "ob_number": filled_item["ob_number"]
            }
        ]
    })
    assert pick_resp.status_code == 200
    po_data = pick_resp.json()
    assert re.match(r"^PICK-\d{8}-\d{3}$", po_data["pick_order_code"])
    assert po_data["status_label"] == "Chờ người dùng CF"

    orders_list = client.get("/api/pick-orders").json()
    found_po = next(o for o in orders_list if o["id"] == po_data["id"])
    assert found_po["status"] == "Chờ người dùng CF"

    # Step 4: Verify unconfirmed CF blocks BBBG, then manual CF unlocks it
    block_resp = client.post("/api/bbbg/create", json={
        "pick_order_id": po_data["id"],
        "sender_name": "Nova_QA"
    })
    assert block_resp.status_code == 400
    assert "BẤT BIẾN NGHIỆP VỤ BỊ VI PHẠM" in block_resp.json()["detail"]

    cf_resp = client.post(f"/api/pick-orders/{po_data['id']}/verify-cf", json={
        "verified_by": "Nova_QA",
        "actual_quantities": {str(line["id"]): line["qty"] for line in client.get(f"/api/pick-orders/{po_data['id']}").json()["items"]},
        "notes": "Verified CF manually"
    })
    assert cf_resp.status_code == 200
    assert cf_resp.json()["cf_status"] == "Đã CF"

    # Step 5: BBBG creation, LIMO mapping, append-only ledger, GOS attachments
    bbbg_resp = client.post("/api/bbbg/create", json={
        "pick_order_id": po_data["id"],
        "sender_name": "GL_EDC_01",
        "receiver_name": "TL_MA_02",
        "storage_bin": "BIN_EDC_09"
    })
    assert bbbg_resp.status_code == 200
    bbbg_json = bbbg_resp.json()
    assert re.match(r"^BBBG-\d{8}-\d{3}$", bbbg_json["bbbg_no"])
    assert bbbg_json["model"] == "VF9"
    assert bbbg_json["general_log_saved"] == 0
    assert bbbg_json["is_draft"] is True
    assert client.get(f"/api/pick-orders/{po_data['id']}").json()["status"] == "Đã CF"
    assert bbbg_json["gos_bbbg_attached"] == 0
    assert bbbg_json["gos_vat_attached"] == 0

def test_excel_grid_ux_js_execution_paste_nav_inline_edit_and_chips():
    """Executes app.js in Node runtime with mock DOM to verify Ctrl+V multi-cell paste,
    leading-zero preservation, Tab/Enter cell navigation, inline editing, and status chips.
    """
    app_js_path = Path(__file__).resolve().parent.parent / "app" / "static" / "js" / "app.js"
    styles_css_path = Path(__file__).resolve().parent.parent / "app" / "static" / "css" / "styles.css"
    assert app_js_path.exists()
    assert styles_css_path.exists()

    css_text = styles_css_path.read_text(encoding="utf-8")
    assert ".chip-ready" in css_text
    assert ".chip-pending" in css_text
    assert ".chip-insufficient" in css_text
    assert ".chip-waiting-pick" in css_text
    assert "--border-excel: #107c41" in css_text

    node_script = f"""
    const fs = require('fs');
    const code = fs.readFileSync({json.dumps(str(app_js_path))}, 'utf8');

    const listeners = {{}};
    const elements = {{}};
    function makeEl(tag) {{
      let _html = '';
      return {{
        tagName: tag.toUpperCase(),
        dataset: {{}},
        style: {{}},
        classList: {{
          _set: new Set(),
          add(c) {{ this._set.add(c); }},
          remove(c) {{ this._set.delete(c); }},
          contains(c) {{ return this._set.has(c); }}
        }},
        children: [],
        get innerHTML() {{ return _html; }},
        set innerHTML(v) {{ _html = v; if (v === '') this.children = []; }},
        textContent: '',
        value: '',
        appendChild(child) {{ this.children.push(child); return child; }},
        addEventListener(ev, fn) {{ this['on_' + ev] = fn; }},
        scrollIntoView() {{}},
        focus() {{}},
        select() {{}}
      }};
    }}
    elements['grid-body'] = makeEl('tbody');
    elements['selection-summary'] = makeEl('div');
    elements['pick-orders-body'] = makeEl('tbody');
    elements['badge-cf-count'] = makeEl('span');

    global.window = global;
    global.window.addEventListener = (ev, fn) => {{ listeners[ev] = fn; }};
    global.document = {{
      activeElement: {{ tagName: 'BODY' }},
      addEventListener: (ev, fn) => {{ listeners[ev] = fn; }},
      getElementById: (id) => elements[id] || null,
      createElement: (tag) => makeEl(tag),
      querySelectorAll: () => [],
      querySelector: () => makeEl('td')
    }};
    global.fetch = async () => ({{ ok: true, json: async () => [] }});
    global.alert = () => {{}};

    eval(code);
    if (listeners['DOMContentLoaded']) listeners['DOMContentLoaded']();

    // Initialize 2 empty rows
    window.gridData = [];
    window.addNewRow(false);
    window.activeRowIndex = 0;
    window.activeColKey = 'material_code';

    // 1. Test multi-row, multi-column Excel paste with leading zeros
    const excelClipboard = "000123456789\\tVF7\\t125\\t0026101301\\tBODY\\t0045001234\\r\\n000987654321\\tVF8 LIMO\\t60\\t0026101309\\tEDC\\t0045009999\\r\\n";
    window.applyPastedText(excelClipboard);

    if (window.gridData.length < 2) throw new Error("Expected at least 2 rows after paste");
    const r0 = window.gridData[0];
    const r1 = window.gridData[1];

    // Verify leading zeros preserved
    if (r0.material_code !== "000123456789") throw new Error("Leading zero lost in r0.material_code: " + r0.material_code);
    if (r0.batch !== "0026101301") throw new Error("Leading zero lost in r0.batch: " + r0.batch);
    if (r0.sto_number !== "0045001234") throw new Error("Leading zero lost in r0.sto_number: " + r0.sto_number);
    if (r0.required_qty !== 125) throw new Error("Qty mismatch: " + r0.required_qty);
    if (r1.material_code !== "000987654321" || r1.batch !== "0026101309") throw new Error("Row 2 paste mismatch");

    // 2. Test keyboard Tab & Enter navigation
    window.selectCell(0, 'model');
    listeners['keydown']({{ key: 'Tab', shiftKey: false, preventDefault() {{}} }});
    const afterTabCol = window.activeColKey;
    listeners['keydown']({{ key: 'Enter', shiftKey: false, preventDefault() {{}} }});
    const afterEnterRow = window.activeRowIndex;

    if (afterTabCol !== 'material_code') throw new Error("Tab navigation failed: " + afterTabCol);
    if (afterEnterRow !== 1) throw new Error("Enter navigation failed: " + afterEnterRow);

    // 3. Test Inline Edit via double-click & Enter commit with leading zero
    const tr0 = elements['grid-body'].children[0];
    const tdBatch = tr0.children.find(td => td.dataset.colKey === 'batch');
    tdBatch.on_dblclick();
    const inputEl = tdBatch.children[0];
    inputEl.value = "0000998877";
    inputEl.on_keydown({{ key: 'Enter', preventDefault() {{}} }});
    if (window.gridData[0].batch !== "0000998877") throw new Error("Inline edit lost leading zero: " + window.gridData[0].batch);

    // 4. Verify status chip rendering for all 4 states
    const statuses = ['Sẵn sàng điền', 'Chưa kiểm tra', 'Thiếu SL OB', 'Chờ pick / CF'];
    const renderedChips = statuses.map(s => {{
      window.gridData[0].status = s;
      window.renderGrid();
      const statusTd = elements['grid-body'].children[0].children.find(td => td.dataset.colKey === 'status');
      return statusTd.innerHTML;
    }});

    console.log(JSON.stringify({{
      paste_ok: true,
      r0_material: r0.material_code,
      r0_batch: window.gridData[0].batch,
      tab_col: afterTabCol,
      enter_row: afterEnterRow,
      chips: renderedChips
    }}));
    """
    proc = subprocess.run(["node", "-e", node_script], capture_output=True, text=True, encoding="utf-8", check=True)
    out = json.loads(proc.stdout.strip())
    assert out["paste_ok"] is True
    assert out["r0_material"] == "000123456789"
    assert out["r0_batch"] == "0000998877"
    assert out["tab_col"] == "material_code"
    assert out["enter_row"] == 1
    assert 'chip-ready">Sẵn sàng điền' in out["chips"][0]
    assert 'chip-pending">Chưa kiểm tra' in out["chips"][1]
    assert 'chip-insufficient">⚠ Thiếu SL OB' in out["chips"][2]
    assert 'chip-waiting-pick">Chờ pick / CF' in out["chips"][3]


def test_navigation_views_and_pilot_labels():
    client = TestClient(app)
    resp = client.get("/")
    assert resp.status_code == 200
    html = resp.text

    # All 7 navigation views must exist in HTML
    for view_id in ["view-dashboard", "view-stock", "view-grid", "view-waiting-cf", "view-bbbg", "view-history", "view-config"]:
        assert f'id="{view_id}"' in html

    # Explicit pilot notice label for unimplemented external integrations
    assert "Chưa hỗ trợ trong bản pilot" in html
    assert "SAP live write" in html or "SAP Live Write" in html
    assert "GOS" in html
    assert "email" in html or "Email" in html
    assert "in" in html or "In" in html
    assert "UNC" in html

    # Verify backend endpoints for navigation views exist and respond
    assert client.get("/api/stock/summary").status_code == 200
    assert client.get("/api/waiting-cf").status_code == 200
    assert client.get("/api/bbbg/records").status_code == 200
    assert client.get("/api/jobs/events?limit=10").status_code == 200
    assert client.get("/api/settings").status_code == 200
    assert client.get("/api/materials").status_code == 200
    assert client.get("/api/model-mappings").status_code == 200
