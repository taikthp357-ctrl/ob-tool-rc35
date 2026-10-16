# -*- coding: utf-8 -*-
"""Test Suite for MA CONTROL FG Zero-Port Architecture & Enterprise Launcher.

Validates:
1. Zero open TCP/UDP sockets (Corporate firewall and Endpoint Security compliance).
2. In-process data synchronization from SQLite to data.js via web_bridge.
3. Zero-Port workflow action dispatching (Check OB, Fill OB, Create Pick, Confirm CF, BBBG).
4. Strict enforcement of Manual CF Gate (Zero Auto-CF Invariant).
5. Manifest SHA256 integrity and Staged Enterprise Deployment.
"""

import json
import os
import socket
import sys
from pathlib import Path
import pytest

from app.core.config import BASE_DIR, DATA_DIR, DB_PATH
from app.core.database import get_connection, init_db
from app.web_bridge import get_system_snapshot, sync_web_dashboard, ZeroPortActionDispatcher
from app.launcher_core import EnterpriseLauncher, compute_sha256


def test_zero_listening_sockets_invariant():
    """BẤT BIẾN ZERO-PORT: Tuyệt đối không mở cổng lắng nghe mạng (TCP 8012 / 8013)."""
    # Test that port 8012 is NOT in use by any listening socket
    for port in [8012, 8013]:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            result = s.connect_ex(("127.0.0.1", port))
            # result != 0 means port is NOT listening (connection refused), which is REQUIRED for Zero-Port!
            assert result != 0, f"VI PHẠM BẤT BIẾN ZERO-PORT: Phát hiện port {port} đang mở lắng nghe trên 127.0.0.1!"


def test_web_bridge_sync_and_data_js_generation():
    """Kiểm tra đồng bộ dữ liệu SQLite sang data.js trong tiến trình nội bộ (In-Process)."""
    init_db()
    res = sync_web_dashboard()
    assert res is not None
    assert "version" in res
    assert res.get("zero_port_mode") is True
    assert "ob_items_body_3001" in res
    assert "pick_orders" in res

    # Check generated data.js file
    data_js_path = BASE_DIR / "app" / "web" / "data.js"
    assert data_js_path.exists(), "data.js không được sinh ra trong app/web/"

    content = data_js_path.read_text(encoding="utf-8")
    assert "window.MA_CONTROL_DATA =" in content
    assert "window.OB_ITEMS_BODY =" in content
    assert "window.PICK_ORDERS =" in content
    assert "100% Zero-Port compliant" in content


def test_in_process_action_dispatcher_workflow():
    """Kiểm tra điều phối luồng nghiệp vụ in-process qua ZeroPortActionDispatcher."""
    init_db()
    # Step 2a: Check OB items
    res_check = ZeroPortActionDispatcher.check_ob_items(branch="BODY_3001")
    assert res_check["status"] == "SUCCESS"
    assert "checked_count" in res_check

    # Step 2b: Fill OB items
    res_fill = ZeroPortActionDispatcher.fill_ob_items(branch="BODY_3001")
    assert res_fill["status"] == "SUCCESS"
    assert "filled_count" in res_fill


def test_zero_auto_cf_invariant_enforcement_in_zero_port():
    """BẤT BIẾN ZERO AUTO-CF: Chốt chặn kiểm tra thủ công tại Bước 4 trong chế độ Zero-Port."""
    init_db()
    # 1. Insert a verified test OB item into database first
    conn = get_connection()
    cur = conn.execute("""
        INSERT INTO ob_items (material_code, model, required_qty, batch, shop, ob_number, status, branch, branch_tab)
        VALUES ('BIW73012711CA', 'VF5 RHD', 50, '26101399', 'BODY', '1012649999', 'READY_FOR_PICK', 'BODY_3001', 'BODY_3001')
    """)
    row_id = cur.lastrowid
    conn.commit()
    conn.close()

    test_items = [
        {
            "id": row_id,
            "material_code": "BIW73012711CA",
            "model": "VF5 RHD",
            "required_qty": 50,
            "batch": "26101399",
            "shop": "BODY",
            "ob_number": "1012649999",
            "status": "READY_FOR_PICK"
        }
    ]
    pick_res = ZeroPortActionDispatcher.create_pick_order(
        shop="BODY", model="VF5 RHD", items=test_items, branch_tab="BODY_3001"
    )
    pick_no = pick_res.get("pick_order_code") or pick_res.get("pick_order_no")
    assert pick_no is not None
    assert pick_res.get("pick_status") in ("WAITING_USER_CF", "Chờ CF")

    # 2. Invariant: Try generating BBBG directly -> MUST FAIL!
    with pytest.raises(Exception) as exc_info:
        ZeroPortActionDispatcher.generate_bbbg(
            pick_order_no=pick_no, sender_name="GL_BODY_01"
        )
    err_str = str(exc_info.value)
    assert ("BẤT BIẾN NGHIỆP VỤ BỊ VI PHẠM" in err_str or "chưa được người dùng CF" in err_str or "USER_CONFIRMED" in err_str)

    # 3. User performs manual CF confirmation
    cf_res = ZeroPortActionDispatcher.confirm_pick_order(
        pick_order_no=pick_no, user_name="GL_BODY_01"
    )
    assert cf_res is not None
    assert cf_res.get("status") in ("success", "already_confirmed", "VALIDATED", "CONFIRMED") or "pick_order_id" in cf_res

    # 4. Now BBBG generation succeeds!
    bbbg_res = ZeroPortActionDispatcher.generate_bbbg(
        pick_order_no=pick_no, sender_name="GL_BODY_01", storage_bin="BIN_ZERO_PORT"
    )
    assert bbbg_res is not None
    assert "bbbg_no" in bbbg_res or "record_id" in bbbg_res or "bbbg_code" in bbbg_res


def test_manifest_and_staged_deployment_integrity():
    """Kiểm tra tính toàn vẹn của manifest.json và bộ cài doanh nghiệp."""
    manifest_path = BASE_DIR / "dist" / "COMPANY_RELEASE_FG" / "manifest.json"
    assert manifest_path.exists(), "manifest.json chưa được sinh ra trong dist/COMPANY_RELEASE_FG/"

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["zero_port_mode"] is True
    assert manifest["network_ports_required"] == 0
    assert manifest["version"] == "v1.0.0-rc52-FG"
    assert len(manifest["files"]) > 0

    # Verify checksum of a sample file from manifest
    payload_dir = BASE_DIR / "dist" / "COMPANY_RELEASE_FG" / "payload"
    sample_file_rel, expected_hash = next(iter(manifest["files"].items()))
    sample_file_path = payload_dir / sample_file_rel
    assert sample_file_path.exists(), f"File {sample_file_rel} không tồn tại trong payload"
    assert compute_sha256(sample_file_path).lower() == expected_hash.lower(), "Sai lệch SHA256!"


def test_launcher_edge_app_mode_zero_port():
    """Kiểm tra launcher_core chuẩn bị runtime sạch và URL file:/// mà không cần cổng mạng."""
    launcher = EnterpriseLauncher(BASE_DIR)
    prep = launcher.prepare_runtime()
    assert prep["status"] == "READY"
    assert prep["zero_port"] is True
    assert Path(prep["runtime_path"]).exists()

    # Verify that index.html exists in runtime
    runtime_html = Path(prep["runtime_path"]) / "app" / "web" / "index.html"
    assert runtime_html.exists(), "app/web/index.html không tồn tại trong runtime"
    file_uri = runtime_html.as_uri()
    assert file_uri.startswith("file:///")
    assert "127.0.0.1" not in file_uri
    assert "localhost" not in file_uri
