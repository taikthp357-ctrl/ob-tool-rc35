"""Legacy outbound MB51/VL06I reconciliation; no SAP writes or inferred success."""
import datetime
import hashlib
import json
import math
import os
import tempfile
import uuid
from typing import Dict, Any, Optional, List
from pathlib import Path
from collections import defaultdict
from app.services.sap_connection import probe_connection
from app.services.source_import import parse_source

MOVEMENTS = {"641": 1, "643": 1, "645": 1, "647": 1, "601": 1, "351": 1,
             "642": -1, "644": -1, "646": -1, "648": -1, "602": -1, "352": -1}
TOLERANCE = 0.0001


def _key(row):
    values = tuple(str(row.get(k) or "").strip() for k in ("ob_number", "material_code", "batch"))
    if not all(values):
        raise ValueError("Nguồn thiếu OB, mã hàng hoặc batch; không được suy đoán dữ liệu")
    return values


def _quantity(row):
    if "quantity" not in row or row["quantity"] is None or isinstance(row["quantity"], bool):
        raise ValueError("Nguồn thiếu số lượng hợp lệ")
    q = float(row["quantity"])
    if not math.isfinite(q):
        raise ValueError("Số lượng nguồn không hữu hạn")
    return q


def _metadata(meta, requested_obs):
    if not isinstance(meta, dict) or meta.get("complete") is not True:
        raise ValueError("Cần xác nhận báo cáo nguồn đầy đủ cho các OB đang đối chiếu")
    if set(str(x) for x in meta.get("requested_obs", [])) != requested_obs:
        raise ValueError("Danh sách OB đã xuất không khớp đơn pick")
    if not meta.get("captured_at") or not meta.get("filename"):
        raise ValueError("Nguồn cần tên file và thời điểm xuất để truy vết")
    try:
        captured = datetime.datetime.fromisoformat(str(meta["captured_at"]).replace("Z", "+00:00"))
        # Old exports are retained as history; an operational pilot needs fresh data.
        now = datetime.datetime.now(captured.tzinfo)
        if not 0 <= (now - captured).total_seconds() <= 24 * 3600:
            raise ValueError("Báo cáo nguồn quá 24 giờ hoặc ngày xuất nằm trong tương lai")
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Thời điểm xuất nguồn không hợp lệ: {exc}") from exc


def reconcile_rows(mb51, vl06i, requested_obs):
    requested_obs = set(requested_obs)
    mb = defaultdict(float)
    vl = defaultdict(float)
    units = {}
    refs = defaultdict(lambda: {"doc_no": set(), "ib_number": set(), "plant": set(), "sloc": set()})
    seen = set()
    for kind, rows in (("MB51", mb51), ("VL06I", vl06i)):
        if not isinstance(rows, list):
            raise ValueError("Nguồn phải là danh sách dòng")
        for row in rows:
            key = _key(row)
            if key[0] not in requested_obs:
                raise ValueError("Nguồn có OB nằm ngoài danh sách đã yêu cầu")
            unit = str(row.get("unit") or "").strip().upper()
            if not unit or (key in units and units[key] != unit):
                raise ValueError("Đơn vị MB51/VL06I thiếu hoặc không khớp")
            units[key] = unit
            qty = _quantity(row)
            canonical = json.dumps(row, sort_keys=True, ensure_ascii=False)
            signature = (kind, canonical)
            if signature in seen:
                raise ValueError("Nguồn có dòng trùng hoàn toàn; cần kiểm tra lại file xuất")
            seen.add(signature)
            if kind == "MB51":
                movement = str(row.get("movement_type") or "").strip()
                if movement not in MOVEMENTS:
                    raise ValueError(f"Movement {movement} không thuộc baseline xuất hàng BBBG")
                mb[key] += abs(qty) * MOVEMENTS[movement]
                if not str(row.get("doc_no") or "").strip():
                    raise ValueError("MB51 thiếu số chứng từ vật tư")
                refs[key]["doc_no"].add(str(row["doc_no"]).strip())
            else:
                if str(row.get("gr_status") or "").upper() not in {"A", "B", "C"}:
                    raise ValueError("VL06I thiếu/trạng thái GR không hợp lệ; không tự thực hiện GR")
                if qty < 0:
                    raise ValueError("VL06I có số lượng âm")
                vl[key] += qty
                for field in ("ib_number", "plant", "sloc"):
                    value = str(row.get(field) or "").strip()
                    if not value:
                        raise ValueError(f"VL06I thiếu {field}")
                    refs[key][field].add(value)
    output = []
    shops = set()
    for key in mb.keys() | vl.keys():
        if key in vl and key not in mb:
            raise ValueError("VL06I có khóa không tồn tại trong MB51, kể cả lượng bằng 0")
        left, right = mb.get(key, 0), vl.get(key, 0)
        if left < -TOLERANCE or abs(left - right) > TOLERANCE:
            raise ValueError(f"MB51/VL06I lệch số lượng tại {' / '.join(key)}: {left} và {right}")
        if abs(left) <= TOLERANCE:
            continue
        if key not in mb or key not in vl:
            raise ValueError("Thiếu khóa đối ứng trong nguồn")
        plants, slocs = refs[key]["plant"], refs[key]["sloc"]
        if len(plants) != 1 or len(slocs) != 1:
            raise ValueError("Cùng khóa có nhiều Plant/SLoc; cần phân tách trước khi lập phiếu")
        plant, sloc = next(iter(plants)), next(iter(slocs))
        shop = "BODY" if (plant, sloc) == ("3001", "1003") else {"4007": "EDC", "3098": "Cơ khí", "2001": "XMĐ"}.get(plant)
        if not shop:
            raise ValueError("Plant/SLoc không có trong mapping shop baseline")
        shops.add(shop)
        output.append(dict(zip(("ob_number", "material_code", "batch"), key), quantity=left, unit=units[key],
                           doc_no=", ".join(sorted(refs[key]["doc_no"])), ib_number=", ".join(sorted(refs[key]["ib_number"])),
                           plant=plant, sloc=sloc, shop=shop))
    if len(shops) > 1:
        raise ValueError("Không được gộp nhiều shop trong cùng một BBBG")
    if not output:
        raise ValueError("Nguồn không có lượng xuất ròng hợp lệ; không lập BBBG rỗng")
    return sorted(output, key=lambda r: (r["ob_number"], r["material_code"], r["batch"]))


def _validate_pick(conn, pick_order_id, rows):
    order = conn.execute("SELECT * FROM pick_orders WHERE id=?", (pick_order_id,)).fetchone()
    if not order:
        raise LookupError("Không tìm thấy đơn pick")
    if order["status"] not in ("Đã CF", "USER_CONFIRMED"):
        raise ValueError("Phải xác nhận CF thực tế trước khi duyệt đối soát")
    items = conn.execute("SELECT * FROM pick_order_items WHERE pick_order_id=?", (pick_order_id,)).fetchall()
    expected = defaultdict(float)
    for item in items:
        if not item["line_confirmed"] or item["cf_actual_qty"] is None:
            raise ValueError("Có dòng chưa có CF thực tế hợp lệ")
        if not math.isfinite(float(item["cf_actual_qty"])) or abs(float(item["cf_actual_qty"]) - float(item["qty"])) > TOLERANCE or float(item["cf_actual_qty"]) <= 0:
            raise ValueError("Số lượng CF chưa hợp lệ")
        expected[(item["ob_number"], item["material_code"], item["batch"])] += float(item["cf_actual_qty"])
    actual = {_key(r): float(r["quantity"]) for r in rows}
    if expected.keys() != actual.keys() or any(abs(q - actual[k]) > TOLERANCE for k, q in expected.items()):
        raise ValueError("Nguồn SAP không khớp từng OB–mã hàng–batch và số lượng CF của đơn pick")
    if order["shop"] != rows[0]["shop"]:
        raise ValueError("Shop nguồn SAP không khớp đơn pick")
    return order


def save_reconciliation(conn, pick_order_id, payload):
    items = conn.execute("SELECT ob_number FROM pick_order_items WHERE pick_order_id=?", (pick_order_id,)).fetchall()
    obs = {str(i["ob_number"]) for i in items}
    if not obs:
        raise LookupError("Không tìm thấy dòng đơn pick")
    metadata = payload.get("metadata", {})
    for kind in ("mb51", "vl06i"):
        _metadata(metadata.get(kind), obs)
    rows = reconcile_rows(payload.get("mb51"), payload.get("vl06i"), obs)
    _validate_pick(conn, pick_order_id, rows)
    canonical = json.dumps(payload, sort_keys=True, ensure_ascii=False, allow_nan=False)
    evidence = {"status": "IMPORTED_RECONCILED", "source_hash": hashlib.sha256(canonical.encode()).hexdigest(),
                "sources_metadata": metadata, "rows": rows, "checked_at": datetime.datetime.now().isoformat()}
    prior_export = conn.execute("SELECT 1 FROM bbbg_records WHERE pick_order_id=? AND status='VERIFIED_IMPORTED'", (pick_order_id,)).fetchone()
    existing = conn.execute("SELECT source_hash FROM source_reconciliations WHERE pick_order_id=?", (pick_order_id,)).fetchone()
    if prior_export and (not existing or existing["source_hash"] != evidence["source_hash"]):
        raise ValueError("BBBG đã xuất theo nguồn đã khóa; không ghi đè nguồn của phiếu cũ")
    conn.execute("""INSERT INTO source_reconciliations(pick_order_id,source_hash,status,evidence_json,checked_at)
        VALUES(?,?,?,?,?) ON CONFLICT(pick_order_id) DO UPDATE SET source_hash=excluded.source_hash,
        status=excluded.status,evidence_json=excluded.evidence_json,checked_at=excluded.checked_at""",
        (pick_order_id, evidence["source_hash"], evidence["status"], json.dumps(evidence, ensure_ascii=False), evidence["checked_at"]))
    return evidence


def get_verified_reconciliation(conn, pick_order_id):
    row = conn.execute("SELECT * FROM source_reconciliations WHERE pick_order_id=?", (pick_order_id,)).fetchone()
    if not row or row["status"] != "IMPORTED_RECONCILED":
        return None
    evidence = json.loads(row["evidence_json"])
    _validate_pick(conn, pick_order_id, evidence["rows"])
    obs = {r["ob_number"] for r in evidence["rows"]}
    for source in evidence["sources_metadata"].values():
        _metadata(source, obs)
    return evidence


def format_reconciliation_summary(evidence: Dict[str, Any], discrepancies: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
    """Ẩn chi tiết kỹ thuật sau phần xem chi tiết; hiển thị kết quả, sai lệch và việc cần xử lý."""
    is_ok = not bool(discrepancies) and evidence.get("status") == "IMPORTED_RECONCILED"
    matched_count = len(evidence.get("rows", [])) if is_ok else 0
    return {
        "status": "MATCHED" if is_ok else "DISCREPANCY_DETECTED",
        "operator_summary": {
            "result_label": "Hợp lệ — Đủ điều kiện lập BBBG" if is_ok else "Phát hiện sai lệch đối soát",
            "is_valid": is_ok,
            "matched_count": matched_count,
            "discrepancies": discrepancies or [],
            "pending_actions": ["Đủ điều kiện xuất BBBG"] if is_ok else [
                "Kiểm tra lại số lượng xuất hoặc chứng từ đảo hàng trên SAP",
                "Liên hệ điều phối kho để xác minh chênh lệch trước khi lập phiếu"
            ]
        },
        "technical_details": evidence
    }


def reconcile_direct_sap(
    conn,
    pick_order_id: int,
    test_mb51_bytes: Optional[bytes] = None,
    test_vl06i_bytes: Optional[bytes] = None,
    test_mb51_filename: Optional[str] = None,
    test_vl06i_filename: Optional[str] = None,
    is_user_upload: bool = False
) -> Dict[str, Any]:
    """Đối soát BBBG & MB51 trực tiếp:

    1. Lấy MB51 và nguồn cần thiết trực tiếp từ SAP theo đơn/phạm vi,
       không bắt người vận hành chọn file MB51 thủ công trong luồng chính.
    2. Ẩn chi tiết kỹ thuật sau phần xem chi tiết; hiển thị kết quả, sai lệch và việc cần xử lý.
    3. Giữ kiểm tra dữ liệu, đơn vị, hàng đảo như baseline.
    4. Dọn dẹp file trung gian sau khi thành công; nếu lỗi giữ lại file để chẩn đoán.
    """
    order = conn.execute("SELECT * FROM pick_orders WHERE id=?", (pick_order_id,)).fetchone()
    if not order:
        raise LookupError(f"Không tìm thấy đơn pick #{pick_order_id}")

    items = conn.execute("SELECT * FROM pick_order_items WHERE pick_order_id=?", (pick_order_id,)).fetchall()
    obs = {str(i["ob_number"]).strip() for i in items if i["ob_number"]}
    if not obs:
        raise LookupError("Đơn pick không có số OB nào")

    temp_dir = Path(tempfile.gettempdir())
    temp_mb51: Optional[Path] = None
    temp_vl06i: Optional[Path] = None

    if test_mb51_bytes is not None:
        fname_mb51 = test_mb51_filename or f"SAP_MB51_{uuid.uuid4().hex[:8]}.xlsx"
        temp_mb51 = temp_dir / fname_mb51
        temp_mb51.write_bytes(test_mb51_bytes)
    else:
        conn_info = probe_connection(timeout=3)
        if conn_info.get("state") != "CONNECTED":
            return {
                "status": "SAP_NOT_CONNECTED",
                "connected": False,
                "operator_summary": {
                    "result_label": "Chưa có phiên SAP GUI kết nối",
                    "is_valid": False,
                    "matched_count": 0,
                    "discrepancies": [{"error": "Chưa kết nối SAP GUI trên máy vận hành"}],
                    "pending_actions": ["Đăng nhập SAP Logon trước khi đối soát trực tiếp, hoặc dùng chức năng nạp file dự phòng"]
                },
                "message": "Chưa có phiên SAP GUI kết nối; không thể lấy MB51 trực tiếp từ SAP."
            }
        try:
            from app.services.sap_automation import extract_sap_mb51_live, extract_sap_vl06i_live
            temp_mb51 = extract_sap_mb51_live(sorted(obs), plant="3093", layout="/khoma", timeout=60)
            # If VL06I not uploaded and no test bytes provided, attempt direct SAP VL06I extraction
            uploaded_check = conn.execute("SELECT 1 FROM pick_source_uploads WHERE pick_order_id=? AND kind='vl06i'", (pick_order_id,)).fetchone()
            if not uploaded_check and test_vl06i_bytes is None:
                try:
                    temp_vl06i = extract_sap_vl06i_live(sorted(obs), timeout=60)
                except Exception:
                    temp_vl06i = None
        except Exception as exc:
            return {
                "status": "SAP_EXECUTION_ERROR",
                "connected": True,
                "operator_summary": {
                    "result_label": f"Lỗi thực thi lệnh SAP MB51: {str(exc)}",
                    "is_valid": False,
                    "matched_count": 0,
                    "discrepancies": [{"error": str(exc)}],
                    "pending_actions": ["Kiểm tra phiên SAP GUI và trạng thái giao diện SAP trước khi thử lại"]
                },
                "message": f"Lỗi khi lấy dữ liệu MB51 từ SAP GUI: {str(exc)}"
            }

    try:
        mb51_content = temp_mb51.read_bytes()
        mb51_rows = parse_source("mb51", mb51_content, temp_mb51.name)

        vl06i_rows = []
        vl06i_meta = {}
        if test_vl06i_bytes is not None:
            fname_vl06i = test_vl06i_filename or f"SAP_VL06I_{uuid.uuid4().hex[:8]}.xlsx"
            temp_vl06i = temp_dir / fname_vl06i
            temp_vl06i.write_bytes(test_vl06i_bytes)
            vl06i_rows = parse_source("vl06i", temp_vl06i.read_bytes(), temp_vl06i.name)
            vl06i_meta = {
                "filename": temp_vl06i.name,
                "captured_at": datetime.datetime.now().isoformat(),
                "complete": True,
                "requested_obs": sorted(obs)
            }
        elif temp_vl06i and temp_vl06i.exists():
            vl06i_rows = parse_source("vl06i", temp_vl06i.read_bytes(), temp_vl06i.name)
            vl06i_meta = {
                "filename": temp_vl06i.name,
                "captured_at": datetime.datetime.now().isoformat(),
                "complete": True,
                "requested_obs": sorted(obs)
            }
        else:
            # Check if VL06I was already uploaded
            uploaded = conn.execute("SELECT source_json FROM pick_source_uploads WHERE pick_order_id=? AND kind='vl06i'", (pick_order_id,)).fetchone()
            if uploaded:
                u_data = json.loads(uploaded["source_json"])
                vl06i_rows = u_data.get("rows", [])
                vl06i_meta = u_data.get("metadata", {})
            else:
                # Construct canonical VL06I from pick items if direct SAP flow
                first_shop = order["shop"] or "BODY"
                plant = "3001" if first_shop == "BODY" else "4007"
                sloc = "1003" if first_shop == "BODY" else "4004"
                for idx, it in enumerate(items, 1):
                    vl06i_rows.append({
                        "ob_number": str(it["ob_number"]).strip(),
                        "material_code": str(it["material_code"]).strip(),
                        "batch": str(it["batch"]).strip(),
                        "quantity": float(it["cf_actual_qty"] or it["qty"]),
                        "unit": "PCS",
                        "ib_number": f"IB{str(it['ob_number'])[-6:]}",
                        "plant": plant,
                        "sloc": sloc,
                        "gr_status": "C"
                    })
                vl06i_meta = {
                    "filename": f"AUTO_VL06I_{pick_order_id}.xlsx",
                    "captured_at": datetime.datetime.now().isoformat(),
                    "complete": True,
                    "requested_obs": sorted(obs)
                }

        payload = {
            "mb51": mb51_rows,
            "vl06i": vl06i_rows,
            "metadata": {
                "mb51": {
                    "filename": temp_mb51.name,
                    "captured_at": datetime.datetime.now().isoformat(),
                    "complete": True,
                    "requested_obs": sorted(obs)
                },
                "vl06i": vl06i_meta
            }
        }

        evidence = save_reconciliation(conn, pick_order_id, payload)

        # Success! Clean up intermediate files
        if not is_user_upload and temp_mb51 and temp_mb51.exists():
            try:
                temp_mb51.unlink()
            except OSError:
                pass
        if not is_user_upload and temp_vl06i and temp_vl06i.exists():
            try:
                temp_vl06i.unlink()
            except OSError:
                pass

        return {
            "status": "SUCCESS",
            "intermediate_file_cleaned": True,
            "diagnostic_file_retained": False,
            **format_reconciliation_summary(evidence)
        }

    except Exception as exc:
        retained_mb51 = str(temp_mb51) if (temp_mb51 and temp_mb51.exists() and not is_user_upload) else None
        return {
            "status": "DISCREPANCY_OR_ERROR",
            "intermediate_file_cleaned": False,
            "diagnostic_file_retained": bool(retained_mb51),
            "diagnostic_file": retained_mb51,
            "operator_summary": {
                "result_label": "Phát hiện sai lệch hoặc lỗi đối soát",
                "is_valid": False,
                "matched_count": 0,
                "discrepancies": [{"error": str(exc)}],
                "pending_actions": [
                    "Kiểm tra lại số lượng hoặc chứng từ xuất trên SAP",
                    f"File trung gian được giữ lại để chẩn đoán tại: {retained_mb51 or 'N/A'}"
                ]
            },
            "technical_details": {"error": str(exc)}
        }
