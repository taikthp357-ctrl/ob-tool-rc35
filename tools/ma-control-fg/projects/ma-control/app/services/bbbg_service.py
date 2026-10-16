# -*- coding: utf-8 -*-
"""Step 5: evidence-backed local BBBG export, never simulated SAP side effects."""

import sqlite3
import uuid
import datetime
import json
import os
from pathlib import Path
from app.services.bbbg_export import export_bbbg, sha256_file, PROJECT_ROOT
from typing import List, Dict, Any, Optional, Union
from app.core.database import get_connection

class BBBGService:
    def __init__(self, conn: Optional[sqlite3.Connection] = None, export_dir=None):
        self._conn = conn
        self.export_dir = Path(export_dir or os.environ.get('MA_CONTROL_EXPORT_DIR') or PROJECT_ROOT / 'data/bbbg_exports')

    def _get_conn(self) -> sqlite3.Connection:
        return self._conn or get_connection()

    def resolve_model(self, material_code: str, default_model: Optional[str] = None) -> str:
        """Resolve model name using model_mappings table."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            mat = str(material_code).strip()
            row = conn.execute("SELECT model_code FROM model_mappings WHERE material_code = ?", (mat,)).fetchone()
            if row and row["model_code"]:
                return row["model_code"]

            # Prefix match for variants like BIW71001001AA -> BIW71001001
            if len(mat) >= 11:
                base_mat = mat[:11]
                prow = conn.execute("SELECT model_code FROM model_mappings WHERE material_code = ?", (base_mat,)).fetchone()
                if prow and prow["model_code"]:
                    return prow["model_code"]

            if default_model and str(default_model).strip().upper() not in ('UNKNOWN', 'N/A', '/'):
                return str(default_model).strip()
            raise ValueError(f'Chưa có mapping model cho mã hàng {mat}; cần bổ sung dữ liệu')
        finally:
            if close_conn:
                conn.close()

    def create_bbbg(
        self,
        pick_order_ref: Union[int, str],
        sender_name: str,
        receiver_name: Optional[str] = None,
        storage_bin: Optional[str] = None,
        slip_no: Optional[str] = None,
        doc_no: Optional[str] = None,
        is_external: int = 0,
        template_mode: str = "Auto",
        font_size: int = 14,
        existing_vat_attachments: Optional[List[str]] = None
    ) -> Dict[str, Any]:
        """Step 5: Generate BBBG Record.

        CRITICAL INVARIANT: Fails with ValueError if pick order status is not 'Đã CF' (or 'USER_CONFIRMED').
        """
        conn = self._get_conn()
        close_conn = self._conn is None
        exported_path = None
        try:
            if not sender_name or not sender_name.strip():
                raise ValueError('Cần tên người bàn giao')
            if template_mode not in ('Auto', 'Set') or not 8 <= font_size <= 24:
                raise ValueError('Chế độ mẫu hoặc cỡ chữ không hợp lệ')
            today_str = datetime.date.today().strftime("%Y%m%d")
            today_iso = datetime.date.today().isoformat()
            now_iso = datetime.datetime.now().isoformat()
            year_str = datetime.date.today().strftime("%Y")

            with conn:
                if not conn.in_transaction:
                    conn.execute('BEGIN IMMEDIATE')
                # 1. Fetch Pick Order
                if isinstance(pick_order_ref, int) or (isinstance(pick_order_ref, str) and pick_order_ref.isdigit()):
                    po = conn.execute("SELECT * FROM pick_orders WHERE id = ?", (int(pick_order_ref),)).fetchone()
                else:
                    po = conn.execute(
                        "SELECT * FROM pick_orders WHERE pick_order_code = ? OR pick_order_no = ?",
                        (str(pick_order_ref), str(pick_order_ref))
                    ).fetchone()

                if not po:
                    raise LookupError(f"Không tìm thấy đơn pick: {pick_order_ref}")

                po_id = po["id"]
                po_code = po["pick_order_code"] or po["pick_order_no"]
                prior = conn.execute('SELECT * FROM bbbg_records WHERE pick_order_id = ? ORDER BY id LIMIT 1', (po_id,)).fetchone()
                if prior:
                    if (slip_no and prior['slip_no'] != slip_no) or (doc_no and prior['doc_no'] != doc_no):
                        raise ValueError('Đơn pick đã có BBBG; không tạo trùng')
                    if prior['status'] == 'DRAFT_CREATED':
                        from app.services.reconciliation_service import get_verified_reconciliation
                        verified_source = get_verified_reconciliation(conn, po_id)
                        if verified_source:
                            final_sender = sender_name.strip() if sender_name else (prior['sender_name'] or '').strip()
                            final_receiver = (receiver_name if receiver_name is not None else prior['receiver_name']) or ''
                            final_receiver = final_receiver.strip()
                            final_storage = storage_bin if storage_bin is not None else prior['storage_bin']
                            if not final_sender or not final_receiver:
                                raise ValueError('BBBG đã đối chiếu cần tên người bàn giao và người nhận thực tế')
                            exported_path = self.export_dir / (prior['bbbg_uuid'] + '-verified-' + uuid.uuid4().hex[:8] + '.xlsx')
                            export_evidence = export_bbbg(exported_path, {'slip_no': prior['slip_no'], 'doc_no': prior['doc_no'], 'sender_name': final_sender, 'receiver_name': final_receiver, 'shop': prior['shop'], 'storage_bin': final_storage, 'handover_date': prior['handover_date'], 'reconciliation': verified_source}, [{**i, 'cf_actual_qty': i['quantity']} for i in verified_source['rows']])
                            qty = sum(float(i['quantity']) for i in verified_source['rows'])
                            conn.execute("UPDATE bbbg_records SET status='VERIFIED_IMPORTED', total_qty=?,sender_name=?,handover_person=?,receiver_name=?,receiver_person=?,storage_bin=?,storage_location_note=? WHERE id=?", (qty, final_sender, final_sender, final_receiver, final_receiver, final_storage, final_storage, prior['id']))
                            conn.execute("INSERT INTO job_events (job_id,business_key,job_type,step_name,event_type,target_ref,status,evidence_json,created_at) VALUES (?,?, 'BBBG','FINALIZE_DRAFT','FINALIZE_DRAFT',?,'SUCCESS',?,?)", (uuid.uuid4().hex, prior['dedup_key'], str(prior['id']), json.dumps(export_evidence, ensure_ascii=False), now_iso))
                            return {**dict(prior), **export_evidence, 'status':'success', 'record_status':'VERIFIED_IMPORTED', 'total_qty':qty, 'sender_name':final_sender,'receiver_name':final_receiver,'storage_bin':final_storage, 'finalized':True}
                    return {**dict(prior), **self._export_evidence(conn, prior['id']), 'status': 'success', 'idempotent': True}

                # 2. STRICT INVARIANT CHECK
                if po["status"] not in ("Đã CF", "USER_CONFIRMED"):
                    raise ValueError(
                        f"BẤT BIẾN NGHIỆP VỤ BỊ VI PHẠM: Đơn pick {po_code} đang ở trạng thái '{po['status']}', "
                        f"chưa được người dùng CF thủ công! Không thể tạo BBBG."
                    )

                # 3. Fetch Items & resolve model
                items = conn.execute("SELECT * FROM pick_order_items WHERE pick_order_id = ?", (po_id,)).fetchall()
                if not items:
                    raise ValueError('Đơn pick không có dòng hàng')
                total_qty = sum(float(i['cf_actual_qty'] or 0) for i in items)
                reconciliation = None
                if conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='source_reconciliations'").fetchone():
                    from app.services.reconciliation_service import get_verified_reconciliation
                    reconciliation = get_verified_reconciliation(conn, po_id)
                export_items = [dict(i) for i in items]
                if reconciliation:
                    if not receiver_name or not receiver_name.strip():
                        raise ValueError('BBBG đã đối chiếu cần tên người nhận thực tế')
                    export_items = [{**i, 'cf_actual_qty': i['quantity']} for i in reconciliation['rows']]
                    total_qty = sum(float(i['cf_actual_qty']) for i in export_items)
                resolved_model = po["model"]
                if items:
                    first_mat = items[0]["material_code"]
                    resolved_model = self.resolve_model(first_mat, default_model=po["model"])

                # 4. Generate Slip No & Dedup Key
                if not slip_no:
                    cnt = conn.execute(
                        "SELECT COUNT(*) FROM bbbg_records WHERE handover_date = ?",
                        (today_iso,)
                    ).fetchone()[0] + 1
                    generated_slip = f"BBBG-{today_str}-{cnt:03d}"
                else:
                    generated_slip = slip_no

                plant_shop = f"{po['branch_tab'] or '3001'}_{po['shop']}"
                doc_type = "FG_HANDOVER"
                dedup_key = f"{doc_type}_{year_str}_{plant_shop}_{generated_slip}"

                # Check append-only uniqueness
                existing = conn.execute("SELECT 1 FROM bbbg_records WHERE dedup_key = ?", (dedup_key,)).fetchone()
                if existing:
                    raise ValueError(f"Khóa sổ tổng đã tồn tại (chống trùng lặp append-only): {dedup_key}")

                # 5. GOS Attachment Logic (avoid duplicate VAT)
                vat_filename = f"VAT_{po_code}_{year_str}.pdf"
                gos_vat_attached = 0
                if existing_vat_attachments and vat_filename in existing_vat_attachments:
                    # Skip duplicate VAT
                    gos_vat_attached = 0

                bbbg_uuid = f"bbbg-{uuid.uuid4().hex}"
                doc_code = doc_no or f"DOC-{today_str}-{po_id}"
                ob_no = items[0]["ob_number"] if items and items[0]["ob_number"] else None
                exported_path = self.export_dir / (bbbg_uuid + '.xlsx')
                header = {'slip_no': generated_slip, 'doc_no': doc_code, 'sender_name': sender_name, 'receiver_name': receiver_name, 'shop': po['shop'], 'storage_bin': storage_bin, 'handover_date': today_iso, 'template_mode': template_mode, 'font_size': font_size}
                if reconciliation:
                    header['reconciliation'] = reconciliation

                # 1. Database First Persistence: Commit bbbg_records to SQLite before writing physical file
                cur = conn.execute("""
                    INSERT INTO bbbg_records (
                        bbbg_uuid, bbbg_no, dedup_key, pick_order_id, pick_order_no,
                        slip_no, doc_no, ob_number, shop, model, total_qty,
                        handover_person, receiver_person, sender_name, receiver_name,
                        storage_location_note, storage_bin, is_external, template_mode,
                        font_size, email_status, print_status, unc_export_status,
                        general_log_saved, shop_log_saved, gos_bbbg_attached, gos_vat_attached,
                        status, handover_date, created_at
                    ) VALUES (
                        ?, ?, ?, ?, ?,
                        ?, ?, ?, ?, ?, ?,
                        ?, ?, ?, ?,
                        ?, ?, ?, ?,
                        ?, 'PENDING', 'PENDING', 'PENDING',
                        0, 0, 0, ?,
                        'DRAFT_CREATED', ?, ?
                    )
                """, (
                    bbbg_uuid, generated_slip, dedup_key, po_id, po_code,
                    generated_slip, doc_code, ob_no, po["shop"], resolved_model, total_qty,
                    sender_name, receiver_name, sender_name, receiver_name,
                    storage_bin, storage_bin, is_external, template_mode,
                    font_size, gos_vat_attached,
                    today_iso, now_iso
                ))
                bbbg_id = cur.lastrowid
                if reconciliation:
                    conn.execute("UPDATE bbbg_records SET status='VERIFIED_IMPORTED' WHERE id=?", (bbbg_id,))

                # 2. Export physical file with intelligent File Lock catching (PermissionError / WinError 32)
                is_file_locked = False
                export_evidence = {}
                try:
                    export_evidence = export_bbbg(exported_path, header, export_items)
                except (PermissionError, OSError) as pe:
                    err_str = str(pe)
                    is_lock = (
                        getattr(pe, "winerror", None) == 32
                        or isinstance(pe, PermissionError)
                        or "used by another process" in err_str
                        or "Permission denied" in err_str
                    )
                    if is_lock:
                        is_file_locked = True
                        conn.execute("UPDATE bbbg_records SET status='FILE_LOCKED' WHERE id=?", (bbbg_id,))
                        conn.execute("""
                            INSERT INTO job_events (
                                job_id, business_key, job_type, step_name, event_type,
                                target_ref, status, evidence, evidence_json, created_at
                            ) VALUES (?, ?, 'BBBG', 'GENERATE_BBBG', 'FILE_LOCKED', ?, 'WARNING', ?, ?, ?)
                        """, (
                            str(uuid.uuid4()), f"BBBG_{generated_slip}", str(bbbg_id),
                            "Tệp Excel đang được mở bởi ứng dụng khác (Microsoft Excel). Quá trình ghi bị từ chối.",
                            json.dumps({"error": err_str, "slip_no": generated_slip, "saved_in_db": True}, ensure_ascii=False),
                            now_iso
                        ))
                    else:
                        raise

                if is_file_locked:
                    return {
                        "status": "FILE_LOCKED",
                        "is_file_locked": True,
                        "id": bbbg_id,
                        "record_id": bbbg_id,
                        "bbbg_uuid": bbbg_uuid,
                        "bbbg_no": generated_slip,
                        "slip_no": generated_slip,
                        "dedup_key": dedup_key,
                        "pick_order_id": po_id,
                        "pick_order_code": po_code,
                        "total_qty": total_qty,
                        "saved_in_db": True,
                        "error_title": "Khóa tệp (Permission Denied)",
                        "reason": "Tệp Excel đang được mở bởi ứng dụng khác (Microsoft Excel). Quá trình ghi bị từ chối.",
                        "db_status": "Toàn bộ dữ liệu bàn giao đã được lưu trữ an toàn trong cơ sở dữ liệu SQLite.",
                        "action_advice": "Vui lòng đóng tệp Excel trên máy và nhấn nút 'Thử lại bước này'."
                    }

                # Audit log for successful export
                conn.execute("""
                    INSERT INTO job_events (
                        job_id, business_key, job_type, step_name, event_type,
                        target_ref, status, evidence, evidence_json, created_at
                    ) VALUES (?, ?, 'BBBG', 'GENERATE_BBBG', 'CREATE_BBBG', ?, 'SUCCESS', ?, ?, ?)
                """, (
                    str(uuid.uuid4()), f"BBBG_{generated_slip}", str(bbbg_id),
                    f"Local draft BBBG {generated_slip}; SAP not verified; qty={total_qty}",
                    json.dumps(export_evidence, ensure_ascii=False), now_iso
                ))

            return {
                "status": "success",
                "id": bbbg_id,
                "bbbg_uuid": bbbg_uuid,
                "bbbg_no": generated_slip,
                "slip_no": generated_slip,
                "dedup_key": dedup_key,
                "pick_order_id": po_id,
                "pick_order_code": po_code,
                "model": resolved_model,
                "total_qty": total_qty,
                "gos_bbbg_attached": 0,
                "gos_vat_attached": gos_vat_attached,
                "general_log_saved": 0,
                "shop_log_saved": 0,
                **export_evidence,
                "message": ('Đã xuất BBBG từ nguồn nhập đã đối chiếu; chưa xác nhận SAP/GR' if reconciliation else 'Đã xuất bản nháp BBBG; cần đối chiếu MB51/VL06I trước bàn giao')
            }
        except Exception:
            if exported_path is not None:
                exported_path.unlink(missing_ok=True)
            raise
        finally:
            if close_conn:
                conn.close()

    def _export_evidence(self, conn, record_id):
        row = conn.execute("SELECT evidence_json FROM job_events WHERE job_type = 'BBBG' AND target_ref = ? ORDER BY id DESC LIMIT 1", (str(record_id),)).fetchone()
        if not row or not row['evidence_json']:
            raise ValueError('BBBG cũ chưa có bằng chứng tệp xuất')
        evidence = json.loads(row['evidence_json'])
        path = Path(evidence['file_path']).resolve()
        if not path.is_relative_to(self.export_dir.resolve()) or not path.is_file() or sha256_file(path) != evidence['file_sha256']:
            raise ValueError('Tệp BBBG bị mất hoặc thay đổi; cần kiểm tra phục hồi')
        return evidence

    def get_export_path(self, record_id):
        conn = self._get_conn()
        try:
            return self._export_evidence(conn, record_id)['file_path']
        finally:
            if self._conn is None:
                conn.close()

    def list_bbbg_records(self) -> List[Dict[str, Any]]:
        """List all BBBG records."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            rows = conn.execute("SELECT * FROM bbbg_records ORDER BY id DESC").fetchall()
            return [dict(r) for r in rows]
        finally:
            if close_conn:
                conn.close()
