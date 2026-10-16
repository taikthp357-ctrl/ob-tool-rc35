# -*- coding: utf-8 -*-
"""Database initialization and SQLite manager for MA CONTROL.

Strictly enforces WAL mode, foreign keys, and TEXT preservation of leading zeroes
for material codes, batches, STOs, OB numbers, and slip codes.
Implements the full 10-table schema per ARCHITECTURE_SPEC_AIM1167.
"""

import sqlite3
import datetime
from pathlib import Path
from typing import Optional
from app.core.config import DB_PATH

def get_connection(db_file: Optional[Path] = None) -> sqlite3.Connection:
    """Get a SQLite connection configured with WAL mode and row factory."""
    target_path = str(db_file or DB_PATH)
    conn = sqlite3.connect(target_path, timeout=30.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA foreign_keys=ON;")
    conn.execute("PRAGMA busy_timeout=5000;")
    return conn

def init_db(db_file: Optional[Path] = None):
    """Initialize SQLite database with all 10 tables and seed data."""
    conn = get_connection(db_file)
    with conn:
        conn.executescript("""
        -- 1. materials
        CREATE TABLE IF NOT EXISTS materials (
            material_code TEXT PRIMARY KEY,
            description TEXT,
            model TEXT,
            shop TEXT,
            standard_pack REAL DEFAULT 1,
            us_rate INTEGER DEFAULT 1,
            std_pack INTEGER DEFAULT 1,
            default_sloc TEXT,
            is_external INTEGER DEFAULT 0,
            updated_at TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );

        -- 2. model_mappings
        CREATE TABLE IF NOT EXISTS model_mappings (
            material_code TEXT PRIMARY KEY,
            model_code TEXT NOT NULL,
            shop_code TEXT,
            note TEXT
        );

        -- 3. stock_snapshots
        CREATE TABLE IF NOT EXISTS stock_snapshots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_type TEXT,
            snapshot_time TEXT,
            plant TEXT,
            sloc TEXT,
            total_rows INTEGER DEFAULT 0,
            total_materials INTEGER DEFAULT 0,
            total_qty REAL DEFAULT 0,
            notes TEXT
        );

        -- 4. stock_items
        CREATE TABLE IF NOT EXISTS stock_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            snapshot_id INTEGER,
            plant TEXT,
            sloc TEXT,
            material_code TEXT,
            batch TEXT,
            stock_type TEXT DEFAULT 'UNRESTRICTED',
            unrestricted_qty REAL DEFAULT 0,
            qty REAL DEFAULT 0,
            unit TEXT DEFAULT 'PCS',
            priority_rank INTEGER DEFAULT 0,
            priority_flag TEXT,
            pull_line_warning TEXT,
            FOREIGN KEY (snapshot_id) REFERENCES stock_snapshots(id) ON DELETE CASCADE
        );

        -- 5. ob_items
        CREATE TABLE IF NOT EXISTS ob_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_uid TEXT UNIQUE,
            branch_tab TEXT NOT NULL,
            branch TEXT,
            row_order INTEGER DEFAULT 0,
            selected INTEGER DEFAULT 1,
            material_code TEXT,
            model TEXT,
            req_qty REAL DEFAULT 0,
            required_qty REAL DEFAULT 0,
            avail_qty REAL DEFAULT 0,
            missing_qty REAL DEFAULT 0,
            batch TEXT,
            shop TEXT,
            sto_number TEXT,
            ob_number TEXT,
            status TEXT DEFAULT 'Chưa kiểm tra',
            rc52_code TEXT,
            note TEXT,
            error_message TEXT,
            pick_order_id INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        );

        -- 6. pick_orders
        CREATE TABLE IF NOT EXISTS pick_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pick_order_code TEXT UNIQUE,
            pick_order_no TEXT,
            branch_tab TEXT,
            shop TEXT,
            model TEXT,
            item_count INTEGER DEFAULT 0,
            total_items INTEGER DEFAULT 0,
            total_qty REAL DEFAULT 0,
            status TEXT DEFAULT 'Chờ người dùng CF',
            cf_verified_by TEXT,
            cf_verified_at TEXT,
            confirmed_by TEXT,
            confirmed_at TEXT,
            notes TEXT,
            created_date TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        );

        -- 7. pick_order_items
        CREATE TABLE IF NOT EXISTS pick_order_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pick_order_id INTEGER,
            pick_order_no TEXT,
            ob_item_id INTEGER,
            material_code TEXT,
            model TEXT,
            batch TEXT,
            shop TEXT,
            sto_number TEXT,
            ob_number TEXT,
            plant TEXT,
            packing_box TEXT,
            rack TEXT,
            is_direct_paste INTEGER DEFAULT 0,
            direct_order_ref TEXT,
            dedup_key TEXT,
            qty REAL DEFAULT 0,
            cf_status TEXT DEFAULT 'PENDING_CF',
            cf_actual_qty REAL DEFAULT 0,
            line_confirmed INTEGER DEFAULT 0,
            missing_qty REAL DEFAULT 0,
            shortage_reason TEXT,
            verified_by TEXT,
            verified_at TEXT,
            evidence_note TEXT,
            cf_source TEXT DEFAULT 'MANUAL_ACTUAL',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (pick_order_id) REFERENCES pick_orders(id) ON DELETE CASCADE
        );

        -- 8. bbbg_records
        CREATE TABLE IF NOT EXISTS bbbg_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bbbg_uuid TEXT UNIQUE,
            bbbg_no TEXT,
            dedup_key TEXT UNIQUE,
            pick_order_id INTEGER,
            pick_order_no TEXT,
            slip_no TEXT,
            doc_no TEXT,
            ob_number TEXT,
            shop TEXT,
            model TEXT,
            total_qty REAL DEFAULT 0,
            handover_person TEXT,
            receiver_person TEXT,
            sender_name TEXT,
            receiver_name TEXT,
            storage_location_note TEXT,
            storage_bin TEXT,
            is_external INTEGER DEFAULT 0,
            template_mode TEXT DEFAULT 'Auto',
            font_size INTEGER DEFAULT 14,
            email_status TEXT DEFAULT 'PENDING',
            print_status TEXT DEFAULT 'PENDING',
            unc_export_status TEXT DEFAULT 'PENDING',
            general_log_saved INTEGER DEFAULT 0,
            shop_log_saved INTEGER DEFAULT 0,
            gos_bbbg_attached INTEGER DEFAULT 0,
            gos_vat_attached INTEGER DEFAULT 0,
            status TEXT DEFAULT 'CREATED',
            handover_date TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );

        -- 9. job_events
        CREATE TABLE IF NOT EXISTS job_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id TEXT,
            business_key TEXT,
            job_type TEXT,
            step_name TEXT,
            event_type TEXT,
            target_ref TEXT,
            status TEXT,
            duration_ms INTEGER DEFAULT 0,
            sap_status_msg TEXT,
            evidence_json TEXT,
            evidence TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );

        -- 10. settings
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value_json TEXT,
            version INTEGER DEFAULT 1,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        );

        -- 11. stock_thresholds (Quản lý ngưỡng cảnh báo tồn thấp theo mã hàng & plant)
        CREATE TABLE IF NOT EXISTS stock_thresholds (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            material_code TEXT,
            plant TEXT,
            threshold_qty REAL NOT NULL DEFAULT 50.0,
            warning_message TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(material_code, plant)
        );

        -- 12. low_stock_allocations (Ghi nhận / Đề xuất phân bổ tồn thấp 2-3 batch tùy chọn)
        CREATE TABLE IF NOT EXISTS low_stock_allocations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            allocation_code TEXT UNIQUE,
            material_code TEXT NOT NULL,
            plant TEXT NOT NULL,
            shop TEXT,
            scope_sloc TEXT,
            required_qty REAL NOT NULL,
            total_allocated_qty REAL NOT NULL,
            status TEXT DEFAULT 'Ghi nhận / Đề xuất',
            result_type TEXT DEFAULT 'PROPOSAL',
            sap_live_status TEXT DEFAULT 'NOT_APPLIED_SAP',
            other_shop_impact TEXT DEFAULT 'Chưa đánh giá được nhu cầu/dự trữ shop khác',
            allocations_json TEXT NOT NULL,
            evidence_json TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """)

        conn.executescript("""
        CREATE TABLE IF NOT EXISTS stock_snapshot_scopes (
            snapshot_id INTEGER NOT NULL REFERENCES stock_snapshots(id) ON DELETE CASCADE,
            plant TEXT NOT NULL,
            sloc TEXT NOT NULL,
            PRIMARY KEY (snapshot_id, plant, sloc)
        );
        INSERT OR IGNORE INTO stock_snapshot_scopes(snapshot_id, plant, sloc)
            SELECT DISTINCT snapshot_id, plant, sloc FROM stock_items
            WHERE snapshot_id IS NOT NULL AND plant IS NOT NULL AND sloc IS NOT NULL;
        CREATE TABLE IF NOT EXISTS source_reconciliations (
            pick_order_id INTEGER PRIMARY KEY REFERENCES pick_orders(id),
            source_hash TEXT NOT NULL,
            status TEXT NOT NULL,
            evidence_json TEXT NOT NULL,
            checked_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS pick_source_uploads (
            pick_order_id INTEGER NOT NULL REFERENCES pick_orders(id),
            kind TEXT NOT NULL,
            source_json TEXT NOT NULL,
            PRIMARY KEY(pick_order_id, kind)
        );
        CREATE TABLE IF NOT EXISTS cf_check_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pick_order_id INTEGER NOT NULL,
            pick_order_item_id INTEGER,
            ob_number TEXT,
            material_code TEXT,
            batch TEXT,
            req_qty REAL DEFAULT 0,
            actual_qty REAL DEFAULT 0,
            missing_qty REAL DEFAULT 0,
            reason TEXT,
            status TEXT,
            verified_by TEXT NOT NULL,
            verified_at TEXT NOT NULL,
            evidence TEXT,
            cf_source TEXT DEFAULT 'MANUAL_ACTUAL',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """)

        poi_cols = {r[1] for r in conn.execute("PRAGMA table_info(pick_order_items)").fetchall()}
        for col_name, col_def in [
            ("missing_qty", "REAL DEFAULT 0"),
            ("shortage_reason", "TEXT"),
            ("verified_by", "TEXT"),
            ("verified_at", "TEXT"),
            ("evidence_note", "TEXT"),
            ("cf_source", "TEXT DEFAULT 'MANUAL_ACTUAL'"),
            ("plant", "TEXT"),
            ("packing_box", "TEXT"),
            ("rack", "TEXT"),
            ("is_direct_paste", "INTEGER DEFAULT 0"),
            ("direct_order_ref", "TEXT"),
            ("dedup_key", "TEXT"),
            ("created_at", "TEXT"),
            ("updated_at", "TEXT"),
        ]:
            if col_name not in poi_cols:
                conn.execute(f"ALTER TABLE pick_order_items ADD COLUMN {col_name} {col_def}")
        conn.execute("CREATE INDEX IF NOT EXISTS ix_poi_direct ON pick_order_items(is_direct_paste)")
        conn.execute("CREATE INDEX IF NOT EXISTS ix_poi_dedup ON pick_order_items(dedup_key)")

        pick_columns = {r[1] for r in conn.execute("PRAGMA table_info(pick_orders)").fetchall()}
        if "source_key" not in pick_columns:
            conn.execute("ALTER TABLE pick_orders ADD COLUMN source_key TEXT")
        conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_pick_source_key ON pick_orders(source_key)")

        # Seed model mappings if empty
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM model_mappings")
        if cursor.fetchone()[0] == 0:
            seed_mappings = [
                ("BIW71001001", "VF8 LIMO", "BODY", "Mapping chuẩn LIMO V2.2.14"),
                ("BIW71001001AA", "VF8 LIMO", "BODY", "Mapping chuẩn LIMO AA"),
                ("BIW70002022", "VF7", "BODY", "Mapping chuẩn VF7 V2.2.14"),
                ("BIW70002022BA", "VF7", "BODY", "Mapping chuẩn VF7 BA"),
                ("BIW73012711CA", "VF5 RHD", "BODY", "Mapping VF5 RHD"),
                ("BIW73012036CA", "VF5 RHD", "BODY", "Mapping VF5 RHD"),
                ("EDC88001122AA", "VF9", "EDC", "Mapping EDC VF9"),
                ("EDC88001199BB", "VF9", "EDC", "Mapping EDC VF9"),
            ]
            baseline_file = Path(__file__).parent / "baseline_materials.json"
            if baseline_file.exists():
                try:
                    import json
                    with open(baseline_file, "r", encoding="utf-8") as f:
                        b_items = json.load(f)
                    for item in b_items:
                        seed_mappings.append((item["material_code"], item["model"], "BODY", "Baseline V3.6 DATA"))
                except Exception:
                    pass
            cursor.executemany("""
                INSERT OR IGNORE INTO model_mappings (material_code, model_code, shop_code, note)
                VALUES (?, ?, ?, ?)
            """, seed_mappings)

        # Seed materials if empty
        cursor.execute("SELECT COUNT(*) FROM materials")
        if cursor.fetchone()[0] == 0:
            seed_materials = [
                ("BIW73012711CA", "Cột B trái VF5", "VF5 RHD", "BODY", 20.0, 1, 20, "3000", 0),
                ("BIW73012036CA", "Khung sàn trước VF5", "VF5 RHD", "BODY", 10.0, 1, 10, "3000", 0),
                ("BIW71001001AA", "Cản sau VF8 LIMO", "VF8 LIMO", "BODY", 15.0, 1, 15, "3000", 0),
                ("BIW70002022BA", "Cánh cửa VF7", "VF7", "BODY", 8.0, 1, 8, "3000", 0),
                ("EDC88001122AA", "Cụm dây điện EDC VF9", "VF9", "EDC", 5.0, 1, 5, "4004", 0),
                ("EDC88001199BB", "Cụm giắc nối EDC VF9", "VF9", "EDC", 5.0, 1, 5, "4004", 0),
            ]
            baseline_file = Path(__file__).parent / "baseline_materials.json"
            if baseline_file.exists():
                try:
                    import json
                    with open(baseline_file, "r", encoding="utf-8") as f:
                        b_items = json.load(f)
                    for item in b_items:
                        seed_materials.append((
                            item["material_code"],
                            f"{item['material_code']} ({item['model']})",
                            item["model"],
                            "BODY",
                            1.0,
                            item.get("us_rate", 1),
                            item.get("us_rate", 1),
                            "1003",
                            0
                        ))
                except Exception:
                    pass
            cursor.executemany("""
                INSERT OR IGNORE INTO materials (material_code, description, model, shop, standard_pack, us_rate, std_pack, default_sloc, is_external)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, seed_materials)

        # Seed settings if empty
        cursor.execute("SELECT COUNT(*) FROM settings")
        if cursor.fetchone()[0] == 0:
            seed_settings = [
                ("sap_connection", '{"status": "NOT_VERIFIED"}', 1),
                ("email_recipients", '{"BODY": "gl_body@vinfast.vn", "EDC": "gl_edc@vinfast.vn"}', 1),
                ("unc_export_path", '{"path": "\\\\vinfast\\shares\\ma_control\\export"}', 1),
                ("template_defaults", '{"font_size": 14, "mode": "Auto"}', 1)
            ]
            cursor.executemany("""
                INSERT OR IGNORE INTO settings (key, value_json, version)
                VALUES (?, ?, ?)
            """, seed_settings)

        # Synthetic operational rows belong only in explicit test mode.
        cursor.execute("SELECT COUNT(*) FROM ob_items")
        import os
        if cursor.fetchone()[0] == 0 and os.environ.get("MA_CONTROL_SAP_MODE", "MANUAL").upper() == "TEST":
            now = datetime.datetime.now().isoformat()
            demo_ob_items = [
                (1, "item-1", "BODY_3001", "BODY_3001", 1, 1, "BIW73012711CA", "VF5 RHD", 486.0, 486.0, 486.0, 0.0, "26101301", "BODY", "", "1012646320", "Sẵn sàng điền", "RC52_READY", None, None, None, now, now),
                (2, "item-2", "BODY_3001", "BODY_3001", 2, 1, "BIW73012036CA", "VF5 RHD", 833.0, 833.0, 833.0, 0.0, "26101301", "BODY", "", "1012646396", "Chưa kiểm tra", "RC52_PENDING", None, None, None, now, now),
                (3, "item-3", "BODY_3001", "BODY_3001", 3, 1, "BIW71001001AA", "VF8 LIMO", 120.0, 120.0, 0.0, 120.0, "26101302", "BODY", "", "", "Thiếu SL OB", "RC52_INSUFFICIENT", "Thiếu tồn kho 3001", "Thiếu tồn kho 3001", None, now, now),
                (4, "item-4", "BODY_3001", "BODY_3001", 4, 1, "BIW70002022BA", "VF7", 250.0, 250.0, 250.0, 0.0, "26101302", "BODY", "", "1012646400", "Chờ pick / CF", "RC52_WAIT_PICK", None, None, None, now, now),
                (5, "item-5", "EDC_4007", "EDC_4007", 1, 1, "EDC88001122AA", "VF9", 45.0, 45.0, 45.0, 0.0, "26101303", "EDC", "4500123890", "1012647100", "Sẵn sàng điền", "RC52_READY", None, None, None, now, now),
                (6, "item-6", "EDC_4007", "EDC_4007", 2, 1, "EDC88001199BB", "VF9", 90.0, 90.0, 90.0, 0.0, "26101303", "EDC", "4500123891", "", "Chưa kiểm tra", "RC52_PENDING", None, None, None, now, now),
            ]
            cursor.executemany("""
                INSERT OR IGNORE INTO ob_items (
                    id, item_uid, branch_tab, branch, row_order, selected, material_code, model,
                    req_qty, required_qty, avail_qty, missing_qty, batch, shop, sto_number, ob_number,
                    status, rc52_code, note, error_message, pick_order_id, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, demo_ob_items)

        # Seed demo pick order if empty
        cursor.execute("SELECT COUNT(*) FROM pick_orders")
        if cursor.fetchone()[0] == 0 and os.environ.get("MA_CONTROL_SAP_MODE", "MANUAL").upper() == "TEST":
            cursor.execute("""
                INSERT OR IGNORE INTO pick_orders (
                    id, pick_order_code, pick_order_no, branch_tab, shop, model,
                    item_count, total_items, total_qty, status, notes, created_date
                ) VALUES (
                    1, 'PICK-20261013-001', 'PICK-20261013-001', 'BODY_3001', 'BODY', 'VF5 RHD',
                    2, 2, 1319.0, 'Chờ người dùng CF', 'Chờ kiểm tra thực tế tại line Body', '2026-10-13'
                )
            """)
            cursor.execute("""
                INSERT OR IGNORE INTO pick_order_items (
                    pick_order_id, pick_order_no, ob_item_id, material_code, model, batch, shop, ob_number, qty, cf_status
                ) VALUES
                (1, 'PICK-20261013-001', 1, 'BIW73012711CA', 'VF5 RHD', '26101301', 'BODY', '1012646320', 486.0, 'PENDING_CF'),
                (1, 'PICK-20261013-001', 2, 'BIW73012036CA', 'VF5 RHD', '26101301', 'BODY', '1012646396', 833.0, 'PENDING_CF')
            """)

    conn.close()

if __name__ == "__main__":
    init_db()
    print(f"[*] Initialized SQLite database successfully at: {DB_PATH}")
