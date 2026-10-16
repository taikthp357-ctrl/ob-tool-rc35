# -*- coding: utf-8 -*-
"""Config, Master Data, History & Health API routes for MA CONTROL."""

import json
import datetime
from typing import List, Dict, Any
from fastapi import APIRouter, HTTPException
from app.core.config import APP_TITLE, APP_VERSION, SAP_MODE
from app.services.sap_connection import probe_connection
from app.core.database import get_connection
from app.models.schemas import MaterialSchema, ModelMappingSchema, SettingsUpdateSchema
from app.services.sap_worker import sap_worker

router = APIRouter(prefix="/api", tags=["Config & History"])

@router.get("/health")
def health_check():
    connection = probe_connection()
    return {
        "status": "healthy",
        "app": APP_TITLE,
        "version": APP_VERSION,
        "sap_connection": connection["state"],
        "sap_mode": SAP_MODE,
        "sap_details": connection,
        "live_write_verified": False,
        "timestamp": datetime.datetime.now().isoformat()
    }

@router.get("/materials")
def list_materials():
    conn = get_connection()
    try:
        rows = conn.execute("SELECT * FROM materials ORDER BY material_code ASC").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()

@router.post("/materials")
def upsert_material(req: MaterialSchema):
    conn = get_connection()
    try:
        with conn:
            conn.execute("""
                INSERT INTO materials (material_code, description, model, shop, standard_pack, default_sloc, is_external, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(material_code) DO UPDATE SET
                    description = excluded.description,
                    model = excluded.model,
                    shop = excluded.shop,
                    standard_pack = excluded.standard_pack,
                    default_sloc = excluded.default_sloc,
                    is_external = excluded.is_external,
                    updated_at = CURRENT_TIMESTAMP
            """, (
                req.material_code.strip(), req.description, req.model,
                req.shop, req.standard_pack, req.default_sloc, req.is_external
            ))
        return {"status": "success", "material_code": req.material_code}
    finally:
        conn.close()

@router.get("/model-mappings")
def list_model_mappings():
    conn = get_connection()
    try:
        rows = conn.execute("SELECT * FROM model_mappings ORDER BY material_code ASC").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()

@router.post("/model-mappings")
def upsert_model_mapping(req: ModelMappingSchema):
    conn = get_connection()
    try:
        with conn:
            conn.execute("""
                INSERT INTO model_mappings (material_code, model_code, shop_code, note)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(material_code) DO UPDATE SET
                    model_code = excluded.model_code,
                    shop_code = excluded.shop_code,
                    note = excluded.note
            """, (req.material_code.strip(), req.model_code.strip(), req.shop_code, req.note))
        return {"status": "success", "material_code": req.material_code}
    finally:
        conn.close()

@router.get("/settings")
def get_all_settings():
    conn = get_connection()
    try:
        rows = conn.execute("SELECT * FROM settings ORDER BY key ASC").fetchall()
        out = {}
        for r in rows:
            try:
                out[r["key"]] = json.loads(r["value_json"])
            except Exception:
                out[r["key"]] = r["value_json"]
        return out
    finally:
        conn.close()

@router.post("/settings")
def update_setting(req: SettingsUpdateSchema):
    conn = get_connection()
    try:
        val_str = json.dumps(req.value, ensure_ascii=False)
        with conn:
            conn.execute("""
                INSERT INTO settings (key, value_json, version, updated_at)
                VALUES (?, ?, 1, CURRENT_TIMESTAMP)
                ON CONFLICT(key) DO UPDATE SET
                    value_json = excluded.value_json,
                    version = settings.version + 1,
                    updated_at = CURRENT_TIMESTAMP
            """, (req.key, val_str))
        return {"status": "success", "key": req.key}
    finally:
        conn.close()

@router.get("/jobs/events")
def get_job_events(limit: int = 100):
    return sap_worker.list_job_events(limit=limit)
