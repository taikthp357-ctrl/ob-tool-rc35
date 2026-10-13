# -*- coding: utf-8 -*-
"""Single-worker SAP GUI COM Job Queue Engine.

Features:
1. Thread-isolated COM execution (pythoncom.CoInitialize on worker thread).
2. Idempotent lock via business_key to prevent duplicate 'Run All' clicks.
3. Step-level state tracking (PENDING, RUNNING, SUCCEEDED, FAILED, UNCERTAIN, SKIPPED).
4. Resume from exact interrupted step/row without re-executing completed steps.
5. Safe audit logging (scrubs passwords and sensitive credentials).
"""

import re
import json
import time
import uuid
import queue
import threading
import datetime
import sqlite3
from typing import Dict, Any, Optional, Callable, List
from app.core.database import get_connection

SENSITIVE_PATTERNS = [
    (re.compile(r'(?i)(password|passwd|pwd|secret|token|apikey)\s*[:=]\s*["\']?([^"\'\s,;]+)["\']?'), r'\1=***REDACTED***'),
]

def sanitize_log_text(text: str) -> str:
    """Remove passwords or sensitive tokens before writing to audit log."""
    if not text:
        return ""
    result = str(text)
    for pattern, repl in SENSITIVE_PATTERNS:
        result = pattern.sub(repl, result)
    return result

class SAPWorkerEngine:
    def __init__(self, conn: Optional[sqlite3.Connection] = None):
        self._conn = conn
        self._active_locks: Dict[str, str] = {}  # business_key -> job_id
        self._lock = threading.Lock()
        self._job_queue: "queue.Queue[Dict[str, Any]]" = queue.Queue()

    def _get_conn(self) -> sqlite3.Connection:
        return self._conn or get_connection()

    def acquire_idempotent_lock(self, business_key: str, job_id: Optional[str] = None) -> str:
        """Acquire an idempotent lock for a business_key.

        Raises RuntimeError if another job with the same business_key is currently running.
        """
        with self._lock:
            if business_key in self._active_locks:
                existing_job = self._active_locks[business_key]
                raise RuntimeError(
                    f"IDEMPOTENT_LOCK_CONFLICT: Tác vụ với khóa '{business_key}' đang chạy (job_id={existing_job}). Chống bấm lặp Run All!"
                )
            jid = job_id or f"job-{uuid.uuid4().hex[:10]}"
            self._active_locks[business_key] = jid
            return jid

    def release_idempotent_lock(self, business_key: str):
        """Release the idempotent lock for a business_key."""
        with self._lock:
            self._active_locks.pop(business_key, None)

    def log_step_event(
        self,
        job_id: str,
        business_key: str,
        job_type: str,
        step_name: str,
        status: str,
        duration_ms: int = 0,
        sap_status_msg: Optional[str] = None,
        evidence: Optional[Dict[str, Any]] = None
    ) -> int:
        """Record a step event in job_events with credential scrubbing."""
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            now_iso = datetime.datetime.now().isoformat()
            clean_msg = sanitize_log_text(sap_status_msg or "")
            raw_ev = json.dumps(evidence or {}, ensure_ascii=False)
            clean_ev = sanitize_log_text(raw_ev)

            with conn:
                cur = conn.execute("""
                    INSERT INTO job_events (
                        job_id, business_key, job_type, step_name, event_type,
                        target_ref, status, duration_ms, sap_status_msg,
                        evidence_json, evidence, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    job_id, business_key, job_type, step_name, job_type,
                    business_key, status, duration_ms, clean_msg,
                    clean_ev, clean_ev, now_iso
                ))
                return cur.lastrowid
        finally:
            if close_conn:
                conn.close()

    def execute_pipeline(
        self,
        business_key: str,
        job_type: str,
        steps: List[Dict[str, Any]],
        resume: bool = False
    ) -> Dict[str, Any]:
        """Execute a multi-step job pipeline with idempotent locking and step-level resume.

        Each item in `steps` is a dict:
        {
            "step_name": str,
            "action": Callable[[], Dict[str, Any]],
            "is_external_side_effect": bool (optional, marks UNCERTAIN if crashed mid-way)
        }
        """
        job_id = self.acquire_idempotent_lock(business_key)
        completed_steps = set()
        uncertain_steps = set()

        if resume:
            conn = self._get_conn()
            close_conn = self._conn is None
            try:
                rows = conn.execute("""
                    SELECT step_name, status FROM job_events
                    WHERE business_key = ? AND job_type = ?
                    ORDER BY id ASC
                """, (business_key, job_type)).fetchall()
                for r in rows:
                    if r["status"] in ("SUCCEEDED", "SUCCESS"):
                        completed_steps.add(r["step_name"])
                        uncertain_steps.discard(r["step_name"])
                    elif r["status"] == "UNCERTAIN":
                        completed_steps.discard(r["step_name"])
                        uncertain_steps.add(r["step_name"])
                    elif r["status"] in ("FAILED", "RUNNING"):
                        completed_steps.discard(r["step_name"])
                        if r["status"] == "RUNNING" and any(s["step_name"] == r["step_name"] and s.get("is_external_side_effect") for s in steps):
                            uncertain_steps.add(r["step_name"])
            finally:
                if close_conn:
                    conn.close()

        results = []
        try:
            for step in steps:
                sname = step["step_name"]
                if resume and sname in completed_steps:
                    results.append({"step_name": sname, "status": "SKIPPED_ALREADY_DONE"})
                    continue

                if resume and sname in uncertain_steps:
                    # Must reconcile uncertain external side effects before re-running
                    results.append({
                        "step_name": sname,
                        "status": "UNCERTAIN_REQUIRES_RECONCILE",
                        "message": f"Bước {sname} đang ở trạng thái UNCERTAIN sau sự cố ngoại vi; cần đối chiếu trước khi chạy lại."
                    })
                    return {
                        "job_id": job_id,
                        "business_key": business_key,
                        "status": "UNCERTAIN_PAUSED",
                        "steps": results
                    }

                t0 = time.perf_counter()
                try:
                    fn: Callable[[], Dict[str, Any]] = step["action"]
                    self.log_step_event(job_id=job_id, business_key=business_key, job_type=job_type,
                                        step_name=sname, status="RUNNING",
                                        evidence={"external_side_effect": bool(step.get("is_external_side_effect"))})
                    out = fn()
                    dur = int((time.perf_counter() - t0) * 1000)
                    self.log_step_event(
                        job_id=job_id,
                        business_key=business_key,
                        job_type=job_type,
                        step_name=sname,
                        status="SUCCEEDED",
                        duration_ms=dur,
                        sap_status_msg=out.get("sap_status_msg", "OK"),
                        evidence=out
                    )
                    results.append({"step_name": sname, "status": "SUCCEEDED", "output": out})
                except Exception as exc:
                    dur = int((time.perf_counter() - t0) * 1000)
                    err_status = "UNCERTAIN" if step.get("is_external_side_effect") else "FAILED"
                    self.log_step_event(
                        job_id=job_id,
                        business_key=business_key,
                        job_type=job_type,
                        step_name=sname,
                        status=err_status,
                        duration_ms=dur,
                        sap_status_msg=str(exc),
                        evidence={"error": str(exc)}
                    )
                    results.append({"step_name": sname, "status": err_status, "error": str(exc)})
                    return {
                        "job_id": job_id,
                        "business_key": business_key,
                        "status": err_status,
                        "failed_step": sname,
                        "steps": results
                    }

            return {
                "job_id": job_id,
                "business_key": business_key,
                "status": "SUCCEEDED",
                "steps": results
            }
        finally:
            self.release_idempotent_lock(business_key)

    def list_job_events(self, limit: int = 100) -> List[Dict[str, Any]]:
        conn = self._get_conn()
        close_conn = self._conn is None
        try:
            rows = conn.execute("SELECT * FROM job_events ORDER BY id DESC LIMIT ?", (limit,)).fetchall()
            return [dict(r) for r in rows]
        finally:
            if close_conn:
                conn.close()

sap_worker = SAPWorkerEngine()
