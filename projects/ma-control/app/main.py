# -*- coding: utf-8 -*-
"""MA CONTROL — FastAPI Application Server (FG Local Webapp)."""

import sys
import uuid
import datetime
from contextlib import asynccontextmanager
from typing import List, Optional, Dict, Any
from pathlib import Path
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from app.core.config import APP_TITLE, APP_VERSION, HOST, PORT, BASE_DIR
from app.core.database import get_connection, init_db
from app.api.routes_workflow import router as workflow_router
from app.api.routes_config import router as config_router

@asynccontextmanager
async def lifespan(application: FastAPI):
    init_db()
    yield

app = FastAPI(title=APP_TITLE, version=APP_VERSION, lifespan=lifespan)

STATIC_DIR = BASE_DIR / "app" / "static"
TEMPLATES_DIR = BASE_DIR / "app" / "templates"

if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

app.include_router(workflow_router)
app.include_router(config_router)

@app.get("/", response_class=HTMLResponse)
def index_page():
    index_file = TEMPLATES_DIR / "index.html"
    if index_file.exists():
        return HTMLResponse(content=index_file.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>MA CONTROL — Error: index.html not found</h1>", status_code=404)

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="MA CONTROL — Điều hành FG")
    parser.add_argument("--gui", "--desktop", action="store_true", help="Launch native desktop GUI (Tkinter)")
    parser.add_argument("--sync", action="store_true", help="Sync SQLite data to data.js")
    parser.add_argument("--server", action="store_true", help="Launch FastAPI web server (Development only)")
    parser.add_argument("--dry-run", action="store_true", help="Print Edge launch command without opening browser")
    args = parser.parse_args()

    if args.sync:
        from app.web_bridge import sync_web_dashboard
        res = sync_web_dashboard()
        print(f"[OK] Synced {len(res['pick_orders'])} pick orders to data.js.")
        sys.exit(0)

    if args.gui:
        from app.gui.main_window import launch_gui
        launch_gui()
        sys.exit(0)

    if args.server:
        import uvicorn
        uvicorn.run("app.main:app", host=HOST, port=PORT, reload=True)
        sys.exit(0)

    # Canonical default: Zero-Port Edge App Mode
    from app.launcher_core import EnterpriseLauncher
    launcher = EnterpriseLauncher(BASE_DIR)
    if args.dry_run:
        prep = launcher.prepare_runtime()
        print("[DRY_RUN] Runtime prepared for Zero-Port launch:", prep)
        sys.exit(0)
    pid = launcher.launch(mode="web")
    print(f"[OK] Launched MA CONTROL FG in Zero-Port Edge App Mode (PID: {pid}).")
