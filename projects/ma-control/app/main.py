# -*- coding: utf-8 -*-
"""MA CONTROL — FastAPI Application Server (FG Local Webapp)."""

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
    import uvicorn
    uvicorn.run("app.main:app", host=HOST, port=PORT, reload=True)
