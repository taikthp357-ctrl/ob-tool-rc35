# -*- coding: utf-8 -*-
"""System configuration for MA CONTROL - Dieu hanh FG."""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent.parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

DB_PATH = Path(os.environ.get("MA_CONTROL_DB_PATH", str(DATA_DIR / "ma_control_fg.db")))
LEGACY_DB_PATH = DATA_DIR / "ma_control.db"

HOST = os.environ.get("MA_CONTROL_HOST", "127.0.0.1")
PORT = int(os.environ.get("MA_CONTROL_PORT", "8012"))
SAP_MODE = os.environ.get("MA_CONTROL_SAP_MODE", "MANUAL").upper()

APP_TITLE = "MA CONTROL — Điều hành FG"
APP_VERSION = "v1.0.0-rc52-FG"
