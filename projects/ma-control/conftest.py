# -*- coding: utf-8 -*-
import sys
import os
import tempfile
import atexit
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
_test_database_dir = tempfile.TemporaryDirectory(prefix="ma-control-tests-")
atexit.register(_test_database_dir.cleanup)
os.environ["MA_CONTROL_DB_PATH"] = str(Path(_test_database_dir.name) / "test.db")
os.environ["MA_CONTROL_SAP_MODE"] = "TEST"
os.environ["MA_CONTROL_EXPORT_DIR"] = str(Path(_test_database_dir.name) / "exports")
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))
