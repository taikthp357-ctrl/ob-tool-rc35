# -*- coding: utf-8 -*-
"""MA CONTROL FG — Enterprise COMPANY Package Builder.

Generates:
1. Portable company distribution structure in dist/COMPANY_RELEASE_FG/
2. Release manifest.json with SHA256 checksums for each file.
3. Distribution archive: dist/MA-Control-FG-v1.0.0-rc52-COMPANY-package.zip and .sha256.
4. Corporate-ready launcher.cmd for 1-click execution from company shared drive.
"""

import os
import sys
import json
import shutil
import hashlib
import zipfile
from pathlib import Path
from datetime import datetime

REPO_ROOT = Path(__file__).resolve().parent.parent
DIST_DIR = REPO_ROOT / "dist"
RELEASE_DIR = DIST_DIR / "COMPANY_RELEASE_FG"
PAYLOAD_DIR = RELEASE_DIR / "payload"
DIST_DIR.mkdir(exist_ok=True)

VERSION = "v1.0.0-rc52-FG"
RELEASE_ID = f"MA-CONTROL-FG-{datetime.now().strftime('%Y%m%d')}"

EXCLUDE_DIRS = {
    ".git",
    ".pytest_cache",
    "__pycache__",
    "data",
    "dist",
    "tests",
    ".idea",
    ".vscode",
    "node_modules",
}

EXCLUDE_FILES = {
    ".env",
    "ma_control.db",
    "ma_control.db-wal",
    "ma_control.db-shm",
    "ma_control_fg.db",
    "ma_control_pilot.db",
    "ma_control_pilot_evening.db",
    "action_queue.json",
}

EXCLUDE_EXTS = {
    ".pyc",
    ".pyo",
    ".pyd",
    ".db",
    ".sqlite",
    ".sqlite3",
    ".log",
    ".tmp",
}


def compute_file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def build_company_package():
    print(f"[*] Starting Enterprise COMPANY Package Build for MA CONTROL FG ({VERSION})...")

    # 1. Clean payload staging (preserve runtime if already copied)
    if PAYLOAD_DIR.exists():
        shutil.rmtree(PAYLOAD_DIR, ignore_errors=True)
    RELEASE_DIR.mkdir(parents=True, exist_ok=True)
    PAYLOAD_DIR.mkdir(parents=True, exist_ok=True)

    # Copy embedded runtime if available
    runtime_src = REPO_ROOT / "runtime"
    runtime_dst = RELEASE_DIR / "runtime"
    if runtime_src.exists() and not (runtime_dst / "python.exe").exists():
        print(f"[*] Copying embedded Python runtime to {runtime_dst}...")
        shutil.copytree(runtime_src, runtime_dst, dirs_exist_ok=True)
    elif (runtime_dst / "python.exe").exists():
        print(f"[+] Reusing existing embedded Python runtime in {runtime_dst}")

    # 2. Collect payload files
    payload_files = {}

    # Copy 'app' tree
    app_src = REPO_ROOT / "app"
    for root, dirs, files in os.walk(app_src):
        dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS and not d.startswith(".")]
        for file in files:
            fp = Path(root) / file
            if fp.suffix.lower() in EXCLUDE_EXTS or fp.name in EXCLUDE_FILES:
                continue
            rel = fp.relative_to(REPO_ROOT)
            dst = PAYLOAD_DIR / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(fp, dst)
            payload_files[str(rel).replace("\\", "/")] = compute_file_sha256(dst)

    # Copy root launcher scripts to Release root (outside payload)
    shutil.copy2(REPO_ROOT / "launcher.cmd", RELEASE_DIR / "launcher.cmd")

    # Copy documentation
    for doc in ["MA_CONTROL_ARCHITECTURE_SPEC.md", "MA_CONTROL_EXCEL_GRID_UI_SPEC.md"]:
        doc_path = REPO_ROOT / "docs" / doc
        if doc_path.exists():
            dst = RELEASE_DIR / "docs" / doc
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(doc_path, dst)

    # Create empty data & logs folders in release root
    (RELEASE_DIR / "data").mkdir(parents=True, exist_ok=True)
    (RELEASE_DIR / "logs").mkdir(parents=True, exist_ok=True)

    # 3. Write manifest.json
    manifest = {
        "release_id": RELEASE_ID,
        "version": VERSION,
        "built_at": datetime.now().isoformat(),
        "zero_port_mode": True,
        "target_os": "Windows 10/11 x64 (Corporate)",
        "network_ports_required": 0,
        "files_count": len(payload_files),
        "files": payload_files
    }
    manifest_path = RELEASE_DIR / "manifest.json"
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
    print(f"[+] Wrote manifest.json with {len(payload_files)} verified files.")

    # 4. Generate portable ZIP archives
    target_zips = [
        "MA_CONTROL_FG_PORTABLE_ZERO_PORT_v1.1.0.zip",
        "MA_CONTROL_FG_ZERO_PORT_V1_0_0.zip",
        "MA-Control-FG-v1.0.0-rc52-COMPANY-package.zip",
    ]

    primary_zip = None
    for zip_name in target_zips:
        zip_path = DIST_DIR / zip_name
        if zip_path.exists():
            zip_path.unlink()

        print(f"[*] Packaging distribution into {zip_path}...")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(RELEASE_DIR):
                for file in files:
                    fp = Path(root) / file
                    arcname = fp.relative_to(RELEASE_DIR)
                    zf.write(fp, arcname)

        zip_hash = compute_file_sha256(zip_path)
        sha_file = DIST_DIR / f"{zip_name}.sha256"
        with open(sha_file, "w", encoding="utf-8") as f:
            f.write(f"{zip_hash} *{zip_name}\n")

        print(f"[SUCCESS] Packaged {zip_name} ({zip_path.stat().st_size / (1024*1024):.2f} MB), SHA256: {zip_hash[:16]}...")
        if primary_zip is None:
            primary_zip = str(zip_path)

    return primary_zip


if __name__ == "__main__":
    build_company_package()
