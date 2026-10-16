param([switch]$Install)
$projectRoot = $PSScriptRoot
Set-Location -LiteralPath $projectRoot
if ($Install) {
    python -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw 'Cannot create Python environment' }
    & '.\.venv\Scripts\python.exe' -m pip install -r requirements-pilot.txt
    if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed' }
}
$pilotPython = if (Test-Path '.\.venv\Scripts\python.exe') { '.\.venv\Scripts\python.exe' } else { 'python' }
$env:MA_CONTROL_DB_PATH = Join-Path $projectRoot 'data\ma_control_pilot_evening.db'
$env:MA_CONTROL_EXPORT_DIR = Join-Path $projectRoot 'data\pilot_exports'
$env:MA_CONTROL_SAP_MODE = 'MANUAL'
$env:MA_CONTROL_PORT = '8012'
$env:PYTHONUTF8 = '1'
Write-Output 'MA Control pilot: http://127.0.0.1:8012 — nhập nguồn, chưa tự ghi SAP'
& $pilotPython -m uvicorn app.main:app --host 127.0.0.1 --port 8012
