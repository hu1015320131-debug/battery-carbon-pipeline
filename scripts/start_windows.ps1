$projectRoot = Split-Path -Parent $PSScriptRoot
$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"
$appPath = Join-Path $projectRoot "app\streamlit_app.py"

if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw "Project virtual environment is missing. Run scripts/setup_windows.ps1 first."
}

$env:PYTHONPATH = Join-Path $projectRoot "src"
& $pythonExe -m streamlit run $appPath --server.address 127.0.0.1 --server.port 8501 --browser.gatherUsageStats false
exit $LASTEXITCODE
