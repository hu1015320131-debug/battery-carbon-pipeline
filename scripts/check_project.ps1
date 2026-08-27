$projectRoot = Split-Path -Parent $PSScriptRoot
$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw "Project virtual environment is missing. Run scripts/setup_windows.ps1 first."
}

$env:PYTHONPATH = Join-Path $projectRoot "src"
& $pythonExe -m carbon_excel_pipeline check
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& $pythonExe -m unittest discover -s (Join-Path $projectRoot "tests") -p "test_*.py"
exit $LASTEXITCODE

