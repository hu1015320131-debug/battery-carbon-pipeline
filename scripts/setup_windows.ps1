$projectRoot = Split-Path -Parent $PSScriptRoot
$venvPath = Join-Path $projectRoot ".venv"

if (-not (Test-Path -LiteralPath $venvPath)) {
    python -m venv $venvPath
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

$pythonExe = Join-Path $venvPath "Scripts\python.exe"
& $pythonExe -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& $pythonExe -m pip install -e "$projectRoot[dev]"
exit $LASTEXITCODE

