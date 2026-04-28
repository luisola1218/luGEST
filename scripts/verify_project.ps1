param(
    [switch]$SkipCompile,
    [switch]$SkipCoreFlows
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$venvPython = Join-Path $repoRoot '.venv\Scripts\python.exe'

if (Test-Path $venvPython) {
    $python = $venvPython
}
else {
    $python = 'python'
}

Write-Host "Python: $python" -ForegroundColor Cyan

if (-not $SkipCompile) {
    Write-Host "A compilar ficheiros Python..." -ForegroundColor Cyan
    $compileScript = @'
import py_compile
from pathlib import Path

excluded = {".venv", ".cad312", "build", "dist", "generated"}
files = [
    path
    for path in Path(".").rglob("*.py")
    if not any(part in excluded for part in path.parts)
]
errors = []
for path in files:
    try:
        py_compile.compile(str(path), doraise=True)
    except Exception as exc:
        errors.append((str(path), exc))

if errors:
    for path, exc in errors:
        print(f"{path}: {exc}")
    raise SystemExit(1)

print(f"Compiled {len(files)} Python files")
'@
    $compileScript | & $python -
}

if (-not $SkipCoreFlows) {
    Write-Host "A correr fluxos principais..." -ForegroundColor Cyan
    & $python (Join-Path $repoRoot 'scripts\verify_core_flows.py')
}

Write-Host "Verificacao concluida." -ForegroundColor Green
