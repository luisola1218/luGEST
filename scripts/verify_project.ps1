param(
    [switch]$SkipCompile,
    [switch]$SkipCoreFlows,
    [switch]$AllowRemoteDatabase
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

if (-not $SkipCoreFlows -and -not $AllowRemoteDatabase) {
    $envFile = Join-Path $repoRoot 'lugest.env'
    if (Test-Path -LiteralPath $envFile) {
        $hostLine = Get-Content -LiteralPath $envFile |
            Where-Object { $_ -match '^\s*LUGEST_DB_HOST\s*=' } |
            Select-Object -First 1
        if ($hostLine) {
            $dbHost = (($hostLine -split '=', 2)[1]).Trim().Trim('"').Trim("'")
            $localHosts = @('', 'localhost', '127.0.0.1', '::1')
            if ($dbHost -notin $localHosts) {
                throw "Verificacao funcional bloqueada: a base configurada e remota ($dbHost). Usa -AllowRemoteDatabase apenas numa base de testes, nunca na base do cliente."
            }
        }
    }
}

if (-not $SkipCompile) {
    Write-Host "A compilar ficheiros Python..." -ForegroundColor Cyan
    $compileScript = @'
import py_compile
from pathlib import Path

excluded = {
    ".venv",
    ".cad312",
    "backups",
    "build",
    "build_qt_stable",
    "dist",
    "dist_qt_stable",
    "generated",
    "output",
    "tmp",
}
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

    Write-Host "A correr performance Qt..." -ForegroundColor Cyan
    & $python (Join-Path $repoRoot 'scripts\verify_qt_performance.py')
}

Write-Host "Verificacao concluida." -ForegroundColor Green
