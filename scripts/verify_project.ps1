param(
    [switch]$SkipCompile,
    [switch]$SkipCoreFlows,
    [switch]$AllowRemoteDatabase,
    [switch]$SafeOnly,
    [switch]$ReadOnlyDatabaseAudit
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

function Invoke-Checked {
    param(
        [string]$Label,
        [string[]]$CommandArgs
    )

    Write-Host $Label -ForegroundColor Cyan
    & $python @CommandArgs
    if ($LASTEXITCODE -ne 0) {
        throw "$Label falhou com codigo $LASTEXITCODE."
    }
}

Push-Location -LiteralPath $repoRoot
try {
if (-not $SafeOnly -and -not $SkipCoreFlows -and -not $AllowRemoteDatabase) {
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
from pathlib import Path
import os

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
    ".git",
    "__pycache__",
    ".dart_tool",
    "node_modules",
}
files = []
for directory, dirs, names in os.walk("."):
    dirs[:] = [name for name in dirs if name not in excluded]
    files.extend(Path(directory) / name for name in names if name.endswith(".py"))
errors = []
for path in files:
    try:
        source = path.read_text(encoding="utf-8-sig")
        compile(source, str(path), "exec", dont_inherit=True)
    except Exception as exc:
        errors.append((str(path), exc))

if errors:
    for path, exc in errors:
        print(f"{path}: {exc}")
    raise SystemExit(1)

print(f"Compiled {len(files)} Python files")
'@
    $compileScript | & $python -
    if ($LASTEXITCODE -ne 0) {
        throw "A compilacao Python falhou com codigo $LASTEXITCODE."
    }
}

if ($SafeOnly) {
    Write-Host "A validar dependencias instaladas..." -ForegroundColor Cyan
    & $python -m pip check
    if ($LASTEXITCODE -ne 0) {
        throw "A validacao de dependencias falhou com codigo $LASTEXITCODE."
    }

    $safeScripts = @(
        'scripts\security_audit.py',
        'scripts\verify_architecture_boundaries.py',
        'scripts\verify_locked_dependencies.py',
        'scripts\verify_app_storage.py',
        'scripts\verify_update_security.py',
        'scripts\verify_runtime_diagnostics.py',
        'scripts\verify_license_foundation.py',
        'scripts\verify_mysql_schema_definition.py',
        'scripts\verify_migration_framework.py',
        'scripts\verify_laser_quote_engine.py',
        'scripts\verify_laser_nesting_flow.py',
        'scripts\verify_quote_ui_controls.py',
        'scripts\verify_partner_ui_layout.py',
        'scripts\verify_runtime_cache.py',
        'scripts\verify_page_modules.py',
        'scripts\verify_backend_modules.py',
        'scripts\verify_snapshot_merge.py',
        'scripts\verify_configuration_repository.py',
        'scripts\verify_operation_costing.py',
        'scripts\verify_clients_integration.py',
        'scripts\verify_business_modules.py',
        'scripts\verify_quote_editors.py',
        'scripts\verify_nesting_study_service.py',
        'scripts\verify_quote_snapshot_save.py'
    )
    foreach ($relativeScript in $safeScripts) {
        Invoke-Checked -Label "A correr $relativeScript..." -CommandArgs @((Join-Path $repoRoot $relativeScript))
    }
}
elseif (-not $SkipCoreFlows) {
    Invoke-Checked -Label "A correr fluxos principais..." -CommandArgs @((Join-Path $repoRoot 'scripts\verify_core_flows.py'))

    Invoke-Checked -Label "A correr performance Qt..." -CommandArgs @((Join-Path $repoRoot 'scripts\verify_qt_performance.py'))
}

if ($ReadOnlyDatabaseAudit) {
    Invoke-Checked -Label "A auditar a base de dados em modo apenas de leitura..." -CommandArgs @((Join-Path $repoRoot 'scripts\audit_database_readonly.py'))
}

Write-Host "Verificacao concluida." -ForegroundColor Green
}
finally {
    Pop-Location
}
