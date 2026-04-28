param(
    [switch]$DryRun,
    [switch]$IncludeFlutterBuild
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$removed = New-Object System.Collections.Generic.List[string]
$found = New-Object System.Collections.Generic.List[string]

function Remove-PathSafe {
    param(
        [string]$PathToRemove
    )

    if (-not (Test-Path $PathToRemove)) {
        return
    }

    $found.Add($PathToRemove) | Out-Null
    if ($DryRun) {
        return
    }

    Remove-Item $PathToRemove -Recurse -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path $PathToRemove)) {
        $removed.Add($PathToRemove) | Out-Null
    }
}

$pathsToRemove = @(
    'build',
    'build_qt_stable',
    'dist_qt_stable',
    'dist\lugest_trial.json',
    '.cad312',
    '.pytest_cache',
    '.mypy_cache',
    '.ruff_cache',
    '.idea',
    'impulse_mobile_app\.dart_tool',
    'impulse_mobile_app\.idea',
    'impulse_mobile_app\android\.gradle',
    'impulse_mobile_app\android\.kotlin',
    'impulse_mobile_app\.flutter-plugins-dependencies',
    'impulse_mobile_api\api_runtime.log',
    'impulse_mobile_api\api_stdout.log',
    'impulse_mobile_api\api_stderr.log',
    'impulse_mobile_api\.server.pid',
    'previews'
)

if ($IncludeFlutterBuild) {
    $pathsToRemove += 'impulse_mobile_app\build'
}

foreach ($relativePath in $pathsToRemove) {
    Remove-PathSafe -PathToRemove (Join-Path $repoRoot $relativePath)
}

Get-ChildItem $repoRoot -Recurse -Directory -Filter '__pycache__' -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -notlike '*\.venv\*' } |
    ForEach-Object { Remove-PathSafe -PathToRemove $_.FullName }

Write-Host ""
if ($DryRun) {
    Write-Host "Limpeza simulada. Itens encontrados:" -ForegroundColor Yellow
    $found | Sort-Object -Unique | ForEach-Object { Write-Host " - $_" }
}
else {
    Write-Host "Limpeza concluida. Itens removidos:" -ForegroundColor Green
    $removed | Sort-Object -Unique | ForEach-Object { Write-Host " - $_" }
}

Write-Host ""
Get-ChildItem $repoRoot -Force | Select-Object Name, Mode, LastWriteTime
