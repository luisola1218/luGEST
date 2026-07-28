param(
    [switch]$DryRun,
    [switch]$CachesOnly
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

$pathsToRemove = if ($CachesOnly) {
    @('build', 'build_qt_stable', '.pytest_cache', '.mypy_cache', '.ruff_cache')
}
else {
    @(
        'backups',
        'build',
        'build_qt_stable',
        'dist',
        'dist_qt_stable',
        'generated',
        'dist\lugest_trial.json',
        'lugest_runtime_state.json',
        'lugest_supplier_seq.json',
        'lugest_transport_seq.json',
        'lugest_trial.json',
        '.cad312',
        '.pytest_cache',
        '.mypy_cache',
        '.ruff_cache',
        '.idea',
        'previews'
    )
}

foreach ($relativePath in $pathsToRemove) {
    Remove-PathSafe -PathToRemove (Join-Path $repoRoot $relativePath)
}

Get-ChildItem $repoRoot -Recurse -Directory -Filter '__pycache__' -ErrorAction SilentlyContinue |
    Where-Object {
        $_.FullName -notlike '*\.venv\*' -and
        $_.FullName -notlike '*\dist\*' -and
        $_.FullName -notlike '*\dist_qt_stable\*' -and
        $_.FullName -notlike '*\generated\*' -and
        $_.FullName -notlike '*\backups\*'
    } |
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
