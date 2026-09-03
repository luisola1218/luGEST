param(
    [switch]$DryRun,
    [switch]$CachesOnly,
    [switch]$TempOnly
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$repoRootFull = [System.IO.Path]::GetFullPath($repoRoot).TrimEnd('\') + '\'
$removed = New-Object System.Collections.Generic.List[string]
$found = New-Object System.Collections.Generic.List[string]

function Remove-PathSafe {
    param(
        [string]$PathToRemove
    )

    $resolvedPath = [System.IO.Path]::GetFullPath($PathToRemove)
    if (-not $resolvedPath.StartsWith($repoRootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "A limpeza recusou um caminho fora do projeto: $resolvedPath"
    }
    if (-not (Test-Path -LiteralPath $resolvedPath)) {
        return
    }

    $found.Add($resolvedPath) | Out-Null
    if ($DryRun) {
        return
    }

    Remove-Item -LiteralPath $resolvedPath -Recurse -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path -LiteralPath $resolvedPath)) {
        $removed.Add($resolvedPath) | Out-Null
    }
}

$cachePaths = @(
    'build',
    'build_qt_stable',
    '.pytest_cache',
    '.mypy_cache',
    '.ruff_cache',
    'impulse_mobile_app\.dart_tool',
    'impulse_mobile_app\build',
    'impulse_mobile_app\android\.gradle',
    'impulse_mobile_app\android\.kotlin'
)

# A execução normal também é deliberadamente conservadora. Distribuições,
# backups, dados gerados e estado local nunca são tratados como cache.
$pathsToRemove = if ($TempOnly) {
    @('tmp', '__pycache__')
}
else {
    $cachePaths
}

foreach ($relativePath in $pathsToRemove) {
    Remove-PathSafe -PathToRemove (Join-Path $repoRoot $relativePath)
}

if (-not $TempOnly) {
    Get-ChildItem $repoRoot -Recurse -Directory -Filter '__pycache__' -ErrorAction SilentlyContinue |
        Where-Object {
            $_.FullName -notlike '*\.venv\*' -and
            $_.FullName -notlike '*\dist\*' -and
            $_.FullName -notlike '*\dist_qt_stable\*' -and
            $_.FullName -notlike '*\generated\*' -and
            $_.FullName -notlike '*\backups\*'
        } |
        ForEach-Object { Remove-PathSafe -PathToRemove $_.FullName }
}

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
