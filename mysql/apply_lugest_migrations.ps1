param(
    [string]$DbHost = '',
    [int]$Port = 0,
    [string]$AdminUser = '',
    [string]$AdminPassword = '',
    [string]$Database = '',
    [string]$Actor = '',
    [switch]$Status,
    [switch]$BaselineCurrent,
    [switch]$LegacyApplyAll,
    [switch]$SkipValidation,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$mysqlDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $mysqlDir 'apply_lugest_migrations.py'
$releaseRoot = Resolve-Path (Join-Path $mysqlDir '..\..')

function Resolve-PythonCommand {
    $candidates = @(
        (Join-Path $releaseRoot '.venv\Scripts\python.exe'),
        'py',
        'python'
    )
    foreach ($candidate in $candidates) {
        if ($candidate -match '\.exe$') {
            if (Test-Path $candidate) {
                return $candidate
            }
            continue
        }
        $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($cmd) {
            return $candidate
        }
    }
    throw 'Nao foi encontrado Python. Instala o Python ou executa a partir da pasta desktop LuisGEST.'
}

$pythonCmd = Resolve-PythonCommand
$arguments = @($scriptPath)

if ($DbHost) {
    $arguments += @('--host', $DbHost)
}
if ($Port -gt 0) {
    $arguments += @('--port', "$Port")
}
if ($Database) {
    $arguments += @('--database', $Database)
}

if ($AdminUser) {
    $arguments += @('--admin-user', $AdminUser)
}
if ($AdminPassword) {
    $arguments += @('--admin-password', $AdminPassword)
}
if ($Actor) {
    $arguments += @('--actor', $Actor)
}
if ($Status) {
    $arguments += '--status'
}
if ($BaselineCurrent) {
    $arguments += '--baseline-current'
}
if ($LegacyApplyAll) {
    $arguments += '--legacy-apply-all'
}
if ($SkipValidation) {
    $arguments += '--skip-validation'
}
if ($DryRun) {
    $arguments += '--dry-run'
}

& $pythonCmd @arguments
exit $LASTEXITCODE
