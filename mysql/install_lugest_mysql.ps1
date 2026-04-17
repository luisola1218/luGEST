param(
    [string]$DbHost = '127.0.0.1',
    [int]$Port = 3306,
    [string]$AdminUser = 'root',
    [string]$AdminPassword = '',
    [string]$Database = 'lugest',
    [string]$AppUser = 'lugest_user',
    [string]$AppPassword = '',
    [string]$AppHost = 'localhost',
    [switch]$ResetDatabase,
    [switch]$SkipBaseSchema,
    [switch]$SkipPatches,
    [switch]$SkipValidation,
    [switch]$ValidateOnly,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$mysqlDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $mysqlDir 'install_lugest_mysql.py'
$releaseRoot = Resolve-Path (Join-Path $mysqlDir '..\..')

function Resolve-PythonCommand {
    $candidates = @(
        (Join-Path $releaseRoot 'Mobile API\.venv\Scripts\python.exe'),
        (Join-Path $releaseRoot 'impulse_mobile_api\.venv\Scripts\python.exe'),
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
    throw 'Nao foi encontrado Python. Instala o Python ou executa primeiro a instalacao da Mobile API.'
}

$pythonCmd = Resolve-PythonCommand
$arguments = @(
    $scriptPath,
    '--host', $DbHost,
    '--port', "$Port",
    '--admin-user', $AdminUser,
    '--admin-password', $AdminPassword,
    '--database', $Database
)

if ($AppUser) {
    $arguments += @('--app-user', $AppUser)
}
if ($AppPassword) {
    $arguments += @('--app-password', $AppPassword)
}
if ($AppHost) {
    $arguments += @('--app-host', $AppHost)
}
if ($ResetDatabase) {
    $arguments += '--reset-database'
}
if ($SkipBaseSchema) {
    $arguments += '--skip-base-schema'
}
if ($SkipPatches) {
    $arguments += '--skip-patches'
}
if ($SkipValidation) {
    $arguments += '--skip-validation'
}
if ($ValidateOnly) {
    $arguments += '--validate-only'
}
if ($DryRun) {
    $arguments += '--dry-run'
}

& $pythonCmd @arguments
exit $LASTEXITCODE
