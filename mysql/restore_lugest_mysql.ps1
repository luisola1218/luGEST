param(
    [string]$DbHost = '127.0.0.1',
    [int]$Port = 3306,
    [string]$AdminUser = '',
    [string]$AdminPassword = '',
    [string]$Database = 'lugest',
    [string]$Input = '',
    [string]$BackupRoot = '',
    [string]$MySqlPath = '',
    [switch]$ResetDatabase,
    [switch]$SkipValidation,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$mysqlDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $mysqlDir 'restore_lugest_mysql.py'
$pythonCandidates = @(
    (Join-Path (Resolve-Path (Join-Path $mysqlDir '..\..')).Path '.venv\Scripts\python.exe'),
    'py',
    'python'
)

$pythonCmd = $null
foreach ($candidate in $pythonCandidates) {
    if ($candidate -match '\.exe$') {
        if (Test-Path $candidate) {
            $pythonCmd = $candidate
            break
        }
        continue
    }
    if (Get-Command $candidate -ErrorAction SilentlyContinue) {
        $pythonCmd = $candidate
        break
    }
}
if (-not $pythonCmd) {
    throw 'Nao foi encontrado Python para executar o restauro.'
}

$arguments = @(
    $scriptPath,
    '--host', $DbHost,
    '--port', "$Port",
    '--admin-user', $AdminUser,
    '--admin-password', $AdminPassword,
    '--database', $Database
)
if ($Input) { $arguments += @('--input', $Input) }
if ($BackupRoot) { $arguments += @('--backup-root', $BackupRoot) }
if ($MySqlPath) { $arguments += @('--mysql-path', $MySqlPath) }
if ($ResetDatabase) { $arguments += '--reset-database' }
if ($SkipValidation) { $arguments += '--skip-validation' }
if ($DryRun) { $arguments += '--dry-run' }

& $pythonCmd @arguments
exit $LASTEXITCODE
