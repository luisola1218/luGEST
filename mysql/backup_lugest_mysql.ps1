param(
    [string]$DbHost = '127.0.0.1',
    [int]$Port = 3306,
    [string]$User = '',
    [string]$Password = '',
    [string]$Database = 'lugest',
    [string]$Label = '',
    [string]$OutputRoot = '',
    [string]$MySqlDumpPath = '',
    [switch]$PlainSql,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$mysqlDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $mysqlDir 'backup_lugest_mysql.py'
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
    throw 'Nao foi encontrado Python para executar o backup.'
}

$arguments = @(
    $scriptPath,
    '--host', $DbHost,
    '--port', "$Port",
    '--user', $User,
    '--password', $Password,
    '--database', $Database
)
if ($Label) { $arguments += @('--label', $Label) }
if ($OutputRoot) { $arguments += @('--output-root', $OutputRoot) }
if ($MySqlDumpPath) { $arguments += @('--mysqldump-path', $MySqlDumpPath) }
if ($PlainSql) { $arguments += '--plain-sql' }
if ($DryRun) { $arguments += '--dry-run' }

& $pythonCmd @arguments
exit $LASTEXITCODE
