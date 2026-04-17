param(
    [string]$DbHost = '127.0.0.1',
    [int]$Port = 3306,
    [string]$User = '',
    [string]$Password = '',
    [string]$Database = 'lugest'
)

$ErrorActionPreference = 'Stop'

$mysqlDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$installScript = Join-Path $mysqlDir 'install_lugest_mysql.ps1'

if (-not $User) {
    $User = 'lugest_user'
}

powershell -ExecutionPolicy Bypass -File $installScript `
    -DbHost $DbHost `
    -Port $Port `
    -AdminUser $User `
    -AdminPassword $Password `
    -Database $Database `
    -ValidateOnly

exit $LASTEXITCODE
