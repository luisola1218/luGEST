param(
    [string]$InstallDir = ""
)

$ErrorActionPreference = 'Stop'

function Read-EnvFile {
    param([string]$Path)
    $map = @{}
    if (-not (Test-Path $Path)) {
        return $map
    }
    foreach ($raw in Get-Content $Path -Encoding UTF8) {
        $line = [string]$raw
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $trimmed = $line.Trim()
        if ($trimmed.StartsWith('#')) { continue }
        $eqIndex = $trimmed.IndexOf('=')
        if ($eqIndex -lt 1) { continue }
        $key = $trimmed.Substring(0, $eqIndex).Trim()
        $value = $trimmed.Substring($eqIndex + 1).Trim()
        if ($key) {
            $map[$key] = $value
        }
    }
    return $map
}

if (-not $InstallDir) {
    if (Test-Path (Join-Path $env:ProgramFiles 'LuisGEST')) {
        $InstallDir = Join-Path $env:ProgramFiles 'LuisGEST'
    }
    else {
        $InstallDir = Join-Path $env:LOCALAPPDATA 'LuisGEST'
    }
}

$envPath = Join-Path $InstallDir 'lugest.env'
if (-not (Test-Path $envPath)) {
    throw "Nao foi encontrado lugest.env em: $envPath"
}

$cfg = Read-EnvFile -Path $envPath
$dbHost = [string]($cfg['LUGEST_DB_HOST'])
$portTxt = [string]($cfg['LUGEST_DB_PORT'])
$dbName = [string]($cfg['LUGEST_DB_NAME'])
$user = [string]($cfg['LUGEST_DB_USER'])
$pass = [string]($cfg['LUGEST_DB_PASS'])

Write-Host ''
Write-Host 'Diagnostico LuisGEST' -ForegroundColor Cyan
Write-Host "Pasta instalada: $InstallDir"
Write-Host "Ficheiro env: $envPath"
Write-Host "Host: $dbHost"
Write-Host "Porta: $portTxt"
Write-Host "Base de dados: $dbName"
Write-Host "Utilizador: $user"

$issues = @()
if (-not $dbHost) { $issues += 'Falta LUGEST_DB_HOST.' }
if (-not $portTxt) { $issues += 'Falta LUGEST_DB_PORT.' }
if (-not $dbName) { $issues += 'Falta LUGEST_DB_NAME.' }
if (-not $user) { $issues += 'Falta LUGEST_DB_USER.' }
if (-not $pass) { $issues += 'Falta LUGEST_DB_PASS.' }
$isLocalHost = $dbHost -eq '127.0.0.1' -or $dbHost -eq 'localhost'
if ($user -eq 'lugest_user') {
    $issues += 'O utilizador MySQL parece ainda ser o exemplo.'
}
if ($pass -eq 'trocar-password') {
    $issues += 'A password MySQL parece ainda ser a de exemplo.'
}

[int]$port = 0
[void][int]::TryParse($portTxt, [ref]$port)
if ($port -gt 0 -and $dbHost) {
    try {
        $tcp = Test-NetConnection -ComputerName $dbHost -Port $port -WarningAction SilentlyContinue
        if ($tcp.TcpTestSucceeded) {
            Write-Host "Ligacao TCP a $dbHost`:$port OK" -ForegroundColor Green
            if ($isLocalHost) {
                Write-Host "MySQL local detetado neste posto." -ForegroundColor DarkGreen
            }
        }
        else {
            $issues += "Nao foi possivel ligar por TCP a $dbHost`:$port."
        }
    }
    catch {
        $issues += "Falhou o teste de rede a $dbHost`:$port. $_"
    }
}

if ($isLocalHost -and -not $issues.Count) {
    Write-Host "Host local aceite porque o servidor MySQL deste posto respondeu." -ForegroundColor DarkGreen
}

if ($issues.Count) {
    Write-Host ''
    Write-Host 'Problemas detetados:' -ForegroundColor Yellow
    foreach ($item in $issues) {
        Write-Host " - $item"
    }
    exit 1
}

Write-Host ''
Write-Host 'Configuracao base validada sem problemas obvios.' -ForegroundColor Green
