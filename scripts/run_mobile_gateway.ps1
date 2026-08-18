param(
    [int]$Port = 8765,
    [ValidateSet('Lan', 'Tunnel')]
    [string]$Mode = 'Lan'
)

$ErrorActionPreference = "Stop"
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$envPath = Join-Path $projectRoot "lugest.env"
$pythonPath = Join-Path $projectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $pythonPath)) {
    throw "Ambiente Python não encontrado em $pythonPath"
}
if (-not (Test-Path -LiteralPath $envPath)) {
    throw "Ficheiro lugest.env não encontrado."
}

$lines = Get-Content -LiteralPath $envPath -Encoding UTF8
$hasToken = $false
foreach ($line in $lines) {
    if ($line -match '^\s*([^#][^=]*)=(.*)$') {
        $name = $matches[1].Trim()
        $value = $matches[2].Trim().Trim('"').Trim("'")
        if ($name) {
            [Environment]::SetEnvironmentVariable($name, $value, 'Process')
            if ($name -eq 'LUGEST_MOBILE_API_TOKEN' -and $value) { $hasToken = $true }
        }
    }
}

if (-not $hasToken) {
    $bytes = New-Object byte[] 32
    $generator = [Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $generator.GetBytes($bytes)
    }
    finally {
        $generator.Dispose()
    }
    $token = ([BitConverter]::ToString($bytes)).Replace('-', '')
    Add-Content -LiteralPath $envPath -Value "`nLUGEST_MOBILE_API_TOKEN=$token" -Encoding UTF8
    [Environment]::SetEnvironmentVariable('LUGEST_MOBILE_API_TOKEN', $token, 'Process')
}

[Environment]::SetEnvironmentVariable('LUGEST_MOBILE_API_PORT', "$Port", 'Process')
$isTunnel = $Mode -eq 'Tunnel'
[Environment]::SetEnvironmentVariable(
    'LUGEST_MOBILE_API_BIND',
    $(if ($isTunnel) { '127.0.0.1' } else { '0.0.0.0' }),
    'Process'
)
[Environment]::SetEnvironmentVariable(
    'LUGEST_MOBILE_ALLOW_LEGACY_TOKEN',
    $(if ($isTunnel) { '0' } else { '1' }),
    'Process'
)

if ($isTunnel) {
    $deviceStore = Join-Path $projectRoot 'mobile_gateway_data\devices.json'
    if (-not (Test-Path -LiteralPath $deviceStore)) {
        Set-Location -LiteralPath $projectRoot
        & $pythonPath -m mobile_gateway.devices import-legacy --name 'Telemóvel principal'
        if ($LASTEXITCODE -ne 0) { throw 'Não foi possível preparar o dispositivo móvel.' }
    }
}

$lanAddress = Get-NetIPAddress -AddressFamily IPv4 |
    Where-Object { $_.IPAddress -notlike '127.*' -and $_.PrefixOrigin -ne 'WellKnown' } |
    Sort-Object InterfaceMetric |
    Select-Object -First 1 -ExpandProperty IPAddress

Write-Host ""
Write-Host "LuGEST Field - ligação ao stock" -ForegroundColor Green
if ($isTunnel) {
    Write-Host "Modo seguro: apenas túnel HTTPS (127.0.0.1:$Port)"
    Write-Host "A porta não está exposta na rede local." -ForegroundColor Green
}
else {
    Write-Host "Modo de testes LAN: http://${lanAddress}:$Port"
    Write-Host "Chave de acesso: $([Environment]::GetEnvironmentVariable('LUGEST_MOBILE_API_TOKEN', 'Process'))"
}
Write-Host "Mantenha esta janela aberta enquanto usa a app." -ForegroundColor Yellow
Write-Host ""

Set-Location -LiteralPath $projectRoot
& $pythonPath -m mobile_gateway.server
