param(
    [Parameter(Mandatory = $true)]
    [string]$Hostname,
    [string]$TunnelName = 'lugest-field'
)

$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$cloudflared = 'C:\Program Files (x86)\cloudflared\cloudflared.exe'
$dataDirectory = Join-Path $projectRoot 'mobile_gateway_data\cloudflare'
$configPath = Join-Path $dataDirectory 'config.yml'

if ($Hostname -notmatch '^[A-Za-z0-9](?:[A-Za-z0-9.-]{1,251})\.[A-Za-z]{2,63}$') {
    throw 'Indica um endereço completo válido, por exemplo field.empresa.pt.'
}
if (-not (Test-Path -LiteralPath $cloudflared)) {
    throw 'cloudflared não está instalado. Executa primeiro o instalador preparado.'
}
New-Item -ItemType Directory -Path $dataDirectory -Force | Out-Null

Write-Host 'Será aberta a autenticação Cloudflare no navegador.' -ForegroundColor Cyan
Write-Host 'Escolha o domínio onde pretende criar o endereço da app.'
& $cloudflared tunnel login
if ($LASTEXITCODE -ne 0) { throw 'A autenticação Cloudflare não foi concluída.' }

$tunnels = @(& $cloudflared tunnel list --output json | ConvertFrom-Json)
$tunnel = $tunnels | Where-Object { $_.name -eq $TunnelName } | Select-Object -First 1
if (-not $tunnel) {
    & $cloudflared tunnel create $TunnelName
    if ($LASTEXITCODE -ne 0) { throw 'Não foi possível criar o túnel.' }
    $tunnels = @(& $cloudflared tunnel list --output json | ConvertFrom-Json)
    $tunnel = $tunnels | Where-Object { $_.name -eq $TunnelName } | Select-Object -First 1
}
if (-not $tunnel -or -not $tunnel.id) { throw 'O túnel criado não foi encontrado.' }

$credentialsPath = Join-Path $env:USERPROFILE ".cloudflared\$($tunnel.id).json"
if (-not (Test-Path -LiteralPath $credentialsPath)) {
    throw "Credenciais do túnel não encontradas em $credentialsPath"
}

$yamlCredentials = $credentialsPath.Replace('\', '/')
$config = @"
tunnel: $($tunnel.id)
credentials-file: $yamlCredentials

ingress:
  - hostname: $Hostname
    service: http://127.0.0.1:8765
    originRequest:
      connectTimeout: 10s
      httpHostHeader: localhost
  - service: http_status:404
"@
Set-Content -LiteralPath $configPath -Value $config -Encoding UTF8

& $cloudflared tunnel route dns $($tunnel.id) $Hostname
if ($LASTEXITCODE -ne 0) { throw 'Não foi possível criar o registo DNS do túnel.' }

Set-Content -LiteralPath (Join-Path $dataDirectory 'public_url.txt') -Value "https://$Hostname" -Encoding ASCII
Write-Host ''
Write-Host 'Túnel permanente preparado com sucesso.' -ForegroundColor Green
Write-Host "Endereço da app: https://$Hostname"
Write-Host "Configuração: $configPath"
Write-Host 'O passo seguinte é instalar o arranque automático.' -ForegroundColor Yellow
