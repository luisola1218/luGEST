param(
    [string]$Url = ''
)

$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$envPath = Join-Path $projectRoot 'lugest.env'
if (-not $Url) {
    $urlPath = Join-Path $projectRoot 'mobile_gateway_data\cloudflare\public_url.txt'
    if (-not (Test-Path -LiteralPath $urlPath)) { throw 'Indica -Url ou configura o túnel permanente.' }
    $Url = (Get-Content -LiteralPath $urlPath -Raw).Trim()
}
$tokenLine = Get-Content -LiteralPath $envPath -Encoding UTF8 |
    Where-Object { $_ -match '^LUGEST_MOBILE_API_TOKEN=.+' } |
    Select-Object -Last 1
if (-not $tokenLine) { throw 'Chave móvel não encontrada.' }
$token = ($tokenLine -split '=', 2)[1].Trim()
$headers = @{ Authorization = "Bearer $token" }

$health = Invoke-RestMethod -Uri "$($Url.TrimEnd('/'))/health" -Headers $headers -TimeoutSec 20
$stock = Invoke-RestMethod -Uri "$($Url.TrimEnd('/'))/v1/stock/summary" -Headers $headers -TimeoutSec 20
if (-not $health.ok) { throw 'O endpoint remoto não respondeu corretamente.' }

Write-Host 'Ligação remota validada.' -ForegroundColor Green
Write-Host "URL: $Url"
Write-Host "Gateway: $($health.version)"
Write-Host "Dispositivo: $($health.device)"
Write-Host "Produtos: $($stock.products) | Materiais: $($stock.materials)"
