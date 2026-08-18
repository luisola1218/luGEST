param()

$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$gatewayScript = Join-Path $projectRoot 'scripts\run_mobile_gateway.ps1'
$cloudflared = 'C:\Program Files (x86)\cloudflared\cloudflared.exe'
$configPath = Join-Path $projectRoot 'mobile_gateway_data\cloudflare\config.yml'

if (-not (Test-Path -LiteralPath $configPath)) {
    throw 'Configuração Cloudflare inexistente. Executa primeiro setup_cloudflare_tunnel.ps1.'
}

$gatewayArguments = "-NoLogo -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$gatewayScript`" -Mode Tunnel"
$tunnelArguments = "tunnel --no-autoupdate --config `"$configPath`" run"
$userId = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Limited
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $userId
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RestartCount 10 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew

$gatewayAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $gatewayArguments -WorkingDirectory $projectRoot
$tunnelAction = New-ScheduledTaskAction -Execute $cloudflared -Argument $tunnelArguments -WorkingDirectory $projectRoot

Register-ScheduledTask -TaskName 'LuGEST Mobile Gateway' -Action $gatewayAction -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
Register-ScheduledTask -TaskName 'LuGEST Mobile Tunnel' -Action $tunnelAction -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null

Start-ScheduledTask -TaskName 'LuGEST Mobile Gateway'
Start-Sleep -Seconds 2
Start-ScheduledTask -TaskName 'LuGEST Mobile Tunnel'

Write-Host 'Arranque automático instalado.' -ForegroundColor Green
Write-Host 'O gateway e o túnel serão iniciados sempre que este utilizador entrar no Windows.'
Write-Host 'O LuGEST desktop não foi alterado e pode continuar a abrir normalmente.'
