param(
    [string]$TaskName = 'LUGEST Impulse API',
    [switch]$OpenFirewall,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$apiRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExe = Join-Path $apiRoot '.venv\Scripts\python.exe'
$runnerScript = Join-Path $apiRoot 'run_impulse_api_server.ps1'
$powershellExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-DotEnvValue {
    param(
        [string]$Key,
        [string]$DefaultValue = ''
    )

    foreach ($candidate in @(
        (Join-Path $apiRoot '.env'),
        (Join-Path $apiRoot '.env.example')
    )) {
        if (-not (Test-Path $candidate)) {
            continue
        }
        foreach ($raw in Get-Content $candidate) {
            $line = ($raw | Out-String).Trim()
            if (-not $line -or $line.StartsWith('#') -or -not $line.Contains('=')) {
                continue
            }
            $parts = $line.Split('=', 2)
            if ($parts[0].Trim() -eq $Key) {
                return $parts[1].Trim().Trim('"').Trim("'")
            }
        }
    }
    return $DefaultValue
}

if ((-not $DryRun) -and (-not (Test-IsAdmin))) {
    throw 'Executa este script num PowerShell com privilegios de Administrador.'
}
if (-not (Test-Path $pythonExe)) {
    throw "Falta a venv da API: $pythonExe"
}
if (-not (Test-Path $runnerScript)) {
    throw "Falta o runner da API: $runnerScript"
}

$apiPort = Get-DotEnvValue -Key 'LUGEST_API_PORT' -DefaultValue '8050'
$firewallRuleName = "LUGEST Impulse API $apiPort"
$taskArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$runnerScript`""

Write-Host "Task.......: $TaskName"
Write-Host "Runner.....: $runnerScript"
Write-Host "Python.....: $pythonExe"
Write-Host "Porta API..: $apiPort"
Write-Host "Firewall...: $firewallRuleName"

if ($DryRun) {
    Write-Host ''
    Write-Host 'DryRun ativo. Nenhuma alteracao foi aplicada.'
    return
}

$action = New-ScheduledTaskAction -Execute $powershellExe -Argument $taskArgs -WorkingDirectory $apiRoot
$trigger = New-ScheduledTaskTrigger -AtStartup
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description 'Arranca a LUGEST Impulse Mobile API no boot do servidor.' `
    -Force | Out-Null

if ($OpenFirewall) {
    $rule = Get-NetFirewallRule -DisplayName $firewallRuleName -ErrorAction SilentlyContinue
    if (-not $rule) {
        New-NetFirewallRule -DisplayName $firewallRuleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort $apiPort | Out-Null
    }
}

Start-ScheduledTask -TaskName $TaskName
Write-Host ''
Write-Host 'Task criada e iniciada com sucesso.'
Write-Host "Logs stdout: $(Join-Path $apiRoot 'api_stdout.log')"
Write-Host "Logs stderr: $(Join-Path $apiRoot 'api_stderr.log')"
