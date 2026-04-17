param(
    [string]$TaskName = 'LUGEST Impulse API',
    [switch]$RemoveFirewall,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$apiRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

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

$apiPort = Get-DotEnvValue -Key 'LUGEST_API_PORT' -DefaultValue '8050'
$firewallRuleNames = @(
    "LUGEST Impulse API $apiPort",
    'LUGEST Impulse API 8050'
) | Select-Object -Unique

Write-Host "Task.......: $TaskName"
Write-Host "Porta API..: $apiPort"

if ($DryRun) {
    Write-Host ''
    Write-Host 'DryRun ativo. Nenhuma alteracao foi aplicada.'
    return
}

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

if ($RemoveFirewall) {
    foreach ($ruleName in $firewallRuleNames) {
        Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue |
            Remove-NetFirewallRule -ErrorAction SilentlyContinue
    }
}

Write-Host ''
Write-Host 'Task removida com sucesso.'
