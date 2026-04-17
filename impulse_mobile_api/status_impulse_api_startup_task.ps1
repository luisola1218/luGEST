param(
    [string]$TaskName = 'LUGEST Impulse API',
    [int]$LogTail = 20
)

$ErrorActionPreference = 'Stop'

$apiRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$stdoutLog = Join-Path $apiRoot 'api_stdout.log'
$stderrLog = Join-Path $apiRoot 'api_stderr.log'
$pidFile = Join-Path $apiRoot '.server.pid'

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

$apiPort = 0
try {
    $apiPort = [int](Get-DotEnvValue -Key 'LUGEST_API_PORT' -DefaultValue '8050')
}
catch {
    $apiPort = 8050
}

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if (-not $task) {
    Write-Host "Task nao encontrada: $TaskName"
}
else {
    $info = Get-ScheduledTaskInfo -TaskName $TaskName
    Write-Host "Task........: $TaskName"
    Write-Host "Estado......: $($task.State)"
    Write-Host "Ultima exec.: $($info.LastRunTime)"
    Write-Host "Ultimo code.: $($info.LastTaskResult)"
    Write-Host "Proxima exec: $($info.NextRunTime)"
}

if (Test-Path $pidFile) {
    Write-Host "Runner PID..: $(Get-Content $pidFile -ErrorAction SilentlyContinue)"
}

$listeners = Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue | Where-Object { $_.LocalPort -eq $apiPort }
if ($listeners) {
    foreach ($listener in $listeners) {
        Write-Host "Porta ativa.: $($listener.LocalAddress):$($listener.LocalPort)"
    }
}
else {
    Write-Host "Porta ativa.: nenhuma escuta detetada em $apiPort"
}

if (Test-Path $stdoutLog) {
    Write-Host ''
    Write-Host '--- Ultimas linhas stdout ---'
    Get-Content $stdoutLog -Tail ([Math]::Max(1, $LogTail))
}

if (Test-Path $stderrLog) {
    Write-Host ''
    Write-Host '--- Ultimas linhas stderr ---'
    Get-Content $stderrLog -Tail ([Math]::Max(1, $LogTail))
}
