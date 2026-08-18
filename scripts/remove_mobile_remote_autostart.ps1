param()

$ErrorActionPreference = 'Stop'
foreach ($taskName in @('LuGEST Mobile Gateway', 'LuGEST Mobile Tunnel')) {
    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($task) {
        Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Host "Removido: $taskName"
    }
}
