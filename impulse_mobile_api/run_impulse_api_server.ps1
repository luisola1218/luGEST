param(
    [int]$RestartDelaySeconds = 5,
    [int]$MaxRestarts = 50
)

$ErrorActionPreference = 'Stop'

$apiRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExe = Join-Path $apiRoot '.venv\Scripts\python.exe'
$verifyScript = Join-Path $apiRoot 'verify_installation.py'
$runScript = Join-Path $apiRoot 'run_server.py'
$stdoutLog = Join-Path $apiRoot 'api_stdout.log'
$stderrLog = Join-Path $apiRoot 'api_stderr.log'
$pidFile = Join-Path $apiRoot '.server.pid'

function Write-LogLine {
    param(
        [string]$Message,
        [string]$Target = 'stdout'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] $Message"
    if ($Target -eq 'stderr') {
        Add-Content -Path $stderrLog -Value $line -Encoding UTF8
        return
    }
    Add-Content -Path $stdoutLog -Value $line -Encoding UTF8
}

if (-not (Test-Path $pythonExe)) {
    throw "Falta o Python da venv: $pythonExe"
}
if (-not (Test-Path $verifyScript)) {
    throw "Falta o preflight da API: $verifyScript"
}
if (-not (Test-Path $runScript)) {
    throw "Falta o arranque da API: $runScript"
}

Push-Location $apiRoot
try {
    Set-Content -Path $pidFile -Value $PID -Encoding ASCII
    Write-LogLine "Runner iniciado. PowerShell PID=$PID"

    $restartCount = 0
    while ($true) {
        Write-LogLine "A executar preflight da API."
        & $pythonExe $verifyScript 1>> $stdoutLog 2>> $stderrLog
        $verifyExit = $LASTEXITCODE
        if ($verifyExit -ne 0) {
            Write-LogLine "Preflight falhou com exit code $verifyExit." 'stderr'
            exit $verifyExit
        }

        Write-LogLine "A arrancar uvicorn."
        & $pythonExe $runScript 1>> $stdoutLog 2>> $stderrLog
        $runExit = $LASTEXITCODE
        Write-LogLine "Uvicorn terminou com exit code $runExit."

        if ($runExit -eq 0) {
            break
        }

        $restartCount += 1
        if ($restartCount -ge [Math]::Max(1, $MaxRestarts)) {
            Write-LogLine "Numero maximo de reinicios atingido ($restartCount)." 'stderr'
            exit $runExit
        }

        Write-LogLine "A aguardar $RestartDelaySeconds segundo(s) antes do reinicio $restartCount."
        Start-Sleep -Seconds ([Math]::Max(1, $RestartDelaySeconds))
    }
}
finally {
    if (Test-Path $pidFile) {
        Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
    }
    Write-LogLine "Runner terminado."
    Pop-Location
}
