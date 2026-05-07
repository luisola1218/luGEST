param(
    [string]$InstallDir = ""
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms

if (-not $InstallDir) {
    if (Test-Path (Join-Path $env:ProgramFiles 'LuisGEST')) {
        $InstallDir = Join-Path $env:ProgramFiles 'LuisGEST'
    }
    else {
        $InstallDir = Join-Path $env:LOCALAPPDATA 'LuisGEST'
    }
}

$exePath = Join-Path $InstallDir 'main.exe'
if (-not (Test-Path $exePath)) {
    $exePath = Join-Path $InstallDir 'lugest_qt.exe'
}

$internalDir = Join-Path $InstallDir '_internal'
$envPath = Join-Path $InstallDir 'lugest.env'
$logPath = Join-Path $InstallDir 'startup_last.log'
$diagExePath = Join-Path $InstallDir 'main.exe'

function Show-ErrorDialog {
    param(
        [string]$Title,
        [string]$Message
    )
    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

try {
    if (-not (Test-Path $exePath)) {
        throw "Nao foi encontrado lugest_qt.exe nem main.exe em: $InstallDir"
    }
    $exeName = [System.IO.Path]::GetFileName($exePath).ToLowerInvariant()
    if ($exeName -eq 'lugest_qt.exe' -and -not (Test-Path $internalDir)) {
        throw "Falta a pasta _internal em: $InstallDir`nA instalacao ficou incompleta."
    }
    if (-not (Test-Path $envPath)) {
        throw "Falta o ficheiro lugest.env em: $InstallDir"
    }

    $startedAt = Get-Date
    "[$($startedAt.ToString('yyyy-MM-dd HH:mm:ss'))] A tentar arrancar $exePath" | Set-Content -Path $logPath -Encoding UTF8
    $startParams = @{
        FilePath = $exePath
        WorkingDirectory = $InstallDir
        PassThru = $true
    }
    $process = Start-Process @startParams
    Start-Sleep -Seconds 5

    if ($process.HasExited) {
        $finishedAt = Get-Date
        Add-Content -Path $logPath -Value "[$($finishedAt.ToString('yyyy-MM-dd HH:mm:ss'))] O processo fechou logo com exit code $($process.ExitCode)."
        if (Test-Path $diagExePath) {
            Add-Content -Path $logPath -Value "[$($finishedAt.ToString('yyyy-MM-dd HH:mm:ss'))] A correr smoke test de diagnostico via $diagExePath --smoke-test"
            try {
                & $diagExePath --smoke-test *>> $logPath
                $diagExitCode = if ($LASTEXITCODE -is [int]) { $LASTEXITCODE } else { 0 }
                Add-Content -Path $logPath -Value "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] Smoke test terminou com exit code $diagExitCode."
            }
            catch {
                Add-Content -Path $logPath -Value "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] Erro ao correr smoke test: $($_.Exception.Message)"
            }
        }
        throw "A aplicacao arrancou e fechou imediatamente (exit $($process.ExitCode)).`n`nConsulta:`n$logPath`n`nCorre tambem: Diagnosticar Ligacao LuisGEST.bat"
    }

    $runningAt = Get-Date
    Add-Content -Path $logPath -Value "[$($runningAt.ToString('yyyy-MM-dd HH:mm:ss'))] Processo a correr (PID $($process.Id))."
}
catch {
    $message = [string]$_
    try {
        Add-Content -Path $logPath -Value "ERRO: $message"
    }
    catch {
    }
    Show-ErrorDialog -Title 'LuisGEST' -Message $message
    exit 1
}
