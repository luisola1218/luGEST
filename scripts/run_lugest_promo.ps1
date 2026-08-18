$ErrorActionPreference = "Stop"

$python = "C:\Users\engenharia\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$generator = "C:\Users\engenharia\VSCodeProjects\teste\scripts\create_lugest_promo.py"
$ffmpeg = "C:\Users\ENGENH~1\AppData\Local\Temp\lugest-video-tools\imageio_ffmpeg\binaries\ffmpeg-win-x86_64-v7.1.exe"
$output = "C:\Users\engenharia\Desktop\luGEST - Pacote Comercial Piloto\Apresentacao Comercial\Video Promocional luGEST PT.mp4"
$work = Join-Path $env:TEMP "lugest-promo-build"
$tools = Join-Path $env:TEMP "lugest-video-tools"

New-Item -ItemType Directory -Force -Path $work | Out-Null
$env:PYTHONPATH = $tools

$stdout = Join-Path $work "build.stdout.log"
$stderr = Join-Path $work "build.stderr.log"
Remove-Item -LiteralPath $stdout, $stderr -Force -ErrorAction SilentlyContinue

$argumentLine = '"{0}" --ffmpeg "{1}" --output "{2}" --work "{3}"' -f $generator, $ffmpeg, $output, $work
$process = Start-Process `
    -FilePath $python `
    -ArgumentList $argumentLine `
    -WorkingDirectory "C:\Users\engenharia\VSCodeProjects\teste" `
    -RedirectStandardOutput $stdout `
    -RedirectStandardError $stderr `
    -WindowStyle Hidden `
    -PassThru

Write-Output $process.Id
