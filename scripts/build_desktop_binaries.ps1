$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$venvPython = Join-Path $repoRoot '.venv\Scripts\python.exe'

if (-not (Test-Path $venvPython)) {
    throw "Falta a .venv do projeto: $venvPython"
}

Write-Host "A usar Python de build: $venvPython"

& $venvPython -c "import PySide6, sys; print(sys.executable); print(PySide6.__file__)"
if ($LASTEXITCODE -ne 0) {
    throw "A .venv nao tem PySide6 funcional. Instala primeiro os requisitos Qt."
}

& $venvPython -m pip show PyInstaller *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Host "PyInstaller nao encontrado na .venv. A instalar..."
    & $venvPython -m pip install PyInstaller
    if ($LASTEXITCODE -ne 0) {
        throw "Falha a instalar PyInstaller na .venv."
    }
}

Write-Host "A gerar main.exe..."
& $venvPython -m PyInstaller --noconfirm main.spec --distpath dist --workpath build
if ($LASTEXITCODE -ne 0) {
    throw "Falha a gerar dist\\main.exe"
}

Write-Host "A gerar lugest_qt.exe..."
& $venvPython -m PyInstaller --noconfirm lugest_qt.spec --distpath dist_qt_stable --workpath build_qt_stable
if ($LASTEXITCODE -ne 0) {
    throw "Falha a gerar dist_qt_stable\\lugest_qt\\lugest_qt.exe"
}

Write-Host "Build desktop concluida."
Get-Item (Join-Path $repoRoot 'dist\main.exe') | Select-Object FullName, Length, LastWriteTime
Get-Item (Join-Path $repoRoot 'dist_qt_stable\lugest_qt\lugest_qt.exe') | Select-Object FullName, Length, LastWriteTime
