$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$venvPython = Join-Path $repoRoot '.venv\Scripts\python.exe'

if (-not (Test-Path $venvPython)) {
    throw "Falta a .venv oficial. Cria com: py -m venv .venv"
}

Push-Location $repoRoot
try {
    @'
import importlib
import sys

required = ["PySide6", "pypdf", "reportlab", "pymysql"]
missing = []
for name in required:
    try:
        importlib.import_module(name)
    except Exception as exc:
        missing.append(f"{name}: {exc}")
if missing:
    print("Dependencias em falta na .venv:")
    print("\n".join(missing))
    sys.exit(1)
print("dependencias-ok")
'@ | & $venvPython -

    & $venvPython -m PyInstaller lugest_qt.spec --noconfirm
    & powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot 'scripts\prepare_final_release.ps1')
}
finally {
    Pop-Location
}
