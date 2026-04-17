@echo off
cd /d %~dp0
python -m PyInstaller --noconfirm lugest_qt.spec --distpath dist_qt_stable --workpath build_qt_stable
