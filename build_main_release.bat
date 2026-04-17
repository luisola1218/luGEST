@echo off
cd /d %~dp0
py -m PyInstaller --noconfirm main.spec
