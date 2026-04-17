@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0apply_lugest_migrations.ps1" -Status %*
