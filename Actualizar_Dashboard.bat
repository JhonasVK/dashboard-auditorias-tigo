@echo off
setlocal
set "PATH=%PATH%;C:\Program Files\GitHub CLI"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0actualizar_dashboard.ps1"
echo.
pause
