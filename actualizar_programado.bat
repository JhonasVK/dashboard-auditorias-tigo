@echo off
:: Version sin pausas para el Programador de tareas (AuditoriasTIGO_0850):
:: descarga respuestas de Google (Firefox), reemplaza la base y publica el dashboard.
setlocal
set "PATH=%PATH%;C:\Program Files\GitHub CLI"
cd /d "%~dp0"

for /f "delims=" %%i in ('powershell -NoProfile -Command "Get-Date -Format 'dd-MM-yyyy HH:mm'"') do set "FECHAHORA=%%i"
echo.
echo ===== %FECHAHORA% =====

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0descargar_base.ps1"
if errorlevel 1 (
    msg "%USERNAME%" "Auditorias TIGO: no se pudo descargar/reemplazar la base. Revisa log_actualizacion.txt"
    node C:\Bases_Tigo\registrar_estado.js auditorias_tigo ERROR "fallo la descarga desde Google Drive, revisar log_actualizacion.txt"
    exit /b 1
)

:: Otros procesos (Reincidencias) dejan activa otra cuenta de GitHub
gh auth switch --hostname github.com --user JhonasVK >nul 2>nul

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0actualizar_dashboard.ps1"
if errorlevel 1 (
    msg "%USERNAME%" "Auditorias TIGO: se descargo la base pero fallo la publicacion. Revisa log_actualizacion.txt"
    node C:\Bases_Tigo\registrar_estado.js auditorias_tigo ERROR "base descargada pero fallo la publicacion en GitHub"
    exit /b 1
)
node C:\Bases_Tigo\registrar_estado.js auditorias_tigo OK "respuestas descargadas y dashboard TIGO publicado en GitHub"
exit /b 0
