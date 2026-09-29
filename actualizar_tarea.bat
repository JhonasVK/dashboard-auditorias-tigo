@echo off
:: Lo ejecuta el Programador de tareas (AuditoriasTIGO_0850). Deja todo en log_actualizacion.txt
cd /d "%~dp0"
call actualizar_programado.bat >> log_actualizacion.txt 2>&1
