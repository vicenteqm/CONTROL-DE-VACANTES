@echo off
cd /d "%~dp0"
where py >nul 2>&1
if errorlevel 1 (
  echo No se detecto Python. Ejecuta primero instalar.bat.
  pause
  exit /b 1
)
py -3 configurar_carpetas.py
if errorlevel 1 pause
