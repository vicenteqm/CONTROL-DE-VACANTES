@echo off
cd /d "%~dp0"
where py >nul 2>&1
if errorlevel 1 (
  echo No se detecto Python. Ejecuta instalar.bat.
  if /I not "%~1"=="--scheduled" pause
  exit /b 1
)
py -3 actualizar_dashboard.py %*
if errorlevel 1 (
  echo.
  echo La actualizacion fallo. Revisa logs\actualizacion.log.
  if /I not "%~1"=="--scheduled" pause
  exit /b 1
)
if /I not "%~1"=="--scheduled" pause
