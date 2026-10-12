@echo off
cd /d "%~dp0"
echo ================================================
echo  CONTROL DE VACANTES INVEX - INSTALADOR
echo ================================================
where py >nul 2>&1
if errorlevel 1 (
  echo No se detecto Python. Instala Python 3 desde https://www.python.org/downloads/windows/
  echo Durante la instalacion marca la opcion para agregar Python al PATH.
  pause
  exit /b 1
)
py -3 -m pip install --user -r requirements.txt
if errorlevel 1 (
  echo No fue posible instalar las dependencias.
  pause
  exit /b 1
)
echo Dependencias instaladas correctamente.
pause
