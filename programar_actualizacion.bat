@echo off
cd /d "%~dp0"
echo ================================================
echo  PROGRAMAR ACTUALIZACION DIARIA INVEX
echo ================================================
set HORA=09:00
set /p HORA=Hora de actualizacion diaria (formato 24h, predeterminado 09:00):
if "%HORA%"=="" set HORA=09:00
schtasks /Create /TN "Control Vacantes INVEX" /SC DAILY /ST %HORA% /TR "\"%~dp0actualizar.bat\" --scheduled" /F
if errorlevel 1 (
 echo No se pudo crear la tarea. Puedes configurarla manualmente en Programador de tareas de Windows.
) else (
 echo Tarea creada. Windows ejecutara el proceso todos los dias a las %HORA%.
 echo El equipo debe estar encendido, conectado a internet y tener los Excel guardados.
)
pause
