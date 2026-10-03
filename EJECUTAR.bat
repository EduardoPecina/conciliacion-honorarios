@echo off
REM Conciliacion de Honorarios: doble clic para ejecutar. Usa Python 3.13 (la version probada) si esta.
cd /d "%~dp0"
py -3.13 -c "import sys" >nul 2>nul && (py -3.13 _Programa.py & exit /b)
where python >nul 2>nul
if %errorlevel%==0 (
    python _Programa.py
) else (
    py _Programa.py
)
