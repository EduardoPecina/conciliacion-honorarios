@echo off
REM Conciliacion de Honorarios: doble clic para ejecutar.
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (
    python conciliar.py
) else (
    py conciliar.py
)
