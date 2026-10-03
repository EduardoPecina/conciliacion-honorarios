@echo off
REM ==========================================================================
REM  Instalador del Portal de Honorarios: doble clic para INSTALAR o ACTUALIZAR.
REM  1) Busca Python 3.13 (la version con la que el programa esta probado); si no
REM     esta, lo instala solo para este usuario (sin permisos de administrador),
REM     junto a cualquier otro Python que ya tenga la PC (ese no se toca).
REM  2) Lo demas (librerias, carpetas, acceso directo) lo hace instalador\instalar.py
REM ==========================================================================
title Instalador - Portal de Honorarios
cd /d "%~dp0"
echo ============================================================
echo   INSTALADOR - PORTAL DE HONORARIOS
echo ============================================================
echo.
echo Revisando Python 3.13...
set "PY="
set "PYNUEVO=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
py -3.13 -c "import sys" >nul 2>nul && set "PY=py -3.13"
if not defined PY python -c "import sys; sys.exit(0 if sys.version_info[:2] == (3, 13) else 1)" >nul 2>nul && set "PY=python"
if not defined PY if exist "%PYNUEVO%" set PY="%PYNUEVO%"
if not defined PY call :instalar_python
if not defined PY goto :sin_python
echo Python 3.13 encontrado.
echo.
%PY% instalador\instalar.py
if errorlevel 1 goto :fallo
echo.
echo Listo. Puedes cerrar esta ventana.
timeout /t 10 >nul
exit /b 0

:instalar_python
echo No se encontro Python 3.13. Se instalara solo para tu usuario.
echo Esto puede tardar unos minutos; no cierres esta ventana.
where winget >nul 2>nul && winget install --id Python.Python.3.13 -e --scope user --silent --accept-package-agreements --accept-source-agreements
if not exist "%PYNUEVO%" (
    echo Descargando Python desde python.org...
    curl.exe -L --retry 5 -o "%TEMP%\instalador-python.exe" https://www.python.org/ftp/python/3.13.1/python-3.13.1-amd64.exe
    if exist "%TEMP%\instalador-python.exe" "%TEMP%\instalador-python.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_launcher=1 InstallLauncherAllUsers=0
)
if exist "%PYNUEVO%" set PY="%PYNUEVO%"
exit /b 0

:sin_python
echo.
echo No se pudo instalar Python automaticamente.
echo Pide a Sistemas que instale Python 3.13 de 64 bits (python.org) y vuelve a correr INSTALAR.bat.
pause
exit /b 1

:fallo
echo.
echo La instalacion no se completo. Revisa el mensaje de arriba.
pause
exit /b 1
