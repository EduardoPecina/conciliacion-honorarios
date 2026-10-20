@echo off
REM ==========================================================================
REM  Instalador del Portal de Honorarios: doble clic para INSTALAR o ACTUALIZAR.
REM  1) Busca Python 3.13 (la version con la que el programa esta probado); si no
REM     esta, lo instala solo para este usuario (sin permisos de administrador),
REM     junto a cualquier otro Python que ya tenga la PC (ese no se toca).
REM  2) Lo demas (librerias, carpetas, acceso directo) lo hace instalador\instalar.py
REM  Todo queda registrado en Honorarios\instalacion.log (los pasos de este .bat se
REM  anotan primero en la carpeta temporal y el instalador los copia ahi).
REM ==========================================================================
title Instalador - Portal de Honorarios
cd /d "%~dp0"
set "HONORARIOS_BITACORA_BAT=%TEMP%\InstalacionHonorarios_bat.log"
> "%HONORARIOS_BITACORA_BAT%" echo INSTALAR.bat %DATE% %TIME% desde %CD%
echo ============================================================
echo   INSTALADOR - PORTAL DE HONORARIOS
echo ============================================================
echo.
echo Revisando Python 3.13...
set "PY="
set "PYNUEVO=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
REM Version de respaldo si no hay winget (winget instala la 3.13 mas reciente)
set "PYVERSION=3.13.7"
set "INSTALADOR_PY=%TEMP%\instalador-python.exe"
py -3.13 -c "import sys" >nul 2>nul && set "PY=py -3.13"
if not defined PY python -c "import sys; sys.exit(0 if sys.version_info[:2] == (3, 13) else 1)" >nul 2>nul && set "PY=python"
if not defined PY if exist "%PYNUEVO%" set PY="%PYNUEVO%"
if not defined PY call :instalar_python
if not defined PY goto :sin_python
>> "%HONORARIOS_BITACORA_BAT%" echo Python 3.13 encontrado: %PY%
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
>> "%HONORARIOS_BITACORA_BAT%" echo No se encontro Python 3.13: se instala para el usuario
where winget >nul 2>nul || goto :sin_winget
echo Instalando Python con winget...
winget install --id Python.Python.3.13 -e --scope user --silent --accept-package-agreements --accept-source-agreements >> "%HONORARIOS_BITACORA_BAT%" 2>&1
>> "%HONORARIOS_BITACORA_BAT%" echo winget termino con codigo %errorlevel%
:sin_winget
if exist "%PYNUEVO%" goto :python_listo
echo Descargando Python %PYVERSION% desde python.org...
>> "%HONORARIOS_BITACORA_BAT%" echo Descargando Python %PYVERSION% desde python.org
del /q "%INSTALADOR_PY%" >nul 2>nul
REM --fail: una pagina de error del proxy no se guarda como si fuera el instalador; --proto: solo HTTPS
curl.exe -sS -L --fail --proto =https --retry 5 -o "%INSTALADOR_PY%" https://www.python.org/ftp/python/%PYVERSION%/python-%PYVERSION%-amd64.exe >> "%HONORARIOS_BITACORA_BAT%" 2>&1
>> "%HONORARIOS_BITACORA_BAT%" echo curl termino con codigo %errorlevel%
if not exist "%INSTALADOR_PY%" goto :python_listo
REM Antes de ejecutarlo: firma digital valida de la Python Software Foundation. Si no la tiene, no se ejecuta.
powershell.exe -NoProfile -NonInteractive -Command "$f = Get-AuthenticodeSignature -LiteralPath $env:INSTALADOR_PY; $f.Status; $f.SignerCertificate.Subject; if ($f.Status -ne 'Valid' -or $f.SignerCertificate.Subject -notmatch 'O=Python Software Foundation') { exit 1 }" >> "%HONORARIOS_BITACORA_BAT%" 2>&1
if errorlevel 1 goto :firma_invalida
>> "%HONORARIOS_BITACORA_BAT%" echo Firma de la Python Software Foundation valida
echo Instalando Python...
"%INSTALADOR_PY%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_launcher=1 InstallLauncherAllUsers=0
>> "%HONORARIOS_BITACORA_BAT%" echo Instalador de python.org termino con codigo %errorlevel%
del /q "%INSTALADOR_PY%" >nul 2>nul
:python_listo
if exist "%PYNUEVO%" set PY="%PYNUEVO%"
exit /b 0

:firma_invalida
>> "%HONORARIOS_BITACORA_BAT%" echo El archivo descargado NO tiene la firma valida de la Python Software Foundation: no se ejecuta
echo El archivo descargado no tiene la firma digital de python.org: no se ejecuta.
del /q "%INSTALADOR_PY%" >nul 2>nul
exit /b 0

:sin_python
>> "%HONORARIOS_BITACORA_BAT%" echo No se pudo instalar Python 3.13
echo.
echo No se pudo instalar Python automaticamente.
echo Pide a Sistemas que instale Python 3.13 de 64 bits (python.org) y vuelve a correr INSTALAR.bat.
echo.
echo El detalle quedo en este archivo (se abre su carpeta); mandalo a quien mantiene el programa:
echo   %HONORARIOS_BITACORA_BAT%
explorer /select,"%HONORARIOS_BITACORA_BAT%"
pause
exit /b 1

:fallo
echo.
echo La instalacion no se completo. Revisa el mensaje de arriba.
echo El detalle quedo en el archivo instalacion.log de la carpeta Honorarios de tu Escritorio.
pause
exit /b 1
