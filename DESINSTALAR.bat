@echo off
REM ==========================================================================
REM  Desinstalador del Portal de Honorarios (doble clic aqui o en el acceso
REM  directo "Desinstalar Portal de Honorarios" de la carpeta Honorarios).
REM  Copia el desinstalador a la carpeta temporal de Windows y lo corre desde
REM  ahi, porque tiene que borrar _Programa, que es donde vive este archivo.
REM ==========================================================================
set "RAIZ=%~dp0.."
for %%I in ("%RAIZ%") do set "RAIZ=%%~fI"
set "TMPDES=%TEMP%\desinstalar-honorarios"
if exist "%TMPDES%" rmdir /s /q "%TMPDES%"
mkdir "%TMPDES%"
copy /y "%~dp0instalador\instalar.py" "%TMPDES%\" >nul
copy /y "%~dp0instalador\desinstalar.py" "%TMPDES%\" >nul
cd /d "%TMPDES%"
where pyw >nul 2>nul && (start "" pyw -3 "%TMPDES%\desinstalar.py" "%RAIZ%" & exit /b 0)
where pythonw >nul 2>nul && (start "" pythonw "%TMPDES%\desinstalar.py" "%RAIZ%" & exit /b 0)
where python >nul 2>nul && (start "" /min python "%TMPDES%\desinstalar.py" "%RAIZ%" & exit /b 0)
echo No se encontro Python. Borra a mano la carpeta _Programa y los accesos directos.
pause
