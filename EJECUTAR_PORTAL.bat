@echo off
REM Portal de Honorarios: doble clic para abrirlo. NO cierres la ventana negra mientras lo uses.
cd /d "%~dp0"
set PY=python
where python >nul 2>nul || set PY=py
echo ============================================================
echo   PORTAL DE HONORARIOS
echo ============================================================
echo Iniciando... el navegador se abrira solo en cuanto el portal este listo.
echo NO CIERRES ESTA VENTANA mientras uses el portal.
echo.
REM Abre el navegador hasta que el portal escuche en el puerto 8501 (espera hasta 2 minutos).
start "" /b powershell -NoProfile -WindowStyle Hidden -Command "for($i=0;$i -lt 120;$i++){try{(New-Object Net.Sockets.TcpClient).Connect('127.0.0.1',8501); Start-Process 'http://localhost:8501'; break}catch{Start-Sleep 1}}"
%PY% -m streamlit run portal\app.py
echo.
echo El portal se detuvo. Si arriba aparece un error, toma captura y compartela.
pause
