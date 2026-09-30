@echo off
REM Portal de Honorarios: abre http://localhost:8501 en el navegador. Deja esta ventana abierta.
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (
    start "" http://localhost:8501
    python -m streamlit run portal\app.py
) else (
    start "" http://localhost:8501
    py -m streamlit run portal\app.py
)
