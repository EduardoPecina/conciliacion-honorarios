@echo off
REM Portal de Honorarios: doble clic para abrirlo en tu navegador. No abre ninguna ventana negra:
REM el portal corre en segundo plano y se cierra con el boton "Cerrar el portal" (o solo, si nadie lo usa).
cd /d "%~dp0"
where pythonw >nul 2>nul && (start "" pythonw portal\iniciar.py & exit /b)
where pyw >nul 2>nul && (start "" pyw portal\iniciar.py & exit /b)
REM Sin pythonw/pyw: se usa python normal, minimizado.
start "" /min python portal\iniciar.py
