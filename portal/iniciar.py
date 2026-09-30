"""Abre el Portal de Honorarios sin ventana negra: lo arranca en segundo plano y abre el navegador predeterminado.

Lo llama EJECUTAR_PORTAL.bat con pythonw (Python sin consola). Si el portal ya está abierto, solo abre otra
pestaña. Si algo falla, muestra una ventana de Windows con el motivo; el detalle queda en portal/portal.log.
"""

import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

RAIZ_REPO = Path(__file__).resolve().parents[1]
PUERTO = 8501                                  # el mismo de .streamlit/config.toml
URL = f"http://localhost:{PUERTO}"
BITACORA = Path(__file__).resolve().parent / "portal.log"
ESPERA_MAXIMA = 120                            # segundos; la primera vez el antivirus puede hacerlo lento


def portal_abierto():
    try:
        with socket.create_connection(("127.0.0.1", PUERTO), timeout=1):
            return True
    except OSError:
        return False


def avisar(mensaje, error=False):
    """Ventana de Windows (no hay consola); en otros sistemas, texto."""
    if sys.platform == "win32":
        import ctypes
        icono = 0x10 if error else 0x40        # MB_ICONERROR / MB_ICONINFORMATION
        ctypes.windll.user32.MessageBoxW(None, mensaje, "Portal de Honorarios", icono)
    else:
        print(mensaje)


def iniciar():
    if portal_abierto():                       # doble clic con el portal ya abierto: solo otra pestaña
        webbrowser.open(URL)
        return 0

    bitacora = open(BITACORA, "w", encoding="utf-8")
    opciones = {"creationflags": subprocess.CREATE_NO_WINDOW} if sys.platform == "win32" else {}
    proceso = subprocess.Popen([sys.executable, "-m", "streamlit", "run", str(Path("portal") / "app.py")],
                               cwd=RAIZ_REPO, stdout=bitacora, stderr=subprocess.STDOUT,
                               stdin=subprocess.DEVNULL, **opciones)
    limite = time.monotonic() + ESPERA_MAXIMA
    while time.monotonic() < limite:
        if portal_abierto():
            webbrowser.open(URL)               # navegador predeterminado: pestaña nueva, o lo abre si estaba cerrado
            return 0
        if proceso.poll() is not None:         # el portal terminó antes de estar listo: hubo un error
            break
        time.sleep(0.5)
    else:
        proceso.terminate()
    bitacora.close()
    ultimas = "".join(BITACORA.read_text(encoding="utf-8", errors="replace").splitlines(keepends=True)[-8:])
    avisar("No se pudo iniciar el Portal de Honorarios.\n\n"
           f"Detalle (últimas líneas de {BITACORA.name}):\n{ultimas or '(vacío)'}\n"
           "Toma captura de este mensaje y compártela con quien mantiene el programa.", error=True)
    return 1


if __name__ == "__main__":
    sys.exit(iniciar())
