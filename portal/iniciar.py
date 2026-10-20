"""Abre el Portal de Honorarios sin ventana negra: lo arranca en segundo plano y abre el navegador predeterminado.

Lo llama EJECUTAR_PORTAL.bat con pythonw (Python sin consola). Si el portal ya está abierto, solo abre otra
pestaña. Si algo falla, muestra una ventana de Windows con el motivo; el detalle queda en
%LOCALAPPDATA%\\Honorarios\\portal.log (fuera de OneDrive).

Clics de más (usuarios desesperados que dan doble clic varias veces): solo un lanzador a la vez trabaja. Los que
llegan mientras otro arranca el portal (o acaba de abrir la pestaña) se retiran sin hacer nada, así que nunca se
arranca más de un portal ni se abren varias pestañas. El turno lo suelta Windows solo cuando el lanzador termina,
aunque se caiga, y un lanzador nunca vive más de VIDA_MAXIMA: un lanzador atorado no deja el acceso sin funcionar.

Portal trabado (el puerto está abierto pero no contesta): se detienen sus procesos y se arranca uno nuevo.
"""

import http.client
import os
import socket
import subprocess
import sys
import tempfile
import threading
import time
import webbrowser
from pathlib import Path

RAIZ_REPO = Path(__file__).resolve().parents[1]
if str(RAIZ_REPO) not in sys.path:
    sys.path.insert(0, str(RAIZ_REPO))           # para importar portal.apagado al correr como script

PUERTO = 8501                                  # el mismo de .streamlit/config.toml
URL = f"http://localhost:{PUERTO}"
# Fuera de _Programa (que puede estar en el Escritorio de OneDrive), junto a la clave del trabajo guardado
CARPETA_LOCAL = Path(os.environ.get("LOCALAPPDATA") or Path.home() / ".config") / "Honorarios"
BITACORA = CARPETA_LOCAL / "portal.log"
ESPERA_MAXIMA = 120                            # segundos; la primera vez el antivirus puede hacerlo lento
VIDA_MAXIMA = ESPERA_MAXIMA + 90               # pase lo que pase, el lanzador no vive más que esto (suelta el turno)
PAUSA_ENTRE_CLICS = 5                          # tras abrir la pestaña, los clics de más en estos segundos no abren otra
CANDADO = Path(tempfile.gettempdir()) / f"portal-honorarios-{PUERTO}.lock"   # en TEMP, no en OneDrive
# Lo que protege al portal va también en la línea de comandos, que gana sobre .streamlit/config.toml, las variables
# STREAMLIT_* y la configuración global del usuario: aunque falte o no se lea ese archivo, solo esta computadora
# puede abrirlo y nada sale a Streamlit.
OPCIONES_STREAMLIT = (
    "--server.address=localhost", f"--server.port={PUERTO}", "--server.headless=true",
    "--server.enableXsrfProtection=true", "--server.enableCORS=true",
    "--server.allowedHosts=localhost", "--server.allowedHosts=127.0.0.1",
    "--browser.gatherUsageStats=false", "--client.showErrorLinks=false",
    "--server.maxUploadSize=500", "--client.toolbarMode=minimal",
)
# Si una página web abierta en el navegador intenta conectarse al portal desde otro sitio, Streamlit la rechaza, pero
# antes pregunta a checkip.amazonaws.com la IP pública de la PC para compararla. Se le da ya resuelta (127.0.0.1, que
# de todos modos acepta): así no sale nada a internet. Se arranca Streamlit con -c en lugar de -m streamlit.
SIN_IP_EXTERNA = "import streamlit.net_util as red; red._external_ip = '127.0.0.1'"
ARRANQUE = SIN_IP_EXTERNA + "; from streamlit.web.cli import main; main()"


def comando_streamlit():
    """El mismo arranque para el acceso directo y para la consola de diagnóstico (python portal/iniciar.py --consola)."""
    return [sys.executable, "-c", ARRANQUE, "run", str(Path("portal") / "app.py"), *OPCIONES_STREAMLIT]


def portal_abierto():
    try:
        with socket.create_connection(("127.0.0.1", PUERTO), timeout=1):
            return True
    except OSError:
        return False


def portal_responde(intentos=4, tiempo_espera=5):
    """True si el portal contesta (aunque sea "todavía arrancando"). Directo a 127.0.0.1, sin el proxy de la red.
    Se le dan varias oportunidades para no confundir un portal ocupado conciliando con uno trabado."""
    for intento in range(intentos):
        conexion = http.client.HTTPConnection("127.0.0.1", PUERTO, timeout=tiempo_espera)
        try:
            conexion.request("GET", "/_stcore/health")
            conexion.getresponse()
            return True
        except (OSError, http.client.HTTPException):
            if intento < intentos - 1:
                time.sleep(1)
        finally:
            conexion.close()
    return False


def avisar(mensaje, error=False):
    """Ventana de Windows (no hay consola); en otros sistemas, texto."""
    if sys.platform == "win32":
        import ctypes
        icono = 0x10 if error else 0x40        # MB_ICONERROR / MB_ICONINFORMATION
        ctypes.windll.user32.MessageBoxW(None, mensaje, "Portal de Honorarios", icono)
    else:
        print(mensaje)


# ------------------------------------------------------------------ Un solo lanzador a la vez
def tomar_turno():
    """Candado para que solo un lanzador a la vez arranque el portal. Devuelve el candado (se suelta con
    soltar_turno), None si otro lanzador ya tiene el turno, o "sin candado" si no se pudo crear el archivo
    (se sigue como antes, sin protección, en lugar de dejar el acceso sin funcionar).
    El sistema operativo suelta el candado solo cuando el lanzador termina, aunque se caiga."""
    try:
        candado = os.open(CANDADO, os.O_RDWR | os.O_CREAT)
    except OSError:
        return "sin candado"
    try:
        if sys.platform == "win32":
            import msvcrt
            msvcrt.locking(candado, msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(candado, fcntl.LOCK_EX | fcntl.LOCK_NB)
        return candado
    except OSError:
        os.close(candado)
        return None


def soltar_turno(candado):
    if isinstance(candado, int):
        os.close(candado)                      # cerrar el archivo suelta el candado


# ------------------------------------------------------------------ Arranque
def detener_portal_trabado():
    """Detiene los procesos de este portal (nunca otro Python) y espera a que se libere el puerto. No toca a este
    lanzador ni a quien lo lanzó (p. ej. pyw.exe: si se detiene, Windows cierra también este lanzador)."""
    from portal.apagado import detener, procesos_del_portal
    try:
        import psutil
        yo = {os.getpid()} | {padre.pid for padre in psutil.Process().parents()}
    except Exception:
        yo = {os.getpid()}
    detener([proceso for proceso in procesos_del_portal() if proceso.pid not in yo])
    limite = time.monotonic() + 15
    while portal_abierto() and time.monotonic() < limite:
        time.sleep(0.5)
    return not portal_abierto()


def arrancar_portal():
    """Arranca el portal en segundo plano y espera a que abra el puerto. None si quedó listo, o el mensaje de error."""
    BITACORA.parent.mkdir(parents=True, exist_ok=True)
    bitacora = open(BITACORA, "w", encoding="utf-8")
    opciones = {"creationflags": subprocess.CREATE_NO_WINDOW} if sys.platform == "win32" else {}
    proceso = subprocess.Popen(comando_streamlit(), cwd=RAIZ_REPO, stdout=bitacora, stderr=subprocess.STDOUT,
                               stdin=subprocess.DEVNULL, **opciones)
    limite = time.monotonic() + ESPERA_MAXIMA
    try:
        while time.monotonic() < limite:
            if portal_abierto():
                return None
            if proceso.poll() is not None:     # terminó antes de estar listo
                if portal_abierto():           # ...porque otro portal ganó el puerto: ese sirve
                    return None
                break
            time.sleep(0.5)
        else:
            proceso.terminate()
    finally:
        bitacora.close()
    ultimas = "".join(BITACORA.read_text(encoding="utf-8", errors="replace").splitlines(keepends=True)[-8:])
    return ("No se pudo iniciar el Portal de Honorarios.\n\n"
            f"Detalle (últimas líneas de {BITACORA.name}):\n{ultimas or '(vacío)'}\n"
            "Toma captura de este mensaje y compártela con quien mantiene el programa.")


def preparar_portal():
    """Deja el portal listo para abrir la pestaña. None si quedó listo, o el mensaje de error."""
    if portal_abierto():
        if portal_responde():
            return None                        # doble clic con el portal ya abierto: solo otra pestaña
        if not detener_portal_trabado():
            return ("El Portal de Honorarios no responde y no se pudo cerrar solo.\n\n"
                    "Reinicia la computadora y vuelve a abrirlo. Si vuelve a pasar, toma captura de este mensaje "
                    "y compártela con quien mantiene el programa.")
    return arrancar_portal()


def iniciar():
    candado = tomar_turno()
    if candado is None:                        # otro clic ya se está encargando: ese abre la pestaña
        return 0
    vigia = threading.Timer(VIDA_MAXIMA, os._exit, args=(1,))
    vigia.daemon = True
    vigia.start()
    try:
        error = preparar_portal()
        if error is None:
            webbrowser.open(URL)               # navegador predeterminado: pestaña nueva, o lo abre si estaba cerrado
            time.sleep(PAUSA_ENTRE_CLICS)      # se conserva el turno: los clics de más no abren otra pestaña
    finally:
        soltar_turno(candado)
        vigia.cancel()
    if error:
        avisar(error, error=True)              # ya sin el turno: con la ventana abierta, otro clic puede reintentar
        return 1
    return 0


if __name__ == "__main__":
    if "--consola" in sys.argv[1:]:            # diagnóstico: el portal con sus mensajes en esta ventana (Ctrl+C lo cierra)
        sys.exit(subprocess.call(comando_streamlit(), cwd=RAIZ_REPO))
    sys.exit(iniciar())
