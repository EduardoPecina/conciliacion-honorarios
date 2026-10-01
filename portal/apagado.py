"""Apagado del portal. Como corre en segundo plano (sin ventana negra que cerrar), se apaga:
- con el botón "Cerrar el portal" del menú, que detiene TODO lo del portal (como Ctrl + C): este portal, portales
  viejos que se hayan quedado abiertos y el lanzador; o
- solo, si pasan INACTIVIDAD minutos sin ninguna pestaña del portal abierta.
Así no queda un portal olvidado ocupando el puerto y bloqueando archivos de _Programa.

Los procesos del portal se reconocen por su línea de comando: Python corriendo portal/app.py o portal/iniciar.py
de ESTA instalación. Cualquier otro Python (por ejemplo, el de VS Code) no se toca."""

import os
import threading
import time
from pathlib import Path

PORTAL = Path(__file__).resolve().parent
SCRIPTS_DEL_PORTAL = {os.path.normcase(str(PORTAL / nombre)) for nombre in ("app.py", "iniciar.py")}

INACTIVIDAD = int(os.environ.get("HONORARIOS_PORTAL_INACTIVIDAD_MIN", "15")) * 60     # segundos
REVISION = min(30, max(INACTIVIDAD // 2, 1))                                          # cada cuánto se revisa


def pestanas_abiertas():
    """Pestañas del navegador conectadas al portal (None si esta versión de Streamlit no lo permite saber)."""
    try:
        from streamlit.runtime import Runtime
        return Runtime.instance()._session_mgr.num_active_sessions()
    except Exception:
        return None


def _vigilar():
    sin_pestanas = 0
    while True:
        time.sleep(REVISION)
        abiertas = pestanas_abiertas()
        if abiertas is None:
            return                               # no se puede saber: no se apaga solo (queda el botón)
        sin_pestanas = sin_pestanas + REVISION if abiertas == 0 else 0
        if sin_pestanas >= INACTIVIDAD:
            os._exit(0)


def vigilar_inactividad():
    threading.Thread(target=_vigilar, name="apagado-por-inactividad", daemon=True).start()


def es_del_portal(linea_de_comando, carpeta):
    """True si el proceso es Python corriendo portal/app.py o portal/iniciar.py de esta instalación."""
    if not linea_de_comando or not Path(linea_de_comando[0]).name.lower().startswith("py"):
        return False                                    # python.exe, pythonw.exe, py.exe, pyw.exe
    for argumento in linea_de_comando[1:]:
        if argumento.lower().endswith((".py", ".pyw")):
            ruta = Path(argumento) if os.path.isabs(argumento) else Path(carpeta) / argumento
            try:
                if os.path.normcase(str(ruta.resolve())) in SCRIPTS_DEL_PORTAL:
                    return True
            except OSError:
                pass
    return False


def procesos_del_portal():
    """Otros procesos de este portal (no el actual). Lista vacía si no se puede saber (sin psutil)."""
    try:
        import psutil
    except ImportError:
        return []
    propios = []
    for proceso in psutil.process_iter():
        if proceso.pid == os.getpid():
            continue
        try:
            if es_del_portal(proceso.cmdline(), proceso.cwd()):
                propios.append(proceso)
        except (psutil.Error, OSError):                 # terminó o es de otro usuario: no es nuestro
            continue
    return propios


def detener(procesos, espera=3):
    """Detiene los procesos y sus hijos: primero les pide terminar y, si no responden, los fuerza."""
    try:
        import psutil
    except ImportError:
        return
    todos = []
    for proceso in procesos:
        try:
            todos += proceso.children(recursive=True) + [proceso]
        except psutil.Error:
            continue
    for proceso in todos:
        try:
            proceso.terminate()
        except psutil.Error:
            pass
    _, vivos = psutil.wait_procs(todos, timeout=espera)
    for proceso in vivos:
        try:
            proceso.kill()
        except psutil.Error:
            pass


def cerrar_portal(espera=1.5):
    """Detiene todo lo del portal (como Ctrl + C): los otros procesos del portal y los hijos de este, y apaga este
    después de unos segundos (para que alcance a mostrarse la página de "El portal se cerró")."""
    try:
        import psutil
        detener(procesos_del_portal() + psutil.Process().children(recursive=True))
    except Exception:
        pass                                            # aunque falle la limpieza, este portal sí se apaga
    threading.Timer(espera, os._exit, args=(0,)).start()
