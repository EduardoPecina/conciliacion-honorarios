"""Apagado del portal. Como corre en segundo plano (sin ventana negra que cerrar), se apaga:
- con el botón "Cerrar el portal" del menú, o
- solo, si pasan INACTIVIDAD minutos sin ninguna pestaña del portal abierta.
Así no queda un portal olvidado ocupando el puerto y bloqueando archivos de _Programa."""

import os
import threading
import time

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


def cerrar_portal(espera=1.5):
    """Apaga el portal después de unos segundos (para que alcance a mostrarse el mensaje)."""
    threading.Timer(espera, os._exit, args=(0,)).start()
