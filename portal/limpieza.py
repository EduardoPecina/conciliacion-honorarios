"""Limpieza al abrir el portal: borra las copias de seguridad que deja una actualización (_Programa_anterior) si el
instalador no pudo borrarlas en su momento (OneDrive sincronizando, antivirus). Así el usuario no ve carpetas raras.

Solo se borra una carpeta si es seguro que es una copia vieja del programa: está junto a _Programa, se llama
_Programa_anterior... y todo lo que contiene son nombres que también tiene el programa actual."""

import os
import shutil
import stat
import sys
import threading
import time
from pathlib import Path

PROGRAMA = Path(__file__).resolve().parents[1]


def es_copia_del_programa(carpeta, programa=PROGRAMA):
    """True si todo lo que hay en la carpeta también existe en el programa actual (aunque esté a medio borrar)."""
    actuales = {p.name.lower() for p in programa.iterdir()}
    try:
        return all(p.name.lower() in actuales for p in carpeta.iterdir())
    except OSError:
        return False


def respaldos_viejos(programa=PROGRAMA):
    if programa.name != "_Programa":                     # en una carpeta de desarrollo no se toca nada
        return []
    return [p for p in sorted(programa.parent.glob("_Programa_anterior*"))
            if p.is_dir() and es_copia_del_programa(p, programa)]


def borrar(carpeta, intentos=5, espera=2.0):
    def quitar_solo_lectura(funcion, camino, _):
        try:
            os.chmod(camino, stat.S_IWRITE)
            funcion(camino)
        except OSError:
            pass
    opcion = {"onexc": quitar_solo_lectura} if sys.version_info >= (3, 12) else {"onerror": quitar_solo_lectura}
    for _ in range(intentos):
        shutil.rmtree(carpeta, **opcion)
        if not carpeta.exists():
            return True
        time.sleep(espera)
    return False


def limpiar_respaldos(programa=PROGRAMA, en_segundo_plano=True):
    """Borra las copias viejas. En segundo plano para no retrasar la apertura del portal."""
    viejos = respaldos_viejos(programa)
    if not viejos:
        return
    trabajo = lambda: [borrar(v) for v in viejos]
    if en_segundo_plano:
        threading.Thread(target=trabajo, name="limpieza-respaldos", daemon=True).start()
    else:
        trabajo()
