"""Genera requirements-instalar.txt: las versiones EXACTAS de requirements-lock.txt con la huella SHA-256 de cada
instalador (wheel) publicado en PyPI. El instalador lo usa con --require-hashes: pip comprueba que cada archivo
descargado sea, byte por byte, el que se publicó para esa versión; si alguien lo alterara en el camino (o PyPI
sirviera otro), la instalación se detiene.

Solo lo corre quien mantiene el programa, cuando cambia requirements-lock.txt (ver README, "Versiones fijas"):

    python instalador/generar_lock.py
"""

import json
import sys
import urllib.request
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
LOCK = RAIZ / "requirements-lock.txt"
SALIDA = RAIZ / "requirements-instalar.txt"
# Solo se instalan en Windows (en Linux no existen o no hacen falta): así el archivo también se puede comprobar en Linux
SOLO_WINDOWS = {"pywin32", "colorama"}


def versiones_fijadas(lock=LOCK):
    lineas = lock.read_text(encoding="utf-8").splitlines()
    return [tuple(l.strip().split("==")) for l in lineas if l.strip() and not l.lstrip().startswith("#")]


def huellas(nombre, version):
    """SHA-256 de todos los wheels de esa versión publicados en PyPI."""
    with urllib.request.urlopen(f"https://pypi.org/pypi/{nombre}/{version}/json", timeout=60) as respuesta:
        datos = json.load(respuesta)
    return sorted(u["digests"]["sha256"] for u in datos["urls"] if u["filename"].endswith(".whl"))


def generar():
    partes = ["# GENERADO por instalador/generar_lock.py a partir de requirements-lock.txt. No editar a mano.",
              "# El instalador lo usa con --require-hashes: cada archivo descargado debe coincidir con su huella SHA-256.",
              ""]
    for nombre, version in versiones_fijadas():
        lista = huellas(nombre, version)
        if not lista:
            sys.exit(f"{nombre}=={version} no tiene instaladores (wheels) en PyPI")
        marcador = ' ; sys_platform == "win32"' if nombre in SOLO_WINDOWS else ""
        partes.append(f"{nombre}=={version}{marcador} \\")
        partes += [f"    --hash=sha256:{h}" + (" \\" if i < len(lista) - 1 else "") for i, h in enumerate(lista)]
    SALIDA.write_text("\n".join(partes) + "\n", encoding="utf-8")
    print(f"{SALIDA.name}: {len(versiones_fijadas())} librerías con sus huellas")


if __name__ == "__main__":
    generar()
