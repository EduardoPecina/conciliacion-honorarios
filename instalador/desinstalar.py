"""Desinstalador del Portal de Honorarios. Lo llama DESINSTALAR.bat (o el acceso directo "Desinstalar Portal de
Honorarios"), que antes lo copia junto con instalar.py a la carpeta temporal de Windows: así puede borrar
_Programa, que es donde vive.

Borra SOLO lo que puso el instalador: _Programa, los accesos directos y LEEME.txt. El catálogo, el historial de
UUID y los reportes se conservan, salvo que el usuario elija borrarlos (con doble confirmación). Nunca borra
completa la carpeta Honorarios: si al final queda vacía, la quita; si el usuario guardó otras cosas, se quedan.
Python y las librerías no se tocan (las pueden usar otros programas y así reinstalar es más rápido).

Uso:  python desinstalar.py "<carpeta Honorarios>"
"""

import shutil
import sys
import time
from pathlib import Path

from instalar import NOMBRE_ACCESO, NOMBRE_DESINSTALAR, avisar, escritorio, portal_abierto

TITULO = "Desinstalar - Portal de Honorarios"
SI, NO, CANCELAR = "si", "no", "cancelar"


def preguntar(mensaje, con_cancelar=False):
    """Ventana de Windows con Sí/No (o Sí/No/Cancelar). En otros sistemas, por consola."""
    if sys.platform == "win32":
        import ctypes
        botones = 0x3 if con_cancelar else 0x4                    # MB_YESNOCANCEL / MB_YESNO
        respuesta = ctypes.windll.user32.MessageBoxW(None, mensaje, TITULO, botones | 0x20)   # MB_ICONQUESTION
        return {6: SI, 7: NO}.get(respuesta, CANCELAR)
    texto = input(f"{mensaje}\n[s = sí / n = no{' / c = cancelar' if con_cancelar else ''}]: ").strip().lower()
    return {"s": SI, "si": SI, "sí": SI, "n": NO, "no": NO}.get(texto, CANCELAR if con_cancelar else NO)


def borrar(ruta, reintentos=5):
    """Borra un archivo o carpeta; reintenta un momento (OneDrive o el antivirus pueden tenerlo ocupado)."""
    for intento in range(reintentos):
        try:
            if ruta.is_dir():
                shutil.rmtree(ruta)
            elif ruta.exists():
                ruta.unlink()
            return True
        except OSError:
            time.sleep(1)
    return not ruta.exists()


def desinstalar(raiz, preguntar=preguntar):
    raiz = Path(raiz)
    if not (raiz / "_Programa").exists() and not (raiz / "_Configuracion").exists():
        avisar(f"No se encontró el Portal de Honorarios en:\n{raiz}", TITULO, error=True)
        return 1
    if portal_abierto():
        avisar("El Portal de Honorarios está abierto.\n\nCiérralo (botón \"Cerrar el portal\") y vuelve a "
               "intentar.", TITULO, error=True)
        return 1
    if preguntar("¿Desinstalar el Portal de Honorarios de esta computadora?\n\n"
                 "Se borrarán el programa y sus accesos directos.") != SI:
        return 0

    conservar = preguntar("¿Conservar tu catálogo de proveedores, el historial de UUID y los reportes?\n\n"
                          "Sí (recomendado): se quedan en la carpeta Honorarios y, si reinstalas, el portal los "
                          "vuelve a usar.\n"
                          "No: se borran también.\n"
                          "Cancelar: no desinstalar.", con_cancelar=True)
    if conservar == CANCELAR:
        return 0
    borrar_datos = conservar == NO
    if borrar_datos and preguntar("Se borrarán el catálogo, el historial de UUID (el que evita pagar dos veces "
                                  "la misma factura) y los reportes.\n\nEsto no se puede deshacer. ¿Seguro?") != SI:
        return 0

    time.sleep(1)                                   # que termine DESINSTALAR.bat (vive en _Programa)
    objetivos = [raiz / "_Programa", raiz / "LEEME.txt", raiz / f"{NOMBRE_ACCESO}.lnk",
                 raiz / f"{NOMBRE_DESINSTALAR}.lnk", escritorio() / f"{NOMBRE_ACCESO}.lnk"]
    if borrar_datos:
        objetivos += [raiz / "_Configuracion", raiz / "Reportes"]
    no_borrados = [ruta for ruta in objetivos if not borrar(ruta)]
    if raiz.exists() and not any(raiz.iterdir()):
        raiz.rmdir()                                # quedó vacía: se quita

    if no_borrados:
        avisar("La desinstalación no se completó: Windows no dejó borrar\n\n"
               + "\n".join(str(r) for r in no_borrados)
               + "\n\nCierra lo que tengas abierto de esa carpeta (Excel, el Explorador, OneDrive sincronizando) "
                 "y vuelve a correr el desinstalador.", TITULO, error=True)
        return 1
    if borrar_datos:
        avisar("El Portal de Honorarios se desinstaló por completo.", TITULO)
    else:
        avisar("El Portal de Honorarios se desinstaló.\n\nTu catálogo, el historial de UUID y los reportes se "
               f"conservaron en:\n{raiz}\n\nPara volver a instalarlo, corre INSTALAR.bat.", TITULO)
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(desinstalar(sys.argv[1]))
