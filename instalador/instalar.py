"""Instalador del Portal de Honorarios. Lo llama INSTALAR.bat (que antes se asegura de que haya Python).

Se puede correr las veces que sea: la primera vez INSTALA y después ACTUALIZA (reemplaza solo el programa;
el catálogo, el historial y los reportes no se tocan).

    1. Instala las librerías que faltan (pandas, openpyxl, streamlit, pywin32...).
    2. Crea en el Escritorio del usuario (el de OneDrive, si Windows lo tiene ahí) la carpeta:
           Honorarios\\
             LEEME.txt          instrucciones cortas
             _Configuracion\\    aquí va el Excel del catálogo de proveedores
             Reportes\\          los Excel de cada conciliación (el portal crea la carpeta de cada mes)
             _Programa\\         el programa
    3. Crea el acceso directo "Portal de Honorarios" en el Escritorio y dentro de Honorarios.

Uso manual (pruebas):  python instalador/instalar.py [--sin-librerias] [--sin-accesos]
"""

import argparse
import os
import shutil
import socket
import subprocess
import sys
from pathlib import Path

ORIGEN = Path(__file__).resolve().parents[1]              # la carpeta del ZIP (o del repositorio)
NOMBRE_CARPETA = "Honorarios"
NOMBRE_ACCESO = "Portal de Honorarios"
CATALOGO = "CatalogoProveedoresTM26.xlsx"
NO_COPIAR = shutil.ignore_patterns(".git", ".github", "__pycache__", "*.pyc", "portal.log", "*.tmp")

LEEME = f"""PORTAL DE HONORARIOS
====================

ABRIR EL PORTAL
  Doble clic en el acceso directo "{NOMBRE_ACCESO}" (en el Escritorio o en esta carpeta).
  Se abre en tu navegador. Para cerrarlo: botón "Cerrar el portal" en el menú de la izquierda.

ANTES DE LA PRIMERA CONCILIACIÓN
  Copia el Excel del catálogo de proveedores ({CATALOGO}) en la carpeta _Configuracion.
  El catálogo lo mantienes tú: cada médico nuevo se da de alta en ese Excel (tabla "Catalogo").

CARPETAS
  _Configuracion  catálogo de proveedores e historial de facturas ya pagadas (HistorialUUID.csv).
                  No borres esta carpeta: el historial evita pagar dos veces la misma factura.
  Reportes        los Excel de cada conciliación, por año y mes.
  _Programa       el programa. No modifiques nada aquí.

ACTUALIZAR
  Cierra el portal y vuelve a correr INSTALAR.bat de la versión nueva. Tu catálogo, tu historial y tus
  reportes no se tocan.
"""


# ------------------------------------------------------------------ Windows
def avisar(mensaje, titulo="Instalador - Portal de Honorarios", error=False):
    print(("ERROR: " if error else "") + mensaje)
    if sys.platform == "win32":
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, mensaje, titulo, 0x10 if error else 0x40)


def escritorio():
    """Escritorio del usuario tal como lo tiene Windows (si está en OneDrive, el de OneDrive)."""
    if os.environ.get("HONORARIOS_ESCRITORIO"):              # para pruebas
        return Path(os.environ["HONORARIOS_ESCRITORIO"])
    if sys.platform == "win32":
        import ctypes
        import uuid
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD), ("Data3", wintypes.WORD),
                        ("Data4", ctypes.c_ubyte * 8)]

        guid = uuid.UUID("{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}")          # FOLDERID_Desktop
        folderid = GUID(guid.fields[0], guid.fields[1], guid.fields[2],
                        (ctypes.c_ubyte * 8).from_buffer_copy(guid.bytes[8:]))
        ruta = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(folderid), 0, None, ctypes.byref(ruta)) == 0:
            try:
                return Path(ruta.value)
            finally:
                ctypes.windll.ole32.CoTaskMemFree(ruta)
    return Path.home() / "Desktop"


def portal_abierto():
    try:
        with socket.create_connection(("127.0.0.1", 8501), timeout=1):
            return True
    except OSError:
        return False


# ------------------------------------------------------------------ Pasos
def instalar_librerias():
    """pip install de requirements-portal.txt (si ya están, solo lo confirma). Reintenta por la red corporativa."""
    print("Revisando librerías (pandas, openpyxl, streamlit, pywin32)...")
    comando = [sys.executable, "-m", "pip", "install", "-r", str(ORIGEN / "requirements-portal.txt"),
               "--disable-pip-version-check", "--retries", "10", "--timeout", "60"]
    for intento in (1, 2):
        if subprocess.run(comando).returncode == 0:
            return True
        print(f"La instalación de librerías falló (intento {intento} de 2). Reintentando...")
    return False


def copiar_programa(destino):
    """Copia el programa a Honorarios\\_Programa. Al actualizar, la versión anterior se reemplaza completa
    (así no quedan archivos viejos). Si el origen ya es el destino (se corrió desde _Programa), no copia."""
    if destino.exists() and destino.resolve() == ORIGEN:
        print("El programa ya está en su lugar (se corrió desde _Programa).")
        return
    anterior = destino.with_name(destino.name + "_anterior")
    if anterior.exists():
        shutil.rmtree(anterior, ignore_errors=True)
    if destino.exists():
        destino.rename(anterior)             # si el portal o un archivo está abierto, Windows no deja: PermissionError
    try:
        shutil.copytree(ORIGEN, destino, ignore=NO_COPIAR)
    except Exception:
        if anterior.exists() and not destino.exists():
            anterior.rename(destino)         # algo falló: se deja la versión anterior como estaba
        raise
    shutil.rmtree(anterior, ignore_errors=True)


def crear_estructura(raiz):
    for carpeta in (raiz / "_Configuracion", raiz / "Reportes"):
        carpeta.mkdir(parents=True, exist_ok=True)
    (raiz / "LEEME.txt").write_text(LEEME, encoding="utf-8-sig")


def crear_icono(programa):
    """Convierte el logo en ícono (.ico) para el acceso directo. Si no se puede, el acceso usa el de Python."""
    logo = programa / "portal" / "recursos" / "logo.png"
    icono = programa / "portal" / "recursos" / "icono.ico"
    try:
        from PIL import Image
        imagen = Image.open(logo).convert("RGBA")
        lado = max(imagen.size)                                  # cuadrado, centrado, fondo transparente
        cuadro = Image.new("RGBA", (lado, lado), (0, 0, 0, 0))
        cuadro.paste(imagen, ((lado - imagen.width) // 2, (lado - imagen.height) // 2))
        cuadro.save(icono, sizes=[(16, 16), (32, 32), (48, 48), (256, 256)])
        return icono
    except Exception:
        return None


def crear_accesos(raiz, programa):
    """Acceso directo con pythonw (sin ventana negra). Se crea en otro proceso de Python porque pywin32 se
    acaba de instalar y en este proceso todavía no se puede importar."""
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    if pythonw.exists():
        destino, argumentos = pythonw, '"portal\\iniciar.py"'
    else:                                                        # sin pythonw: el .bat (lo hace minimizado)
        destino, argumentos = programa / "EJECUTAR_PORTAL.bat", ""
    icono = crear_icono(programa)
    codigo = (
        "import sys, win32com.client\n"
        "shell = win32com.client.Dispatch('WScript.Shell')\n"
        "for ruta in sys.argv[5:]:\n"
        "    acceso = shell.CreateShortcut(ruta)\n"
        "    acceso.TargetPath, acceso.Arguments, acceso.WorkingDirectory = sys.argv[1], sys.argv[2], sys.argv[3]\n"
        "    if sys.argv[4]: acceso.IconLocation = sys.argv[4]\n"
        "    acceso.Description = 'Abre el Portal de Honorarios en el navegador'\n"
        "    acceso.Save()\n")
    rutas = [str(raiz.parent / f"{NOMBRE_ACCESO}.lnk"), str(raiz / f"{NOMBRE_ACCESO}.lnk")]
    resultado = subprocess.run([sys.executable, "-c", codigo, str(destino), argumentos, str(programa),
                                str(icono or ""), *rutas])
    return resultado.returncode == 0


def instalar(sin_librerias=False, sin_accesos=False):
    raiz = escritorio() / NOMBRE_CARPETA
    programa = raiz / "_Programa"
    print(f"Instalando en: {raiz}")

    if portal_abierto():
        avisar("El Portal de Honorarios está abierto.\n\nCiérralo (botón \"Cerrar el portal\") y vuelve a "
               "correr INSTALAR.bat.", error=True)
        return 1
    if not sin_librerias and not instalar_librerias():
        avisar("No se pudieron instalar las librerías de Python (¿sin internet o la red las bloquea?).\n\n"
               "Vuelve a intentarlo; si sigue fallando, toma captura de la ventana y compártela con quien "
               "mantiene el programa.", error=True)
        return 1
    actualizacion = programa.exists()
    try:
        crear_estructura(raiz)
        copiar_programa(programa)
    except PermissionError:
        avisar("Windows no dejó reemplazar la carpeta _Programa: algún archivo está abierto (el portal, un "
               "Excel, una ventana de esa carpeta u OneDrive sincronizando).\n\nCiérralos y vuelve a correr "
               "INSTALAR.bat.", error=True)
        return 1
    accesos = sin_accesos or sys.platform != "win32" or crear_accesos(raiz, programa)

    falta_catalogo = not (raiz / "_Configuracion" / CATALOGO).exists()
    mensaje = (f"{'Actualización' if actualizacion else 'Instalación'} terminada.\n\n"
               f"Carpeta: {raiz}\n\n"
               f"Para abrir el portal: doble clic en \"{NOMBRE_ACCESO}\" en tu Escritorio.")
    if not accesos:
        mensaje += "\n\n(No se pudo crear el acceso directo: usa EJECUTAR_PORTAL.bat dentro de _Programa.)"
    if falta_catalogo:
        mensaje += (f"\n\nFALTA UN PASO: copia el Excel del catálogo de proveedores ({CATALOGO}) en la carpeta "
                    "_Configuracion. Se abrirá la carpeta Honorarios.")
    avisar(mensaje)
    if sys.platform == "win32":
        os.startfile(raiz)                                        # abre la carpeta en el Explorador
    return 0


if __name__ == "__main__":
    opciones = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    opciones.add_argument("--sin-librerias", action="store_true", help="no correr pip (pruebas)")
    opciones.add_argument("--sin-accesos", action="store_true", help="no crear accesos directos (pruebas)")
    argumentos = opciones.parse_args()
    sys.exit(instalar(argumentos.sin_librerias, argumentos.sin_accesos))
