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
import getpass
import os
import platform
import re
import shutil
import socket
import stat
import subprocess
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

ORIGEN = Path(__file__).resolve().parents[1]              # la carpeta del ZIP (o del repositorio)
NOMBRE_CARPETA = "Honorarios"
NOMBRE_ACCESO = "Portal de Honorarios"
CATALOGO = "CatalogoProveedoresTM26.xlsx"
NO_COPIAR = shutil.ignore_patterns(".git", ".github", "__pycache__", "*.pyc", "portal.log", "*.tmp")

BITACORA = "instalacion.log"                      # en la carpeta Honorarios: todo lo que muestra la ventana negra
TAMANO_MAXIMO = 2_000_000                         # bytes; si crece más, se conservan las instalaciones más recientes
PIDE_BITACORA = (f"\n\nSi necesitas ayuda, manda el archivo {BITACORA} (carpeta Honorarios de tu Escritorio) a "
                 "quien mantiene el programa: ahí quedó el detalle de lo que pasó.")

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

SI ALGO FALLA AL INSTALAR
  Manda el archivo instalacion.log (está en esta carpeta) a quien mantiene el programa.
"""


# ------------------------------------------------------------------ Bitácora
class _Doble:
    """Escribe en la ventana negra y en la bitácora a la vez (si una falla, la otra sigue)."""

    def __init__(self, *destinos):
        self.destinos = destinos

    def write(self, texto):
        for destino in self.destinos:
            try:
                destino.write(texto)
                destino.flush()
            except (OSError, ValueError, UnicodeError):
                pass
        return len(texto)

    def flush(self):
        pass


class Bitacora:
    """Copia a Honorarios\\instalacion.log todo lo que el instalador muestra en la ventana negra (incluido lo que
    imprime pip), con un encabezado por instalación. Si un usuario tiene un problema, manda este archivo y se sabe
    exactamente qué pasó, aunque haya cerrado la ventana. Cada instalación se agrega al final (historial)."""

    def __init__(self, ruta):
        self.ruta = ruta

    def __enter__(self):
        self.ruta.parent.mkdir(parents=True, exist_ok=True)
        self._recortar()
        self.archivo = open(self.ruta, "a", encoding="utf-8")
        self._salida, self._errores = sys.stdout, sys.stderr
        sys.stdout = _Doble(self._salida, self.archivo)
        sys.stderr = _Doble(self._errores, self.archivo)
        return self

    def __exit__(self, *_):
        sys.stdout, sys.stderr = self._salida, self._errores
        self.archivo.close()

    def solo_archivo(self, texto):
        self.archivo.write(texto if texto.endswith("\n") else texto + "\n")
        self.archivo.flush()

    def _recortar(self):
        """Si el archivo es muy grande, conserva solo las instalaciones más recientes."""
        try:
            if self.ruta.stat().st_size <= TAMANO_MAXIMO:
                return
            texto = self.ruta.read_text(encoding="utf-8", errors="replace")[-TAMANO_MAXIMO // 2:]
            inicio = texto.find("\n=====")
            self.ruta.write_text(texto[inicio + 1:] if inicio >= 0 else texto, encoding="utf-8")
        except OSError:
            pass


def version_del_programa():
    try:
        texto = (ORIGEN / "honorarios" / "__init__.py").read_text(encoding="utf-8")
        return re.search(r'__version__ = "([^"]+)"', texto).group(1)
    except (OSError, AttributeError):
        return "desconocida"


def encabezado(raiz):
    return "\n".join([
        "=" * 78,
        f"INSTALACIÓN {datetime.now():%d/%m/%Y %H:%M:%S}",
        f"Versión del programa: {version_del_programa()}  (desde {ORIGEN})",
        f"Python: {sys.version.split()[0]}  ({sys.executable})",
        f"Windows: {platform.platform()}",
        f"Usuario: {getpass.getuser()}",
        f"Carpeta: {raiz}",
        "=" * 78,
    ])


def pasos_de_instalar_bat():
    """Lo que INSTALAR.bat anotó antes de llamar a este programa (búsqueda o instalación de Python)."""
    ruta = os.environ.get("HONORARIOS_BITACORA_BAT")
    if not ruta or not Path(ruta).exists():
        return ""
    try:
        return Path(ruta).read_text(encoding="utf-8", errors="replace").strip()
    except OSError:
        return ""


def ejecutar(comando):
    """Corre un comando mostrando su salida en la ventana y en la bitácora. Devuelve el código de salida."""
    entorno = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
    with subprocess.Popen(comando, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                          text=True, encoding="utf-8", errors="replace", env=entorno) as proceso:
        for linea in proceso.stdout:
            print(linea, end="")
        return proceso.wait()


# ------------------------------------------------------------------ Windows
def avisar(mensaje, titulo="Instalador - Portal de Honorarios", error=False):
    if error:
        mensaje += PIDE_BITACORA
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
PYTHON_PROBADO = (3, 13)                                  # el mismo que instala INSTALAR.bat


def instalar_librerias():
    """pip install de requirements-instalar.txt: las versiones EXACTAS de requirements-lock.txt, cada una con la
    huella SHA-256 de sus instaladores (--require-hashes: si un archivo descargado no coincide byte por byte, se
    detiene). Si ya están, solo lo confirma. Solo instaladores precompilados: nunca intenta compilar. Reintenta por
    la red corporativa."""
    print("Revisando librerías (pandas, openpyxl, streamlit, pywin32) en las versiones probadas...")
    comando = [sys.executable, "-m", "pip", "install", "-r", str(ORIGEN / "requirements-instalar.txt"),
               "--require-hashes", "--only-binary=:all:",
               "--disable-pip-version-check", "--progress-bar", "off", "--retries", "10", "--timeout", "60"]
    for intento in (1, 2):
        if ejecutar(comando) == 0:
            return True
        print(f"La instalación de librerías falló (intento {intento} de 2). Reintentando...")
    return False


def borrar_carpeta(ruta, intentos=10, espera=1.0):
    """Borra una carpeta aunque tenga archivos de solo lectura. Reintenta unos segundos: OneDrive (o el antivirus)
    a veces deja archivos abiertos un momento después de copiarlos. True si quedó borrada."""
    def quitar_solo_lectura(funcion, camino, _):
        try:
            os.chmod(camino, stat.S_IWRITE)
            funcion(camino)
        except OSError:
            pass
    opcion = {"onexc": quitar_solo_lectura} if sys.version_info >= (3, 12) else {"onerror": quitar_solo_lectura}
    for intento in range(intentos):
        if not ruta.exists():
            return True
        shutil.rmtree(ruta, **opcion)
        if ruta.exists() and intento < intentos - 1:
            time.sleep(espera)
    return not ruta.exists()


def respaldos(destino):
    """Copias de seguridad que deja una actualización: _Programa_anterior, _Programa_anterior_2..."""
    return sorted(p for p in destino.parent.glob(destino.name + "_anterior*") if p.is_dir())


def copiar_programa(destino):
    """Copia el programa a Honorarios\\_Programa. Al actualizar, la versión anterior se reemplaza completa
    (así no quedan archivos viejos). Si el origen ya es el destino (se corrió desde _Programa), no copia.
    Mientras copia, la versión anterior queda como respaldo (_Programa_anterior); al terminar se borra. Si OneDrive
    no deja borrarla en ese momento, la borra el portal la próxima vez que se abra."""
    if destino.exists() and destino.resolve() == ORIGEN:
        print("El programa ya está en su lugar (se corrió desde _Programa).")
        return
    for viejo in respaldos(destino):                     # restos de actualizaciones anteriores
        borrar_carpeta(viejo, intentos=3)
    anterior, n = destino.with_name(destino.name + "_anterior"), 2
    while anterior.exists():                             # uno viejo sigue bloqueado: se usa otro nombre
        anterior, n = destino.with_name(f"{destino.name}_anterior_{n}"), n + 1
    if destino.exists():
        destino.rename(anterior)             # si el portal o un archivo está abierto, Windows no deja: PermissionError
    try:
        shutil.copytree(ORIGEN, destino, ignore=NO_COPIAR)
    except Exception:
        if anterior.exists() and not destino.exists():
            anterior.rename(destino)         # algo falló: se deja la versión anterior como estaba
        raise
    if not borrar_carpeta(anterior):
        print(f"OneDrive todavía usa {anterior.name}: se borrará sola la próxima vez que abras el portal.")


# Archivos que dejaron versiones anteriores del instalador y ya no existen (se quitan al actualizar)
OBSOLETOS = ["Desinstalar Portal de Honorarios.lnk"]           # el desinstalador se retiró en la versión 1.10.1


def crear_estructura(raiz):
    for carpeta in (raiz / "_Configuracion", raiz / "Reportes"):
        carpeta.mkdir(parents=True, exist_ok=True)
    (raiz / "LEEME.txt").write_text(LEEME, encoding="utf-8-sig")
    for nombre in OBSOLETOS:
        (raiz / nombre).unlink(missing_ok=True)


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
    return ejecutar([sys.executable, "-c", codigo, str(destino), argumentos, str(programa),
                     str(icono or ""), *rutas]) == 0


def instalar(sin_librerias=False, sin_accesos=False):
    """Instala o actualiza, dejando todo registrado en Honorarios\\instalacion.log."""
    raiz = escritorio() / NOMBRE_CARPETA
    try:
        bitacora = Bitacora(raiz / BITACORA).__enter__()
    except OSError:                                   # sin bitácora se instala igual
        return _instalar(raiz, sin_librerias, sin_accesos)
    try:
        bitacora.solo_archivo("\n" + encabezado(raiz))
        bat = pasos_de_instalar_bat()
        if bat:
            bitacora.solo_archivo("-- INSTALAR.bat --\n" + bat + "\n-- Instalador --")
        try:
            resultado = _instalar(raiz, sin_librerias, sin_accesos)
        except Exception:                             # noqa: BLE001 - cualquier error inesperado queda registrado
            bitacora.solo_archivo(traceback.format_exc())
            avisar("La instalación se detuvo por un error inesperado.", error=True)
            resultado = 1
        bitacora.solo_archivo(f"RESULTADO: {'terminó bien' if resultado == 0 else 'NO se completó'} "
                              f"({datetime.now():%d/%m/%Y %H:%M:%S})")
        return resultado
    finally:
        bitacora.__exit__()


def _instalar(raiz, sin_librerias, sin_accesos):
    programa = raiz / "_Programa"
    print(f"Instalando en: {raiz}")

    if sys.version_info[:2] != PYTHON_PROBADO and not sin_librerias:
        version = ".".join(map(str, PYTHON_PROBADO))
        avisar(f"Este instalador debe correr con Python {version} (se está usando "
               f"{sys.version_info.major}.{sys.version_info.minor}).\n\nUsa INSTALAR.bat: busca Python {version} y, "
               "si no está, lo instala solo para tu usuario.", error=True)
        return 1
    if portal_abierto():
        avisar("El Portal de Honorarios está abierto.\n\nCiérralo (botón \"Cerrar el portal\") y vuelve a "
               "correr INSTALAR.bat.", error=True)
        return 1
    if not sin_librerias and not instalar_librerias():
        avisar("No se pudieron instalar las librerías de Python (¿sin internet o la red las bloquea?).\n\n"
               "Vuelve a intentarlo en unos minutos.", error=True)
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
