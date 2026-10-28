"""Guarda lo que se sube al portal en una carpeta de trabajo con la estructura que espera el motor:

    <carpeta>/Entrada/<archivo>.txt
    <carpeta>/Entrada/XML/...        (XML sueltos y el contenido .xml de los .zip)
    <carpeta>/Salida/

No depende de Streamlit (se puede probar solo). Sin reglas de negocio: solo archivos."""

import re
import zipfile
from io import BytesIO
from pathlib import Path

from honorarios.entradas import datos_basicos_cfdi
from honorarios.utilidades import ErrorUsuario

MAX_DESCOMPRIMIDO = 500 * 1024 * 1024        # 500 MB por .zip: evita "zip bombs"


def _nombre_libre(carpeta, relativa):
    """Evita que dos archivos con el mismo nombre se pisen: 'a.xml' -> 'a (2).xml'."""
    destino = carpeta / relativa
    k = 2
    while destino.exists():
        destino = carpeta / relativa.parent / f"{relativa.stem} ({k}){relativa.suffix}"
        k += 1
    return destino


PROHIBIDOS = re.compile(r'[<>:"|?*\x00-\x1f]')          # Windows no los acepta en un nombre (':' = unidad o flujo)
RESERVADOS = ({"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"}   # dispositivos de Windows, no archivos
              | {f"{d}{n}" for d in ("COM", "LPT") for n in "123456789¹²³"})


def _parte_segura(parte):
    """Una parte de la ruta tal como Windows la guardaría sin sorpresas, o None si no es un nombre ('', '.', '..',
    '...', '.. ': Windows recorta los puntos y espacios del final)."""
    parte = PROHIBIDOS.sub("_", parte).rstrip(" .")
    if not parte.strip(" ."):
        return None
    if parte.split(".")[0].rstrip(" ").upper() in RESERVADOS:     # 'CON.xml', 'COM1.xml'...
        parte = "_" + parte
    return parte


def _partes_seguras(nombre):
    """Partes de la ruta sin '..', sin unidad ni raíz y sin caracteres que Windows no acepta. Se separa a mano (no con
    PurePosixPath): '//servidor/x.xml' sería en Windows una ruta de red y 'a/D:b/x.xml' otra unidad. Si el nombre
    del archivo (la última parte) no sirve, no se guarda: lista vacía (no se usa el nombre de su carpeta)."""
    partes = [_parte_segura(p) for p in nombre.replace("\\", "/").split("/")]
    if partes[-1] is None:
        return []
    return [p for p in partes if p]


def _ruta_segura(nombre):
    """Ruta relativa sin '..', sin unidad ni raíz: un .zip no puede escribir fuera de la carpeta XML."""
    partes = _partes_seguras(nombre)
    return Path(*partes) if partes else None


def _guardar_xml(carpeta_xml, nombre, contenido):
    relativa = _ruta_segura(nombre)
    if relativa is None:
        return 0
    destino = _nombre_libre(carpeta_xml, relativa)
    if not destino.resolve().is_relative_to(Path(carpeta_xml).resolve()):     # segunda revisión, por si acaso
        return 0
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        destino.write_bytes(contenido)
    except OSError as e:
        raise ErrorUsuario(f"No se pudo guardar '{nombre}' ({e.strerror or e}). Cámbiale el nombre y vuelve a "
                           "subirlo.")
    return 1


def _extraer_zip(carpeta_xml, nombre_zip, contenido):
    try:
        archivo = zipfile.ZipFile(BytesIO(contenido))
    except zipfile.BadZipFile:
        raise ErrorUsuario(f"'{nombre_zip}' no es un .zip válido.")
    miembros = [m for m in archivo.infolist() if not m.is_dir() and m.filename.lower().endswith(".xml")]
    if sum(m.file_size for m in miembros) > MAX_DESCOMPRIMIDO:
        raise ErrorUsuario(f"'{nombre_zip}' es demasiado grande al descomprimirlo (más de 500 MB).")
    carpeta_zip = Path(_ruta_segura(Path(nombre_zip).stem) or "zip")
    # Se une como texto: en Windows, Path("Lote") / "//servidor/x.xml" o / "D:x.xml" descartaría "Lote"
    return sum(_guardar_xml(carpeta_xml, f"{carpeta_zip.as_posix()}/{m.filename}", archivo.read(m))
               for m in miembros)


def preparar_carpeta(carpeta, txt, facturas):
    """txt: (nombre, bytes). facturas: lista de (nombre, bytes) de archivos .xml o .zip.
    Devuelve cuántos XML quedaron en la carpeta."""
    carpeta = Path(carpeta)
    entrada, carpeta_xml = carpeta / "Entrada", carpeta / "Entrada" / "XML"
    carpeta_xml.mkdir(parents=True, exist_ok=True)
    (carpeta / "Salida").mkdir(exist_ok=True)

    nombre_txt, contenido_txt = txt
    (entrada / (Path(nombre_txt).stem + ".txt")).write_bytes(contenido_txt)
    return preparar_xmls(carpeta_xml, facturas)


def preparar_xmls(carpeta_xml, facturas):
    """Solo los XML (sueltos o en .zip) en carpeta_xml: para las facturas pendientes, que no traen TXT.
    Devuelve cuántos XML quedaron."""
    carpeta_xml = Path(carpeta_xml)
    carpeta_xml.mkdir(parents=True, exist_ok=True)
    total = 0
    for nombre, contenido in facturas:
        if nombre.lower().endswith(".zip"):
            total += _extraer_zip(carpeta_xml, nombre, contenido)
        elif nombre.lower().endswith(".xml"):
            total += _guardar_xml(carpeta_xml, Path(nombre).name, contenido)
    if not total:
        raise ErrorUsuario("No se encontró ningún archivo .xml en lo que se subió.")
    return total


# ------------------------------------------------------------------ Administrar carpeta de XML
PATRON_UUID = re.compile(r"[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}", re.IGNORECASE)


def _clave_ruta(texto):
    """Ruta para comparar: con /, sin comillas ni espacios de más, sin distinguir mayúsculas."""
    return texto.strip().strip('"').strip("'").strip().replace("\\", "/").strip("/").lower()


def leer_zip(nombre_zip, contenido):
    """-> (zipfile abierto, [{"Ruta": ruta dentro del ZIP, "ArchivoXML": como la escribe el Excel de la conciliación,
    "UUID", "RFCEmisor", "NombreEmisor", "FechaXML", "Total"}] de cada XML). ErrorUsuario si no es un ZIP que se
    pueda leer o es demasiado grande."""
    try:
        archivo = zipfile.ZipFile(BytesIO(contenido))
        miembros = [m for m in archivo.infolist() if not m.is_dir()]
    except zipfile.BadZipFile:
        raise ErrorUsuario(f"'{nombre_zip}' no es un .zip válido.")
    if sum(m.file_size for m in miembros) > MAX_DESCOMPRIMIDO:
        raise ErrorUsuario(f"'{nombre_zip}' es demasiado grande al descomprimirlo (más de 500 MB).")
    prefijo = (_ruta_segura(Path(nombre_zip).stem) or Path("zip")).as_posix()
    xmls = []
    for m in miembros:
        if not m.filename.lower().endswith(".xml"):
            continue
        try:
            datos = datos_basicos_cfdi(archivo.read(m))
        except (RuntimeError, zipfile.BadZipFile, NotImplementedError) as e:   # protegido con contraseña, dañado
            raise ErrorUsuario(f"No se pudo leer '{m.filename}' dentro de '{nombre_zip}' ({e}).")
        relativa = _ruta_segura(m.filename)
        xmls.append({"Ruta": m.filename,
                     "ArchivoXML": f"{prefijo}/{relativa.as_posix()}" if relativa else m.filename, **datos})
    return archivo, xmls


def buscar_en_zip(xmls, pedidos):
    """pedidos: renglones que pegó el usuario (ruta de la columna ArchivoXML del Excel, ruta dentro del ZIP, solo el
    nombre del archivo, o el UUID). -> (XML encontrados, [(renglón, por qué no se encontró)]).
    Un nombre de archivo que se repite en varias carpetas no se adivina: se pide la ruta completa."""
    encontrados, problemas = [], []
    for renglon in pedidos:
        texto = renglon.strip().strip('"').strip("'").strip()
        if not texto or texto.lower() in ("archivoxml", "uuid", "ruta"):     # vacío o el encabezado de la columna
            continue
        uuid = PATRON_UUID.fullmatch(texto)
        if uuid:
            hallados = [x for x in xmls if x["UUID"] == texto.upper()]
        else:
            clave = _clave_ruta(texto)
            hallados = [x for x in xmls if clave in (_clave_ruta(x["ArchivoXML"]), _clave_ruta(x["Ruta"]))]
            if not hallados:
                por_nombre = [x for x in xmls if _clave_ruta(x["Ruta"]).rsplit("/", 1)[-1] == clave.rsplit("/", 1)[-1]]
                if len(por_nombre) == 1:
                    hallados = por_nombre
                elif len(por_nombre) > 1:
                    problemas.append((texto, f"Hay {len(por_nombre)} archivos con ese nombre en distintas carpetas: "
                                             "pega la ruta completa (columna ArchivoXML del Excel) o el UUID"))
                    continue
        if not hallados:
            problemas.append((texto, "No está en el ZIP" + (" (ningún XML tiene ese UUID)" if uuid else "")))
        for x in hallados:
            if x not in encontrados:
                encontrados.append(x)
    return encontrados, problemas


def zip_sin(archivo, quitar, con_mismo_nombre=True):
    """ZIP nuevo igual al original (mismas carpetas, nombres, fechas y contenido) pero sin los XML de quitar.
    con_mismo_nombre: quita también los archivos de la misma carpeta con el mismo nombre y otra extensión (el PDF de
    esa factura). -> (bytes del ZIP nuevo, rutas que se quitaron)."""
    rutas = {x["Ruta"] for x in quitar}
    if con_mismo_nombre:
        bases = {r.rsplit(".", 1)[0].lower() for r in rutas}
        rutas |= {m.filename for m in archivo.infolist()
                  if not m.is_dir() and m.filename.rsplit(".", 1)[0].lower() in bases}
    salida = BytesIO()
    with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as nuevo:
        for m in archivo.infolist():
            if m.filename not in rutas:
                nuevo.writestr(m, archivo.read(m))
    return salida.getvalue(), sorted(rutas)
