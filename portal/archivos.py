"""Guarda lo que se sube al portal en una carpeta de trabajo con la estructura que espera el motor:

    <carpeta>/Entrada/<archivo>.txt
    <carpeta>/Entrada/XML/...        (XML sueltos y el contenido .xml de los .zip)
    <carpeta>/Salida/

No depende de Streamlit (se puede probar solo). Sin reglas de negocio: solo archivos."""

import zipfile
from io import BytesIO
from pathlib import Path, PurePosixPath

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


def _ruta_segura(nombre):
    """Ruta relativa sin '..', sin unidad ni raíz: un .zip no puede escribir fuera de la carpeta XML."""
    partes = [p for p in PurePosixPath(nombre.replace("\\", "/")).parts if p not in ("", ".", "..", "/")]
    partes = [p for p in partes if not p.endswith(":")]
    return Path(*partes) if partes else None


def _guardar_xml(carpeta_xml, nombre, contenido):
    relativa = _ruta_segura(nombre)
    if relativa is None:
        return 0
    destino = _nombre_libre(carpeta_xml, relativa)
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_bytes(contenido)
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
    return sum(_guardar_xml(carpeta_xml, str(carpeta_zip / m.filename), archivo.read(m)) for m in miembros)


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
