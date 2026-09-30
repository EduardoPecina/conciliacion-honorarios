"""Funciones auxiliares: lectura de importes, normalización de IDs, RFC y nombres, archivos."""

import hashlib
import re
import unicodedata
from decimal import Decimal, InvalidOperation

from .config import PARECE_RFC_MX, PATRON_RFC, RFC_GENERICO_EXTRANJERO


def clasificar_rfc(rfc):
    """'MX' = RFC mexicano válido; 'Extranjero' = ID fiscal de otro país (EIN, CUIT, RUT...: no emite
    CFDI); 'Inválido' = parece RFC mexicano pero está mal escrito; 'Vacío'."""
    if not rfc:
        return "Vacío"
    if rfc == RFC_GENERICO_EXTRANJERO:
        return "Extranjero"
    if PATRON_RFC.fullmatch(rfc):
        return "MX"
    return "Inválido" if PARECE_RFC_MX.match(rfc) else "Extranjero"


class ErrorUsuario(Exception):
    """Error explicado en español para quien usa el programa."""


# ------------------------------------------------------------------
# UTILIDADES
# ------------------------------------------------------------------
def a_decimal(texto):
    try:
        valor = Decimal(str(texto).replace(",", "").strip())
    except (InvalidOperation, ValueError):
        return None
    return valor if valor.is_finite() else None           # descarta "NaN", "Infinity"


def importe_xml(nodo, atributo, obligatorio=True):
    """Lee un importe de un nodo del XML; si es obligatorio y no es número, marca el XML como ilegible."""
    texto = nodo.get(atributo)
    if texto is None and not obligatorio:
        return Decimal("0")
    valor = a_decimal(texto)
    if valor is None:
        raise ValueError(f"el atributo {atributo}='{texto}' no es un importe válido")
    return valor


# Importe del TXT: "1160.00", "1,160.00", "1160" o con signo SAP al final ("1,160.00-").
# Rechaza lo ambiguo ("1.160,00", "1,16", tres decimales): antes se leía mal sin avisar.
PATRON_IMPORTE_TXT = re.compile(r"(-?)(\d{1,3}(?:,\d{3})+|\d+)(\.\d{1,2})?(-?)")


def importe_txt(texto):
    m = PATRON_IMPORTE_TXT.fullmatch(str(texto).strip())
    if not m or (m.group(1) and m.group(4)):
        return None
    valor = Decimal(m.group(2).replace(",", "") + (m.group(3) or ""))
    return -valor if (m.group(1) or m.group(4)) else valor


def normalizar_id(valor):
    texto = "" if valor is None else str(valor).strip()
    if re.fullmatch(r"\d+\.0+", texto):          # 100002.0 -> 100002
        texto = texto.split(".")[0]
    if texto.isdigit():                           # 0000100002 (SAP) -> 100002 (catálogo)
        texto = texto.lstrip("0") or "0"
    return texto


VALORES_ACTIVO = {"1": "1", "SI": "1", "S": "1", "TRUE": "1", "VERDADERO": "1", "ACTIVO": "1", "": "1",
                  "0": "0", "NO": "0", "N": "0", "FALSE": "0", "FALSO": "0", "INACTIVO": "0",
                  "BLOQUEADO": "0", "BAJA": "0"}


def normalizar_activo(valor):
    """Devuelve '1' o '0'. Acepta 1/0, Sí/No, VERDADERO/FALSO (booleanos de Excel), Activo/Inactivo.
    Un valor no reconocido devuelve '?' y el catálogo se rechaza (no se asume activo)."""
    if isinstance(valor, bool):
        return "1" if valor else "0"
    texto = normalizar_nombre(normalizar_id(valor))
    return VALORES_ACTIVO.get(texto, "?")


def normalizar_nombre(texto):
    """Mayúsculas, sin acentos ni signos: para comparar razones sociales."""
    texto = unicodedata.normalize("NFKD", texto or "")
    texto = "".join(c for c in texto if not unicodedata.combining(c)).upper()
    return re.sub(r"[^A-Z0-9Ñ&]", "", texto)


def leer_texto(ruta):
    for codificacion in ("utf-8-sig", "latin-1"):
        try:
            return ruta.read_text(encoding=codificacion)
        except UnicodeDecodeError:
            continue
    raise ErrorUsuario(f"No se pudo leer el archivo {ruta.name}.")


def huella(ruta):
    """SHA-256 del archivo: permite demostrar después con qué entrada exacta se hizo la conciliación."""
    return hashlib.sha256(ruta.read_bytes()).hexdigest()


def guardar_seguro(ruta, escribir):
    """Escribe primero a un temporal y luego lo reemplaza: si algo falla no se corrompe el archivo."""
    tmp = ruta.with_name(ruta.name + ".tmp")
    escribir(tmp)
    tmp.replace(ruta)
