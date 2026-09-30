"""
Conciliación de Honorarios: cruce interno TXT -> Catálogo -> XML con reglas de negocio
(opcional: códigos de retención en SAP). No consulta servicios externos.

Estructura esperada (relativa a este archivo, funciona en cualquier computadora):
    Honorarios/
    ├── _Programa/conciliar.py                 <- este archivo
    ├── _Configuracion/
    │   ├── CatalogoProveedoresTM26.xlsx       (tabla "Catalogo")
    │   └── HistorialUUID.csv                  (se crea solo: UUID ya conciliados en meses anteriores)
    └── <Año>/<MM Mes>/
        ├── Entrada/  (un solo .txt suelto + carpeta XML/)
        └── Salida/   (aquí se guarda el resultado)

Regla de oro: el Estatus "Cruzado" solo se asigna si la línea pasó TODAS las validaciones
de la lista VALIDACIONES. Si alguna no se pudo ejecutar (historial ilegible, SAP pedido pero
sin respuesta...), la línea sale "Cruzado con observaciones" y dice cuál faltó.
"""

import hashlib
import re
import unicodedata
import xml.etree.ElementTree as ET
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from bisect import bisect_left
from itertools import combinations
from math import comb
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------
# CONFIGURACIÓN (lo único que se podría ajustar)
# ------------------------------------------------------------------
RFC_SOCIEDAD = "SIN070530EI5"            # RFC receptor obligatorio en los XML
SOCIEDAD_SAP = "TM26"

# Datos del receptor en CFDI 4.0. Si se dejan vacíos, esa validación no se hace (se avisa al iniciar).
NOMBRE_SOCIEDAD = "SERVICIOS INTEGRALES NOVA DE MONTERREY"   # Razón social de la Constancia, sin "S.A. de C.V."
CP_SOCIEDAD = ""                         # Código postal del domicilio fiscal, ej. "64000"
REGIMEN_RECEPTOR = "601"                 # General de Ley Personas Morales
USOS_CFDI_OK = {"G03"}                   # Gastos en general
VERSIONES_CFDI_OK = {"4.0"}

ARCHIVO_CATALOGO = "CatalogoProveedoresTM26.xlsx"
NOMBRE_TABLA_CATALOGO = "Catalogo"
TOLERANCIA = Decimal("1.00")             # diferencia máxima de importe aceptada (cruce TXT vs XML)
TOLERANCIA_IMPUESTOS = Decimal("1.00")   # variación de impuestos aceptada como redondeo
UMBRAL_OK = Decimal("0.05")              # hasta aquí la variación se etiqueta OK; arriba, "Redondeo"

# True: una diferencia de importe entre 0.05 y 1.00 deja la línea "Cruzado con observaciones".
DIFERENCIA_IMPORTE_ES_OBSERVACION = True
# UUID repetido en la carpeta XML (la misma factura copiada dos veces):
#   False -> si solo UNA línea del TXT la paga, queda "Cruzado con observaciones" (no hay pago doble);
#            se detiene solo si dos líneas distintas terminan pagando la misma factura.
#   True  -> siempre "Detenida".
UUID_REPETIDO_SIEMPRE_DETIENE = False
MSG_UUID_REPETIDO = "Verificar con Responsables de Envío de XML"
MSG_SAP_SIN_IH = ("Discrepancia en Sistema e Información Enviada por Customer Services NOVA, "
                  "revisar Constancia Fiscal y Alta en el Sistema")
# True: si la fecha de la línea del TXT no es del periodo, se anota como observación.
FECHA_TXT_DEBE_SER_DEL_PERIODO = True

# Tasas de impuestos (las usa el programa para decidir el Estatus y el Excel para mostrar el cálculo)
TASA_IVA = Decimal("0.16")
TASA_ISR_GENERAL = Decimal("0.10")       # honorarios, régimen 612
TASA_ISR_RESICO = Decimal("0.0125")      # régimen 626

# Claves SAT de producto/servicio que se consideran servicios profesionales (primeros 2 dígitos)
PREFIJOS_CLAVE_SERVICIOS = ("80", "81", "82", "83", "84", "85", "86", "93")

ARCHIVO_HISTORIAL = "HistorialUUID.csv"

MESES = ["01 Ene", "02 Feb", "03 Mar", "04 Abr", "05 May", "06 Jun",
         "07 Jul", "08 Ago", "09 Sep", "10 Oct", "11 Nov", "12 Dic"]

COLUMNAS_TXT = ["Fecha", "IDProveedor", "Folio", "Importe", "FijoX",
                "CodigoImpuesto", "Fijo1", "CentroCosto", "Cuenta", "PM", "Referencia"]
CODIGOS_IMPUESTO_TXT = {"E0", "F2"}      # E0 sin IVA, F2 con IVA

METODOS_PAGO_OK = {"PUE", "PPD"}         # métodos de pago aceptados
FORMAS_PAGO_OK = {"03", "99"}            # transferencia electrónica / por definir

FORMAS_PAGO = {"01": "Efectivo", "02": "Cheque nominativo", "03": "Transferencia electrónica",
               "04": "Tarjeta de crédito", "28": "Tarjeta de débito", "99": "Por definir"}

REGIMENES = {
    "601": "General de Ley Personas Morales", "603": "Personas Morales con Fines no Lucrativos",
    "605": "Sueldos y Salarios", "606": "Arrendamiento", "607": "Enajenación o Adquisición de Bienes",
    "608": "Demás ingresos", "610": "Residentes en el Extranjero",
    "612": "Personas Físicas con Actividades Empresariales y Profesionales",
    "614": "Ingresos por intereses", "616": "Sin obligaciones fiscales",
    "621": "Incorporación Fiscal", "625": "Plataformas Tecnológicas",
    "626": "Régimen Simplificado de Confianza (RESICO)",
}

# Códigos de retención en SAP (tabla LFBW)
ISR_SAP = {"RS": "RESICO", "RH": "No RESICO", "RM": "No RESICO", "RE": "Extranjero"}

# Validaciones que una línea DEBE pasar para quedar "Cruzado" (clave -> nombre para el usuario).
# "sap" solo se exige si en esa corrida se eligió validar contra SAP.
VALIDACIONES = {
    "cfdi": "reglas del CFDI",
    "impuestos": "recálculo de impuestos",
    "sap": "códigos de retención en SAP",
    "historial": "historial de UUID conciliados",
}

# Observaciones que no quitan el "Cruzado" (son informativas, no hay nada que revisar)
INFORMATIVAS = ("Pago quebrado", "Factura PPD")
# Observaciones que mencionan SAP pero no son discrepancias con SAP
NO_ES_DISCREPANCIA_SAP = ("Lleva IVA: falta validar", "Validaciones no realizadas", "No se pudo validar en SAP",
                          "ID fiscal extranjero")

RAIZ = Path(__file__).resolve().parent.parent
CONFIG = RAIZ / "_Configuracion"
PATRON_RFC = re.compile(r"[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}")
PARECE_RFC_MX = re.compile(r"[A-ZÑ&]{3,4}\d")          # empieza como RFC mexicano (3-4 letras y dígitos)
RFC_GENERICO_EXTRANJERO = "XEXX010101000"


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


# ------------------------------------------------------------------
# 1. RUTAS
# ------------------------------------------------------------------
def pedir_periodo():
    hoy = datetime.now()
    print("=" * 60)
    print("  CONCILIACIÓN DE HONORARIOS")
    print("=" * 60)
    anio = input(f"Año [{hoy.year}]: ").strip() or str(hoy.year)
    mes = input(f"Mes (1-12) [{hoy.month}]: ").strip() or str(hoy.month)
    if not anio.isdigit() or not mes.isdigit() or not 1 <= int(mes) <= 12:
        raise ErrorUsuario("El año o el mes no son válidos. Ejemplo: año 2026, mes 9.")
    return int(anio), int(mes)


def obtener_rutas(anio, mes):
    base = RAIZ / str(anio) / MESES[mes - 1]
    entrada, salida = base / "Entrada", base / "Salida"
    carpeta_xml = entrada / "XML"
    catalogo = CONFIG / ARCHIVO_CATALOGO

    if not base.exists():
        raise ErrorUsuario(f"No existe la carpeta del periodo:\n  {base}")
    if not catalogo.exists():
        raise ErrorUsuario(f"No se encontró el catálogo. Debe estar en:\n  {catalogo}")
    if not carpeta_xml.exists():
        raise ErrorUsuario(f"No existe la carpeta de XML:\n  {carpeta_xml}")

    txts = sorted(entrada.glob("*.txt"))
    if not txts:
        raise ErrorUsuario(f"No hay ningún archivo .txt en:\n  {entrada}")
    if len(txts) > 1:
        nombres = "\n  ".join(t.name for t in txts)
        raise ErrorUsuario(f"Hay más de un .txt en la carpeta Entrada. Deja solo uno:\n  {nombres}")

    xmls = sorted(p for p in carpeta_xml.rglob("*") if p.suffix.lower() == ".xml")
    if not xmls:
        raise ErrorUsuario(f"La carpeta XML está vacía:\n  {carpeta_xml}")

    salida.mkdir(parents=True, exist_ok=True)
    return {"catalogo": catalogo, "txt": txts[0], "xmls": xmls, "carpeta_xml": carpeta_xml,
            "salida": salida}


# ------------------------------------------------------------------
# 2. LECTURA DE ENTRADAS
# ------------------------------------------------------------------
def leer_catalogo(ruta):
    try:
        libro = load_workbook(ruta, read_only=False, data_only=True)
    except PermissionError:
        raise ErrorUsuario(f"El catálogo está abierto o bloqueado. Ciérralo y vuelve a intentar:\n  {ruta}")

    for hoja in libro.worksheets:
        if NOMBRE_TABLA_CATALOGO in hoja.tables:
            rango = hoja.tables[NOMBRE_TABLA_CATALOGO].ref
            filas = [[c.value for c in fila] for fila in hoja[rango]]
            break
    else:
        raise ErrorUsuario(f"No se encontró la tabla '{NOMBRE_TABLA_CATALOGO}' en {ruta.name}.")

    encabezados = [str(h).strip() for h in filas[0]]
    faltan = {"IDProveedor", "Nombre", "RFC"} - set(encabezados)
    if faltan:
        raise ErrorUsuario(f"Al catálogo le faltan columnas: {', '.join(sorted(faltan))}")

    catalogo = pd.DataFrame(filas[1:], columns=encabezados)
    catalogo["IDProveedor"] = catalogo["IDProveedor"].map(normalizar_id)
    catalogo = catalogo[catalogo["IDProveedor"] != ""]                 # renglones vacíos de la tabla
    catalogo["RFC"] = (catalogo["RFC"].fillna("").astype(str).str.upper()
                       .str.replace(r"[\s\-\.]", "", regex=True))          # "PEGJ-800101 AB1" -> "PEGJ800101AB1"
    catalogo["RFC"] = catalogo["RFC"].map(lambda r: normalizar_id(r) if re.fullmatch(r"\d+\.0+", r) else r)
    catalogo["TipoRFC"] = catalogo["RFC"].map(clasificar_rfc)
    # Si el catálogo trae la columna RFCTipo, manda sobre la regla por formato:
    #   EXTRANJERO -> extranjero aunque el ID fiscal traiga letras (DE318966062, B84720473...)
    #   otro valor -> se espera RFC mexicano; si no tiene el formato, es un RFC mal escrito
    col_tipo = next((c for c in ("RFCTipo", "TipoRFC_Catalogo", "Tipo RFC") if c in catalogo.columns), None)
    if col_tipo:
        declarado = catalogo[col_tipo].fillna("").astype(str).map(normalizar_nombre)
        catalogo["TipoRFC"] = [
            ("Extranjero" if d == "EXTRANJERO" and t != "Vacío"
             else "Inválido" if d and t == "Extranjero"
             else t)
            for d, t in zip(declarado, catalogo["TipoRFC"])]
    catalogo["Nombre"] = catalogo["Nombre"].fillna("").astype(str).str.strip()
    if "Activo" in catalogo.columns:
        originales = catalogo["Activo"]
        catalogo["Activo"] = originales.map(normalizar_activo)
        raros = [f"{i} ('{v}')" for i, v, a in zip(catalogo["IDProveedor"], originales, catalogo["Activo"])
                 if a == "?"]
        if raros:
            raise ErrorUsuario("Valores no reconocidos en la columna Activo (usa 1/0 o Sí/No):\n  "
                               + "\n  ".join(raros[:10]))
    else:
        catalogo["Activo"] = "1"

    duplicados = catalogo[catalogo["IDProveedor"].duplicated()]["IDProveedor"].tolist()
    if duplicados:
        raise ErrorUsuario(f"El catálogo tiene IDs repetidos: {', '.join(duplicados[:10])}")
    # Un RFC raro ya no detiene todo el programa: solo afecta a las líneas de ese proveedor.
    malos = [f"{i} ({r})" for i, r, t in zip(catalogo["IDProveedor"], catalogo["RFC"], catalogo["TipoRFC"])
             if t == "Inválido"]
    if malos:
        print(f"  Aviso: {len(malos)} proveedor(es) con RFC mal escrito; sus líneas saldrán "
              f"'RFC inválido en catálogo':\n    " + "\n    ".join(malos[:10])
              + ("\n    ..." if len(malos) > 10 else ""))
    return catalogo.reset_index(drop=True)


def leer_txt(ruta):
    lineas = leer_texto(ruta).splitlines()
    filas, errores, sin_id = [], [], []
    for n, linea in enumerate(lineas, start=1):
        if not linea.strip():
            continue
        campos = [c.strip() for c in linea.split("\t")]
        if n == 1 and not re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", campos[0]):
            continue                                   # es un encabezado, se omite
        if len(campos) != len(COLUMNAS_TXT):
            errores.append(f"Línea {n}: trae {len(campos)} columnas y se esperaban {len(COLUMNAS_TXT)}")
            continue
        fila = dict(zip(COLUMNAS_TXT, campos))
        try:
            fila["_fecha"] = datetime.strptime(fila["Fecha"], "%d.%m.%Y")
        except ValueError:
            errores.append(f"Línea {n}: la fecha '{fila['Fecha']}' no es válida (formato DD.MM.AAAA)")
            continue
        fila["Linea"] = n
        fila["IDProveedor"] = normalizar_id(fila["IDProveedor"])
        if not fila["IDProveedor"]:
            sin_id.append(n)                           # no detiene: sale en el reporte con su importe
        fila["CodigoImpuesto"] = fila["CodigoImpuesto"].upper()
        fila["Importe"] = importe_txt(fila["Importe"])
        if fila["Importe"] is None:
            errores.append(f"Línea {n}: el importe '{campos[3]}' no es válido "
                           "(formato esperado 1234.56 o 1,234.56, máximo 2 decimales)")
            continue
        filas.append(fila)

    if errores:
        raise ErrorUsuario("El TXT tiene líneas con formato incorrecto:\n  " + "\n  ".join(errores[:15]))
    if not filas:
        raise ErrorUsuario(f"El TXT no tiene líneas para procesar: {ruta.name}")
    if sin_id:
        print(f"  Aviso: {len(sin_id)} línea(s) sin IDProveedor ({', '.join(map(str, sin_id[:15]))}"
              f"{'...' if len(sin_id) > 15 else ''}): salen en el reporte como 'Línea sin IDProveedor'")
    return filas


def leer_cfdi(ruta, carpeta):
    raiz = ET.parse(ruta).getroot()
    if not raiz.tag.endswith("Comprobante"):
        raise ValueError("no es un CFDI (Comprobante)")
    ns_cfdi = raiz.tag.split("}")[0].strip("{")          # soporta CFDI 3.3 y 4.0
    ns = {"cfdi": ns_cfdi, "tfd": "http://www.sat.gob.mx/TimbreFiscalDigital"}

    emisor = raiz.find("cfdi:Emisor", ns)
    receptor = raiz.find("cfdi:Receptor", ns)
    timbre = raiz.find("cfdi:Complemento/tfd:TimbreFiscalDigital", ns)
    impuestos = raiz.find("cfdi:Impuestos", ns)
    emisor = emisor if emisor is not None else {}
    receptor = receptor if receptor is not None else {}

    ret_isr = ret_iva = iva = otros_tras = otras_ret = Decimal("0")
    iva_exento, tasas_iva = False, set()
    if impuestos is not None:
        for r in impuestos.findall("cfdi:Retenciones/cfdi:Retencion", ns):
            imp = importe_xml(r, "Importe")
            if r.get("Impuesto") == "001":
                ret_isr += imp
            elif r.get("Impuesto") == "002":
                ret_iva += imp
            else:
                otras_ret += imp
        for t in impuestos.findall("cfdi:Traslados/cfdi:Traslado", ns):
            if t.get("TipoFactor") == "Exento":
                iva_exento = iva_exento or t.get("Impuesto") == "002"
                continue
            imp = importe_xml(t, "Importe")
            if t.get("Impuesto") == "002":
                iva += imp
                tasas_iva.add(t.get("TasaOCuota", ""))
            else:
                otros_tras += imp

    # Impuestos locales (complemento implocal): afectan el Total
    loc_ret = loc_tras = Decimal("0")
    for el in raiz.iter():
        if el.tag.endswith("}ImpuestosLocales"):
            loc_ret += importe_xml(el, "TotaldeRetenciones", obligatorio=False)
            loc_tras += importe_xml(el, "TotaldeTraslados", obligatorio=False)

    conceptos = raiz.findall("cfdi:Conceptos/cfdi:Concepto", ns)
    subtotal = importe_xml(raiz, "SubTotal")
    descuento = importe_xml(raiz, "Descuento", obligatorio=False)
    total = importe_xml(raiz, "Total")
    total_calc = subtotal - descuento + iva + otros_tras - ret_isr - ret_iva - otras_ret + loc_tras - loc_ret
    return {
        "Descuento": descuento,
        "ArchivoXML": str(ruta.relative_to(carpeta)),     # ruta relativa: evita choques de nombres en subcarpetas
        "UUID": (timbre.get("UUID", "") if timbre is not None else "").upper().strip(),
        "Version": raiz.get("Version", ""),
        "FechaXML": (raiz.get("Fecha") or "")[:10],
        "MetodoPago": raiz.get("MetodoPago", ""),
        "FormaPago": raiz.get("FormaPago", ""),
        "TipoComprobante": raiz.get("TipoDeComprobante", ""),
        "Moneda": raiz.get("Moneda", ""),
        "RFCEmisor": emisor.get("Rfc", "").upper().strip(),
        "NombreEmisor": emisor.get("Nombre", ""),
        "RegimenEmisor": emisor.get("RegimenFiscal", ""),
        "RFCReceptor": receptor.get("Rfc", "").upper().strip(),
        "NombreReceptor": receptor.get("Nombre", ""),
        "RegimenReceptor": receptor.get("RegimenFiscalReceptor", ""),
        "CPReceptor": receptor.get("DomicilioFiscalReceptor", ""),
        "UsoCFDI": receptor.get("UsoCFDI", ""),
        "SubTotal": subtotal,
        "IVA": iva,
        "IVAExento": "Sí" if iva_exento else "No",
        "TasasIVA": tasas_iva,
        "OtrosImpuestos": otros_tras + otras_ret + loc_tras + loc_ret,
        "RetISR": ret_isr,
        "RetIVA": ret_iva,
        "Total": total,
        "TotalTexto": raiz.get("Total", ""),
        "DifTotalXML": total - total_calc,
        "ImporteConIVA": subtotal - descuento + iva,
        "SumaConceptos": sum((importe_xml(c, "Importe") for c in conceptos), Decimal("0")),
        "ClavesProdServ": [c.get("ClaveProdServ", "") for c in conceptos],
        "ObjetosImp": [c.get("ObjetoImp", "") for c in conceptos],
        # Conceptos 02 (sí objeto de impuesto) que no traen NINGÚN impuesto (ni traslado ni retención)
        "Conceptos02SinImpuestos": sum(
            1 for c in conceptos if c.get("ObjetoImp") == "02"
            and c.find("cfdi:Impuestos/cfdi:Traslados/cfdi:Traslado", ns) is None
            and c.find("cfdi:Impuestos/cfdi:Retenciones/cfdi:Retencion", ns) is None),
    }


def leer_xmls(rutas, carpeta):
    facturas, errores = [], []
    for ruta in rutas:
        try:
            facturas.append(leer_cfdi(ruta, carpeta))
        except Exception as e:
            errores.append({"Archivo": str(ruta.relative_to(carpeta)), "Error": f"No se pudo leer: {e}"})
    return facturas, errores


# ------------------------------------------------------------------
# 2b. HISTORIAL DE UUID (archivo interno: evita pagar dos veces la misma factura)
# ------------------------------------------------------------------
def cargar_historial():
    ruta = CONFIG / ARCHIVO_HISTORIAL
    if not ruta.exists():
        return {}, "Nuevo (se crea al terminar)"
    try:
        df = pd.read_csv(ruta, dtype=str, encoding="utf-8-sig").fillna("")
    except Exception as e:
        return None, f"No se pudo leer {ARCHIVO_HISTORIAL}: {e}"
    faltan = {"UUID", "Periodo"} - set(df.columns)
    if faltan:
        return None, f"A {ARCHIVO_HISTORIAL} le faltan columnas: {', '.join(sorted(faltan))}"
    historial = {}
    for r in df.to_dict("records"):
        historial.setdefault(r["UUID"].upper(), []).append(r)
    return historial, f"{len(historial):,} UUID registrados"


def guardar_historial(resultado, anio, mes):
    """Registra los UUID cruzados en este periodo (reemplaza lo que hubiera del mismo periodo)."""
    ruta = CONFIG / ARCHIVO_HISTORIAL
    periodo = f"{anio}-{mes:02d}"
    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    nuevos = []
    for fila in resultado:
        if str(fila.get("Estatus", "")).startswith("Cruzado"):
            for uuid in str(fila.get("UUID", "")).split(","):
                if uuid.strip():
                    nuevos.append({"UUID": uuid.strip(), "Periodo": periodo, "IDProveedor": fila["IDProveedor"],
                                   "Linea": str(fila["Linea"]), "Importe": str(fila["Importe"]),
                                   "Registrado": ahora})
    try:
        previo = (pd.read_csv(ruta, dtype=str, encoding="utf-8-sig").fillna("") if ruta.exists()
                  else pd.DataFrame(columns=["UUID", "Periodo", "IDProveedor", "Linea", "Importe", "Registrado"]))
        previo = previo[previo["Periodo"] != periodo]
        total = pd.concat([previo, pd.DataFrame(nuevos)], ignore_index=True).drop_duplicates(["UUID", "Periodo"])
        guardar_seguro(ruta, lambda tmp: total.to_csv(tmp, index=False, encoding="utf-8-sig"))
    except PermissionError:
        return f"No se pudo guardar {ARCHIVO_HISTORIAL} (¿está abierto?). Vuelve a correr al cerrarlo."
    return f"{len({n['UUID'] for n in nuevos})} UUID registrados para {periodo} en {ARCHIVO_HISTORIAL}"


# ------------------------------------------------------------------
# 3. CRUCE
# ------------------------------------------------------------------
# Pago quebrado = cualquier grupo de N líneas del TXT contra M facturas del MISMO proveedor
# (N ó M pueden ser 1, pero no ambos: eso es el cruce uno a uno).
MAX_PIEZAS_GRUPO = 10          # máximo de piezas (líneas + facturas) en un pago quebrado
MAX_COMBINACIONES = 400_000    # presupuesto de combinaciones por búsqueda (acota el tiempo)
MAX_CANDIDATOS = 200           # grupos alternativos que se guardan para detectar ambigüedad

CAMPOS_XML = ("SubTotal", "Descuento", "IVA", "RetISR", "RetIVA", "Total")


def a_centavos(valor):
    return int(redondear(valor) * 100)


def buscar_grupo(lineas, xmls, tolerancia, preferencia=lambda idxs: 0):
    """Busca el grupo más chico de líneas y facturas cuyas sumas coinciden dentro de la tolerancia.

    lineas, xmls: listas de (id, importe) de UN proveedor. Se prueban tamaños de grupo crecientes
    (3 piezas, luego 4...) y para cada tamaño todas las formas de repartirlo (1 línea + 2 XML,
    2 líneas + 1 XML...). Las sumas de un lado se ordenan y el otro lado se busca por bisección,
    así que cada combinación cuesta O(log n) y no O(n).

    Devuelve (ids_lineas, ids_xml, ambigua, truncada) o None si no hay grupo.
      ambigua  = hay otro grupo del mismo tamaño que comparte alguna pieza y tiene importes distintos
                 (si solo intercambia piezas del mismo importe, la elección no cambia el resultado).
      truncada = hubo repartos que no se exploraron por el presupuesto de combinaciones."""
    tol = a_centavos(tolerancia)
    L = [(i, v, a_centavos(v)) for i, v in lineas]
    X = [(i, v, a_centavos(v)) for i, v in xmls]
    truncada, trabajo = False, 0
    for k in range(3, min(MAX_PIEZAS_GRUPO, len(L) + len(X)) + 1):
        candidatos = []
        for s in range(1, k):
            t = k - s
            if s > len(L) or t > len(X):
                continue
            costo = comb(len(L), s) + comb(len(X), t)
            if trabajo + costo > MAX_COMBINACIONES:
                truncada = True
                continue
            trabajo += costo
            sumas_x = sorted((sum(c for *_, c in combo), combo) for combo in combinations(X, t))
            claves = [c for c, _ in sumas_x]
            for combo_l in combinations(L, s):
                suma_l = sum(c for *_, c in combo_l)
                # margen de 1 centavo por el redondeo a centavos; abajo se verifica con el importe exacto
                j = bisect_left(claves, suma_l - tol - 1)
                while j < len(claves) and claves[j] <= suma_l + tol + 1:
                    combo_x = sumas_x[j][1]
                    dif = sum(v for _, v, _ in combo_x) - sum(v for _, v, _ in combo_l)
                    if abs(dif) <= tolerancia:
                        ids_l, ids_x = tuple(i for i, *_ in combo_l), tuple(i for i, *_ in combo_x)
                        firma = (tuple(sorted(c for *_, c in combo_l)), tuple(sorted(c for *_, c in combo_x)))
                        candidatos.append(((preferencia(ids_x), abs(dif), ids_l, ids_x), ids_l, ids_x, firma))
                        if len(candidatos) >= MAX_CANDIDATOS:
                            break
                    j += 1
                if len(candidatos) >= MAX_CANDIDATOS:
                    break
        if candidatos:
            _, ids_l, ids_x, firma = min(candidatos)
            piezas = set(ids_l) | {("x", i) for i in ids_x}
            ambigua = any((set(l) | {("x", i) for i in x}) & piezas and f != firma
                          for _, l, x, f in candidatos)
            return list(ids_l), list(ids_x), ambigua, truncada
    return None


def validar_xml(f, ctx):
    """Reglas de negocio de una factura.
    Devuelve (motivos de rechazo, observaciones, motivos para detener, validaciones aprobadas)."""
    r, o, d, ok = [], [], [], set()

    # Timbre y versión
    if not f["UUID"]:
        r.append("El XML no tiene Timbre Fiscal Digital (no está timbrado)")
    if f["Version"] not in VERSIONES_CFDI_OK:
        r.append(f"CFDI versión {f['Version'] or 'desconocida'}, solo se acepta {', '.join(VERSIONES_CFDI_OK)}")

    # Receptor
    if f["RFCReceptor"] != RFC_SOCIEDAD:
        r.append(f"Facturado a {f['RFCReceptor'] or 'RFC vacío'}, debe ser {RFC_SOCIEDAD}")
    if REGIMEN_RECEPTOR and f["RegimenReceptor"] != REGIMEN_RECEPTOR:
        r.append(f"Régimen del receptor {f['RegimenReceptor'] or 'vacío'}, debe ser {REGIMEN_RECEPTOR} "
                 f"{REGIMENES.get(REGIMEN_RECEPTOR, '')}".strip())
    if CP_SOCIEDAD and f["CPReceptor"] != CP_SOCIEDAD:
        r.append(f"Código postal del receptor {f['CPReceptor'] or 'vacío'}, debe ser {CP_SOCIEDAD}")
    if NOMBRE_SOCIEDAD and normalizar_nombre(f["NombreReceptor"]) != normalizar_nombre(NOMBRE_SOCIEDAD):
        r.append(f"Nombre del receptor '{f['NombreReceptor']}', debe ser '{NOMBRE_SOCIEDAD}'")
    if f["UsoCFDI"] not in USOS_CFDI_OK:
        r.append(f"Uso de CFDI {f['UsoCFDI'] or 'vacío'}, debe ser {', '.join(sorted(USOS_CFDI_OK))}")

    # Emisor
    if not PATRON_RFC.fullmatch(f["RFCEmisor"]):
        r.append(f"RFC del emisor '{f['RFCEmisor']}' con formato inválido")
    elif len(f["RFCEmisor"]) != 13:
        d.append("El emisor es persona moral (RFC de 12 caracteres): las retenciones de honorarios "
                 "de persona física no aplican")

    # Periodo, tipo, método y forma de pago
    if f["FechaXML"][:7] != ctx["periodo"]:
        r.append(f"Factura del {f['FechaXML']}, fuera del periodo {ctx['mes']:02d}/{ctx['anio']}: "
                 "pedir refacturación")
    if f["TipoComprobante"] != "I":
        r.append(f"Tipo de comprobante {f['TipoComprobante']}, se esperaba I (Ingreso)")
    if f["MetodoPago"] not in METODOS_PAGO_OK:
        r.append(f"Método de pago {f['MetodoPago'] or 'vacío'}, debe ser PUE o PPD")
    if f["FormaPago"] not in FORMAS_PAGO_OK:
        nombre = FORMAS_PAGO.get(f["FormaPago"], "")
        r.append(f"Forma de pago {f['FormaPago'] or 'vacía'} {nombre}".strip()
                 + ", solo se acepta 03 Transferencia o 99 Por definir")
    if f["MetodoPago"] == "PPD" and f["FormaPago"] != "99":
        r.append(f"Método PPD con forma {f['FormaPago'] or 'vacía'}: en PPD la forma debe ser 99 Por definir")
    if f["MetodoPago"] == "PPD":
        o.append("Factura PPD: el doctor debe enviar el comprobante de pago (complemento de pago / REP) "
                 "después de recibir el pago")
    if f["Moneda"] != "MXN":
        d.append(f"Moneda {f['Moneda'] or 'vacía'}: los importes no son comparables contra el TXT en pesos")

    # Duplicados dentro del mes
    veces = ctx["uuids"].get(f["UUID"], 0)
    if f["UUID"] and veces > 1:
        texto = f"{MSG_UUID_REPETIDO}: el mismo UUID viene en {veces} archivos XML"
        (d if UUID_REPETIDO_SIEMPRE_DETIENE else o).append(texto)

    # Aritmética del propio XML
    if abs(f["DifTotalXML"]) > TOLERANCIA_IMPUESTOS:
        d.append(f"El Total del XML ({f['Total']:,.2f}) no cuadra con subtotal − descuento + impuestos − "
                 f"retenciones (diferencia {f['DifTotalXML']:,.2f})")
    if f["OtrosImpuestos"]:
        o.append(f"La factura trae otros impuestos (IEPS o locales) por {f['OtrosImpuestos']:,.2f}")

    # Conceptos
    if not f["ClavesProdServ"]:
        r.append("La factura no trae conceptos")
    elif abs(f["SumaConceptos"] - f["SubTotal"]) > TOLERANCIA_IMPUESTOS:
        d.append(f"Los conceptos suman {f['SumaConceptos']:,.2f} pero el SubTotal dice {f['SubTotal']:,.2f}: "
                 "XML alterado o mal emitido")
    malas = sorted({c for c in f["ClavesProdServ"] if not c.startswith(PREFIJOS_CLAVE_SERVICIOS)})
    if malas:
        o.append(f"Clave(s) de producto/servicio {', '.join(malas)} no corresponden a servicios profesionales")
    if f["IVA"] > 0 and f["ObjetosImp"] and all(x == "01" for x in f["ObjetosImp"]):
        o.append("Conceptos marcados 01 (no objeto de impuesto) pero la factura trae IVA")
    # Un concepto 02 con solo retención de ISR (honorarios E0, p. ej. médicos) es válido: no se avisa.
    # Si el TXT dice F2 y falta el IVA lo detiene el recálculo; si SAP tiene IH, la validación SAP.
    if f.get("Conceptos02SinImpuestos"):
        o.append(f"{f['Conceptos02SinImpuestos']} concepto(s) marcados 02 (objeto de impuesto) sin ningún "
                 "impuesto desglosado: XML mal emitido")
    ok.add("cfdi")

    # Historial de meses anteriores
    if ctx["historial"] is not None:
        previos = [x for x in ctx["historial"].get(f["UUID"], []) if x.get("Periodo") != ctx["periodo"]]
        if previos:
            p = previos[0]
            d.append(f"Este UUID ya se concilió en el periodo {p.get('Periodo')} (línea {p.get('Linea')} "
                     f"del TXT, importe {p.get('Importe')}): verificar que no se pague dos veces")
        ok.add("historial")
    return r, o, d, ok


def validar_sap(codigo_txt, regimen_xml, codigos, xml_con_iva):
    """Compara TXT (E0/F2), IVA y régimen del XML contra los códigos de retención de SAP.
    Devuelve (régimen según SAP, observaciones, motivos para detener la factura)."""
    obs, detener = [], []
    c = set(codigos)
    if (codigo_txt == "F2" or xml_con_iva) and "IH" not in c:
        origen = "el TXT dice F2" if codigo_txt == "F2" else "el XML trae IVA"
        detener.append(f"{MSG_SAP_SIN_IH} (lleva IVA porque {origen}, pero SAP no tiene IH)")
    if "IP" in c:
        detener.append("SAP tiene IP (Ret. IVA Pitex 100%): retendría todo el IVA, no corresponde a honorarios")
    if codigo_txt == "E0" and "IH" in c:
        obs.append("TXT dice E0 (sin IVA) pero SAP tiene IH (retención de IVA)")
    isr = c & set(ISR_SAP)
    if len(isr) != 1:
        detener.append(f"SAP con indicadores de ISR {sorted(isr) or 'ninguno'}, se esperaba uno: "
                       "no se sabe cuánto ISR retendrá")
        return "", obs, detener
    regimen_sap = ISR_SAP[isr.pop()]
    if regimen_sap == "Extranjero":
        detener.append("SAP marca al proveedor como Extranjero (RE) pero factura con CFDI nacional")
    for reg in [r.strip() for r in regimen_xml.split(",") if r.strip()]:
        if regimen_sap == "RESICO" and reg != "626":
            detener.append(f"SAP retendrá ISR RESICO 1.25% (RS) pero el XML trae régimen {reg}: "
                           "el pago no coincidirá con la factura")
        elif regimen_sap == "No RESICO" and reg == "626":
            detener.append("SAP retendrá ISR 10% (RH) pero el XML es RESICO 626: "
                           "el pago no coincidirá con la factura")
    return regimen_sap, obs, detener


def redondear(valor):
    """Redondea a 2 decimales igual que ROUND de Excel (0.005 sube)."""
    return valor.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def revisar_impuestos(fila):
    """Recalcula IVA y retenciones de una línea cruzada (mismas reglas que la hoja Cálculos).
    Devuelve el motivo para detenerla, o "" si la variación está dentro de la tolerancia."""
    base = fila["SubTotal"] - fila["Descuento"]
    con_iva = fila["CodigoImpuesto"] == "F2"
    regimenes = {r.strip() for r in fila.get("RegimenEmisor", "").split(",") if r.strip()}
    resico = regimenes == {"626"}
    iva_calc = redondear(base * TASA_IVA) if con_iva else Decimal("0")
    isr_calc = redondear(base * (TASA_ISR_RESICO if resico else TASA_ISR_GENERAL))
    riva_calc = redondear(iva_calc * 2 / 3)
    conceptos = [
        ("IVA", fila["IVA"], iva_calc, "16%" if con_iva else "el TXT dice E0: sin IVA"),
        ("Ret. ISR", fila["RetISR"], isr_calc, "RESICO 1.25%" if resico else "10%"),
        ("Ret. IVA", fila["RetIVA"], riva_calc, "2/3 del IVA" if con_iva else "sin IVA no se retiene"),
    ]
    variacion = sum(abs(xml - calc) for _, xml, calc, _ in conceptos)
    if variacion <= TOLERANCIA_IMPUESTOS:
        return ""
    partes = [f"{nombre} {xml:,.2f}, debería ser {calc:,.2f} ({regla})"
              for nombre, xml, calc, regla in conceptos if abs(xml - calc) > UMBRAL_OK]
    if fila["RetIVA"] > 0 and fila["IVA"] == 0:
        partes.append("retiene IVA sin cobrar IVA (¿capturó la retención de ISR como IVA?)")
    tasas = {a_decimal(t) for t in fila.get("_tasas_iva", set())} - {None}
    if any(t != TASA_IVA for t in tasas):
        partes.append("el XML usa tasa de IVA " + ", ".join(f"{t:.0%}" for t in sorted(tasas))
                      + " (¿tasa fronteriza 8%?)")
    if len(regimenes) > 1:
        partes.append(f"las facturas traen regímenes distintos ({', '.join(sorted(regimenes))})")
    return f"Impuestos no cuadran (variación {variacion:,.2f}): " + "; ".join(partes)


def conciliar(lineas_txt, catalogo, facturas, anio, mes, sap=None, historial=None):
    por_id = catalogo.set_index("IDProveedor").to_dict("index")
    uuids_vistos = {}
    for f in facturas:
        if f["UUID"]:
            uuids_vistos[f["UUID"]] = uuids_vistos.get(f["UUID"], 0) + 1
    ctx = {"uuids": uuids_vistos, "anio": anio, "mes": mes, "periodo": f"{anio}-{mes:02d}",
           "historial": historial}
    validacion = [validar_xml(f, ctx) for f in facturas]      # se calcula una vez por factura

    xml_por_rfc = {}
    for i, f in enumerate(facturas):
        xml_por_rfc.setdefault(f["RFCEmisor"], []).append(i)

    filas, obs, rech, deten, pasadas = [], [], [], [], []
    pendientes = {}                 # RFC -> índices de líneas aún sin XML
    repetidas = {}                  # (proveedor, folio, importe) -> líneas
    for n, linea in enumerate(lineas_txt):
        fila = {k: linea[k] for k in ("Linea", "Fecha", "IDProveedor", "Folio", "Importe",
                                      "CodigoImpuesto", "CentroCosto", "Cuenta")}
        fila["Estatus"] = "Sin XML"
        o, d = [], []
        if linea["CodigoImpuesto"] not in CODIGOS_IMPUESTO_TXT:
            d.append(f"Código de impuesto '{linea['CodigoImpuesto']}' desconocido en el TXT (se espera E0 o F2)")
        if FECHA_TXT_DEBE_SER_DEL_PERIODO and linea["_fecha"].strftime("%Y-%m") != ctx["periodo"]:
            o.append(f"La fecha del TXT ({linea['Fecha']}) no es del periodo {mes:02d}/{anio}")
        repetidas.setdefault((linea["IDProveedor"], linea["Folio"], linea["Importe"]), []).append(n)

        prov = por_id.get(linea["IDProveedor"])
        if not linea["IDProveedor"]:
            fila.update(Estatus="Línea sin IDProveedor", RFC="", Nombre="")
            d.append("La línea del TXT no trae IDProveedor: no se puede saber a quién se le paga")
        elif prov is None:
            fila.update(Estatus="ID no está en catálogo", RFC="", Nombre="")
        else:
            fila["RFC"], fila["Nombre"] = prov["RFC"], prov["Nombre"]
            if prov.get("Activo") == "0":
                d.append("Proveedor bloqueado (inactivo) en el catálogo")
            tipo = prov.get("TipoRFC") or clasificar_rfc(prov["RFC"])
            if tipo == "Vacío":
                fila["Estatus"] = "Proveedor sin RFC en catálogo"
            elif tipo == "Extranjero":
                fila["Estatus"] = "Proveedor extranjero"
                o.append(f"ID fiscal extranjero '{prov['RFC']}': no emite CFDI, se revisa con su factura "
                         "del extranjero (y retención RE en SAP)")
            elif tipo == "Inválido":
                fila["Estatus"] = "RFC inválido en catálogo"
                d.append(f"RFC '{prov['RFC']}' mal escrito en el catálogo: corregirlo para poder cruzar")
            elif linea["Importe"] <= 0:
                d.append(f"Importe {linea['Importe']:,.2f} en el TXT: debe ser mayor a cero")
            else:
                pendientes.setdefault(prov["RFC"], []).append(n)
        filas.append(fila)
        obs.append(o)
        rech.append([])
        deten.append(d)
        pasadas.append(set())

    for ns in repetidas.values():
        if len(ns) > 1:
            nums = ", ".join(str(lineas_txt[n]["Linea"]) for n in ns)
            for n in ns:
                deten[n].append(f"Líneas {nums} del TXT repetidas (mismo proveedor, folio e importe): "
                                "posible pago doble")

    usadas = set()
    libres = lambda rfc: [i for i in xml_por_rfc.get(rfc, []) if i not in usadas]

    def poner_xml(n, idxs, estatus, notas=(), con_importes=True):
        fs = [facturas[i] for i in idxs]
        fila = filas[n]
        fila["Estatus"] = estatus
        fila["_grupo"] = tuple(sorted(idxs))
        fila["ArchivoXML"] = ", ".join(f["ArchivoXML"] for f in fs)
        fila["UUID"] = ", ".join(f["UUID"] for f in fs)
        fila["FechaXML"] = ", ".join(sorted({f["FechaXML"] for f in fs}))
        fila["RegimenEmisor"] = ", ".join(sorted({f["RegimenEmisor"] for f in fs}))
        if con_importes:
            for c in CAMPOS_XML:
                fila[c] = sum((f[c] for f in fs), Decimal("0"))
        fila["_iva_xml"] = any(f["IVA"] > 0 for f in fs)
        fila["_tasas_iva"] = set().union(*(f["TasasIVA"] for f in fs))
        fila["NombreRegimen"] = ", ".join(sorted({REGIMENES.get(f["RegimenEmisor"], "Desconocido")
                                                  for f in fs}))
        fila["MetodoPago"] = ", ".join(sorted({f["MetodoPago"] for f in fs}))
        fila["FormaPago"] = ", ".join(sorted({f"{f['FormaPago']} {FORMAS_PAGO.get(f['FormaPago'], '')}".strip()
                                              for f in fs}))
        aprobadas = None
        for i in idxs:
            r, o, d, ok = validacion[i]
            rech[n].extend(r)
            obs[n].extend(o)
            deten[n].extend(d)
            aprobadas = set(ok) if aprobadas is None else aprobadas & ok   # debe aprobar en TODAS las facturas
        pasadas[n] |= aprobadas
        obs[n].extend(notas)

    def calidad(i):
        """0 = factura válida, 1 = con motivo para detener, 2 = a rechazar. Ante importes iguales se
        prefiere la válida: p. ej. el proveedor canceló y reemitió por el mismo importe."""
        r, _, d, _ = validacion[i]
        return 2 if r else (1 if d else 0)

    def emparejar(lineas, candidatas_de):
        """Emparejamiento máximo línea <-> XML (algoritmo de Kuhn). A diferencia de asignar en orden,
        ninguna línea se queda sin factura si existe un acomodo en el que todas cruzan.
        candidatas_de(n) devuelve los XML posibles ya ordenados por preferencia."""
        asignada = {}                                  # xml -> línea

        def intentar(n, vistos):
            for i in candidatas_de(n):
                if i not in vistos:
                    vistos.add(i)
                    if i not in asignada or intentar(asignada[i], vistos):
                        asignada[i] = n
                        return True
            return False

        for n in lineas:
            intentar(n, set())
        return {n: i for i, n in asignada.items()}

    def uno_a_uno(limite):
        """Una línea = una factura, con emparejamiento máximo por proveedor."""
        for rfc, lineas in pendientes.items():
            pool = libres(rfc)
            cerca = lambda n, lim: [i for i in pool
                                    if abs(facturas[i]["ImporteConIVA"] - lineas_txt[n]["Importe"]) <= lim]
            orden = lambda n: sorted(cerca(n, limite), key=lambda i: (
                calidad(i), abs(facturas[i]["ImporteConIVA"] - lineas_txt[n]["Importe"])))
            asignacion = emparejar(list(lineas), orden)
            tomadas = set(asignacion.values())
            for n in sorted(asignacion):
                i = asignacion[n]
                otras = [j for j in cerca(n, TOLERANCIA) if j != i and j not in tomadas
                         and facturas[j]["UUID"] != facturas[i]["UUID"]]     # copia del mismo UUID: otro aviso
                notas = ([f"{len(otras) + 1} facturas distintas (UUID diferente) del mismo RFC por el mismo "
                          "importe: posible facturación doble del proveedor"]
                         if otras else [])
                usadas.add(i)
                poner_xml(n, [i], "Cruzado", notas)
                lineas.remove(n)

    def poner_grupo(ns, idxs, notas):
        ns = sorted(ns, key=lambda n: lineas_txt[n]["Linea"])
        usadas.update(idxs)
        for k, n in enumerate(ns):
            poner_xml(n, idxs, "Cruzado (pago quebrado)", notas, con_importes=(k == 0))
            pendientes[filas[n]["RFC"]].remove(n)

    def describir(ns, idxs):
        nums = ", ".join(str(lineas_txt[n]["Linea"]) for n in sorted(ns, key=lambda n: lineas_txt[n]["Linea"]))
        archivos = ", ".join(facturas[i]["ArchivoXML"] for i in idxs)
        total_txt = sum((lineas_txt[n]["Importe"] for n in ns), Decimal("0"))
        total_xml = sum((facturas[i]["ImporteConIVA"] for i in idxs), Decimal("0"))
        return (f"Pago quebrado: {len(ns)} línea(s) del TXT ({nums}) por {total_txt:,.2f} contra "
                f"{len(idxs)} factura(s) ({archivos}) por {total_xml:,.2f}")

    def lineas_por_id(lineas):
        """El ID del proveedor delimita el grupo: nunca se mezclan líneas de IDs distintos,
        aunque en el catálogo compartan RFC."""
        por_id = {}
        for n in lineas:
            por_id.setdefault(lineas_txt[n]["IDProveedor"], []).append(n)
        return por_id.values()

    def agrupados(tolerancia):
        """Pago quebrado N:M. Se repite por proveedor hasta que ya no aparece ningún grupo."""
        for rfc in pendientes:
            for grupo_id in lineas_por_id(list(pendientes[rfc])):
                grupo_id = list(grupo_id)
                while grupo_id:
                    pool = [i for i in libres(rfc) if facturas[i]["ImporteConIVA"] > 0]
                    if not pool or len(grupo_id) + len(pool) < 3:
                        break
                    hallado = buscar_grupo([(n, lineas_txt[n]["Importe"]) for n in grupo_id],
                                           [(i, facturas[i]["ImporteConIVA"]) for i in pool],
                                           tolerancia, lambda idxs: sum(calidad(i) for i in idxs))
                    if hallado is None:
                        break
                    ns, idxs, ambigua, truncada = hallado
                    notas = [describir(ns, idxs)]
                    if ambigua:
                        notas.append("Hay más de una combinación de líneas y facturas que suma lo mismo: "
                                     "confirmar qué pago corresponde a qué factura")
                    if truncada:
                        notas.append(f"El proveedor tiene muchas piezas pendientes; la búsqueda se limitó a "
                                     f"grupos de hasta {MAX_PIEZAS_GRUPO} y pudo no encontrar el más chico")
                    poner_grupo(ns, idxs, notas)
                    grupo_id = [n for n in grupo_id if n not in ns]

    def por_total_del_proveedor():
        """Red de seguridad: si lo pendiente del TXT de un proveedor suma lo mismo que sus facturas
        pendientes, cuadra en total aunque no se sepa qué pago corresponde a qué factura."""
        for rfc, lineas in pendientes.items():
            ids = {lineas_txt[n]["IDProveedor"] for n in lineas}
            pool = [i for i in libres(rfc) if facturas[i]["ImporteConIVA"] > 0]
            if len(ids) != 1 or not pool or len(lineas) + len(pool) < 3:
                continue
            ns = list(lineas)
            dif = (sum(facturas[i]["ImporteConIVA"] for i in pool) - sum(lineas_txt[n]["Importe"] for n in ns))
            if abs(dif) <= TOLERANCIA:
                poner_grupo(ns, pool, [describir(ns, pool),
                                       "Cuadre por total del proveedor: no se pudo determinar qué pago "
                                       "corresponde a qué factura, revisar el reparto"])

    # Orden de los cruces: primero la evidencia más fuerte (exacto) y los grupos más chicos.
    #   1. Uno a uno exacto (hasta 0.05)
    #   2. Pago quebrado exacto (N líneas contra M facturas, hasta 0.05)
    #   3. Uno a uno con tolerancia (hasta 1.00)
    #   4. Pago quebrado con tolerancia (hasta 1.00): queda con observación por la diferencia
    #   5. Total del proveedor
    uno_a_uno(UMBRAL_OK)
    agrupados(UMBRAL_OK)
    uno_a_uno(TOLERANCIA)
    agrupados(TOLERANCIA)
    por_total_del_proveedor()

    # Lo que quedó sin cruzar
    for rfc, lineas in pendientes.items():
        disponibles = libres(rfc)
        pendiente_txt = sum((lineas_txt[n]["Importe"] for n in lineas), Decimal("0"))
        pendiente_xml = sum((facturas[i]["ImporteConIVA"] for i in disponibles), Decimal("0"))
        for n in lineas:
            imp = lineas_txt[n]["Importe"]
            if disponibles:
                i = min(disponibles, key=lambda i: abs(facturas[i]["ImporteConIVA"] - imp))
                filas[n]["Estatus"] = "Diferencia de importe"
                obs[n].append(f"XML más cercano: {facturas[i]['ArchivoXML']} por "
                              f"{facturas[i]['ImporteConIVA']:,.2f} "
                              f"(diferencia {facturas[i]['ImporteConIVA'] - imp:,.2f})")
                obs[n].append(f"Pendiente del proveedor: TXT {pendiente_txt:,.2f} en {len(lineas)} línea(s) "
                              f"contra XML {pendiente_xml:,.2f} en {len(disponibles)} factura(s) "
                              f"(diferencia {pendiente_xml - pendiente_txt:,.2f})")
            else:
                filas[n]["Estatus"] = "Sin XML"

    # Por cada factura (o grupo de facturas) cruzada: diferencia de importe e impuestos.
    # En pago quebrado todas las líneas del grupo comparten el resultado.
    grupos = {}
    for n, fila in enumerate(filas):
        if "_grupo" in fila:
            grupos.setdefault(fila["_grupo"], []).append(n)
    for grupo, ns in grupos.items():
        dif = (sum(facturas[i]["ImporteConIVA"] for i in grupo)
               - sum(lineas_txt[n]["Importe"] for n in ns))
        if DIFERENCIA_IMPORTE_ES_OBSERVACION and abs(dif) > UMBRAL_OK:
            for n in ns:
                obs[n].append(f"Diferencia de importe de {dif:,.2f} entre factura y TXT "
                              f"(dentro de la tolerancia de {TOLERANCIA})")
        con_importes = next(n for n in ns if filas[n].get("SubTotal") is not None)
        motivo = revisar_impuestos(filas[con_importes])
        codigos = sorted({lineas_txt[n]["CodigoImpuesto"] for n in ns})
        if len(codigos) > 1:
            nums = ", ".join(str(lineas_txt[n]["Linea"]) for n in ns)
            motivo = "; ".join(filter(None, [motivo, f"Las líneas {nums} pagan la misma factura con códigos de "
                                                     f"impuesto distintos ({', '.join(codigos)})"]))
        for n in ns:
            pasadas[n].add("impuestos")
            if motivo:
                deten[n].append(motivo)

    # Pago doble real: el mismo UUID quedó usado en más de un XML cruzado (dos líneas pagan la misma factura)
    usos_uuid = {}
    for n, fila in enumerate(filas):
        for i in fila.get("_grupo", ()):
            if facturas[i]["UUID"]:
                usos_uuid.setdefault(facturas[i]["UUID"], set()).add(i)
    for n, fila in enumerate(filas):
        dobles = [u for u, idxs in usos_uuid.items()
                  if len(idxs) > 1 and any(facturas[i]["UUID"] == u for i in fila.get("_grupo", ()))]
        if dobles:
            lineas_doble = sorted({lineas_txt[m]["Linea"] for m, x in enumerate(filas)
                                   if any(facturas[i]["UUID"] in dobles for i in x.get("_grupo", ()))})
            deten[n].append(f"Pago doble: la misma factura (UUID {', '.join(dobles)}) la pagan las líneas "
                            f"{', '.join(map(str, lineas_doble))} del TXT. {MSG_UUID_REPETIDO}")

    # Validación contra SAP
    for n, fila in enumerate(filas):
        if "_grupo" not in fila:
            continue
        if sap is None:
            continue                                   # no se pidió validar SAP en esta corrida
        consulta = sap.get(fila["IDProveedor"])
        if consulta is None:
            continue                                   # no se consultó: queda como validación no realizada
        codigos, error = consulta
        fila["CodigosSAP"] = ", ".join(codigos)
        if error:
            obs[n].append(f"No se pudo validar en SAP: {error}")
            continue
        fila["RegimenSAP"], o, d = validar_sap(lineas_txt[n]["CodigoImpuesto"], fila.get("RegimenEmisor", ""),
                                               codigos, fila.get("_iva_xml", False))
        obs[n].extend(o)
        deten[n].extend(d)
        pasadas[n].add("sap")

    # Estatus final. "Cruzado" solo si no hay rechazos, ni motivos para detener, ni observaciones
    # que requieran revisión, y además se aprobaron TODAS las validaciones de la lista.
    for n, (fila, o, r, d) in enumerate(zip(filas, obs, rech, deten)):
        cruzada = fila["Estatus"].startswith("Cruzado")
        if cruzada:
            faltan = [nombre for clave, nombre in VALIDACIONES.items()
                      if clave not in pasadas[n] and not (clave == "sap" and sap is None)]
            if faltan and not r:                     # si ya se rechaza, el faltante no aporta
                o.append("Validaciones no realizadas: " + ", ".join(faltan))
        fila["DiscrepanciaSAP"] = any("SAP" in x and not x.startswith(NO_ES_DISCREPANCIA_SAP) for x in o + d)
        if cruzada and r:
            fila["Estatus"] = "Rechazar"
        elif cruzada and d:
            fila["Estatus"] = "Detenida"
        elif cruzada and any(not x.startswith(INFORMATIVAS) for x in o):
            fila["Estatus"] = "Cruzado con observaciones"
        elif not cruzada and d and lineas_txt[n]["Importe"] <= 0:
            fila["Estatus"] = "Detenida"             # el propio TXT está mal; en los demás casos se conserva
                                                     # el estatus (Sin XML, etc.) y el motivo va en MotivoRechazo
        fila["MotivoRechazo"] = "; ".join(dict.fromkeys(r + d))
        fila["Observaciones"] = "; ".join(dict.fromkeys(o))

    # XML que ninguna línea del TXT usó
    rfc_a_id = {}
    for i, r in zip(catalogo["IDProveedor"], catalogo["RFC"]):
        if r:
            rfc_a_id.setdefault(r, []).append(i)
    sobrantes = []
    for i, f in enumerate(facturas):
        if i in usadas:
            continue
        s = {k: f[k] for k in ("ArchivoXML", "UUID", "FechaXML", "RFCEmisor", "NombreEmisor",
                               "RegimenEmisor", "RFCReceptor", "MetodoPago", "FormaPago",
                               "SubTotal", "Descuento", "IVA", "RetISR", "RetIVA", "Total",
                               "ImporteConIVA")}
        s["NombreRegimen"] = REGIMENES.get(f["RegimenEmisor"], "Desconocido")
        r, o_xml, d, _ = validacion[i]
        motivos = r + d
        copia = sorted(facturas[j]["ArchivoXML"] for j in usadas
                       if j != i and f["UUID"] and facturas[j]["UUID"] == f["UUID"])
        if copia:                                     # la otra copia ya se cruzó: esta sobra
            motivos.append(f"Copia repetida de {', '.join(copia)} (ya cruzado): {MSG_UUID_REPETIDO}")
        else:
            motivos += [x for x in o_xml if x.startswith(MSG_UUID_REPETIDO)]
        s["MotivoRechazo"] = "; ".join(motivos)
        s["IDProveedorCatalogo"] = ", ".join(rfc_a_id.get(f["RFCEmisor"], ["No está en catálogo"]))
        sobrantes.append(s)

    verificar_integridad(lineas_txt, facturas, filas, sobrantes)
    return filas, sobrantes


def verificar_integridad(lineas_txt, facturas, filas, sobrantes):
    """Controles de cuadre del propio programa. Si alguno falla es un error de lógica, no de los datos:
    se detiene ANTES de generar un reporte o tocar el historial."""
    fallas = []
    if sorted(f["Linea"] for f in filas) != sorted(l["Linea"] for l in lineas_txt):
        fallas.append("el resultado no tiene exactamente una fila por cada línea del TXT")
    grupos = {}
    for n, fila in enumerate(filas):
        if "_grupo" in fila:
            grupos.setdefault(fila["_grupo"], []).append(n)
    usados = [i for g in grupos for i in g]
    if len(usados) != len(set(usados)):
        fallas.append("un mismo XML quedó asignado a dos cruces distintos")
    if len(usados) + len(sobrantes) != len(facturas):
        fallas.append(f"XML usados ({len(usados)}) + sin TXT ({len(sobrantes)}) no suman los leídos "
                      f"({len(facturas)})")
    for g, ns in grupos.items():
        dif = (sum(facturas[i]["ImporteConIVA"] for i in g) - sum(lineas_txt[n]["Importe"] for n in ns))
        if abs(dif) > TOLERANCIA:
            fallas.append(f"el cruce de las líneas {[lineas_txt[n]['Linea'] for n in ns]} difiere {dif}")
        if len({lineas_txt[n]["IDProveedor"] for n in ns}) != 1 \
                or {facturas[i]["RFCEmisor"] for i in g} != {filas[ns[0]]["RFC"]}:
            fallas.append(f"el cruce de las líneas {[lineas_txt[n]['Linea'] for n in ns]} mezcla proveedores")
    for fila in filas:
        if fila["Estatus"] in ("Cruzado", "Cruzado (pago quebrado)") \
                and fila["MotivoRechazo"]:
            fallas.append(f"la línea {fila['Linea']} quedó Cruzado con motivo de rechazo")
    if fallas:
        raise RuntimeError("Falló el control de integridad de la conciliación:\n  " + "\n  ".join(fallas))



# ------------------------------------------------------------------
# 4. REPORTE
# ------------------------------------------------------------------
ORDEN_ESTATUS = ["Cruzado", "Cruzado (pago quebrado)",
                 "Cruzado con observaciones", "Detenida", "Rechazar", "Diferencia de importe", "Sin XML",
                 "Proveedor extranjero", "RFC inválido en catálogo",
                 "Proveedor sin RFC en catálogo", "ID no está en catálogo", "Línea sin IDProveedor"]
COLORES = {"Cruzado": "C6EFCE", "Cruzado (pago quebrado)": "C6EFCE",
           "Cruzado con observaciones": "FFEB9C",
           "Diferencia de importe": "FFEB9C", "Proveedor extranjero": "DDEBF7"}
COLORES["Detenida"] = "F8CBAD"
ROJO = "FFC7CE"


def a_float(df):
    for col in df.columns:
        if df[col].map(lambda v: isinstance(v, Decimal)).any():
            df[col] = df[col].map(lambda v: float(v) if isinstance(v, Decimal) else v)
    return df


def formatear_hoja(hoja, colorear_estatus=False):
    hoja.freeze_panes = "A2"
    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="305496")
    for i, col in enumerate(hoja.columns, start=1):
        ancho = max(len(str(c.value or "")) for c in col)
        hoja.column_dimensions[get_column_letter(i)].width = min(max(ancho + 2, 10), 60)
        for c in col[1:]:
            if isinstance(c.value, float):
                c.number_format = "#,##0.00"
    if colorear_estatus:
        encabezados = [c.value for c in hoja[1]]
        if "Estatus" in encabezados:
            col = encabezados.index("Estatus")
            for fila in hoja.iter_rows(min_row=2):
                color = COLORES.get(fila[col].value, ROJO)
                fila[col].fill = PatternFill("solid", fgColor=color)


def escribir_resumen(hoja, detalle_cols, n_filas, n_xml):
    """Hoja Resumen con fórmulas que leen la hoja Detalle (se recalcula sola en Excel)."""
    col = lambda nombre: get_column_letter(detalle_cols.index(nombre) + 1)
    E, I, S = col("Estatus"), col("Importe"), col("DiscrepanciaSAP")
    R, V = col("Revisión impuestos"), col("Variación impuestos")
    RI, VI = col("Revisión importe"), col("Variación importe")
    OB = col("Observaciones")
    rango = lambda c: f"Detalle!${c}$2:${c}${n_filas + 1}"

    hoja.append(["Estatus", "Líneas", "Importe TXT", "% del importe", "Cómo se calcula"])
    primera = 2
    for estatus in ORDEN_ESTATUS:
        r = hoja.max_row + 1
        hoja.append([estatus,
                     f"=COUNTIF({rango(E)},A{r})",
                     f"=SUMIF({rango(E)},A{r},{rango(I)})",
                     f"=IF($C${primera + len(ORDEN_ESTATUS)}=0,0,C{r}/$C${primera + len(ORDEN_ESTATUS)})",
                     f"Líneas del Detalle con Estatus = \"{estatus}\""])
    ultima = hoja.max_row
    r_total = ultima + 1
    hoja.append(["TOTAL TXT", f"=SUM(B{primera}:B{ultima})", f"=SUM(C{primera}:C{ultima})",
                 f"=SUM(D{primera}:D{ultima})", "Suma de los estatus de arriba"])
    hoja.append(["Cuadre contra Detalle", f"=COUNTA({rango(E)})-B{r_total}",
                 f"=SUM({rango(I)})-C{r_total}", None,
                 "Debe ser 0: todas las líneas del TXT tienen un estatus"])
    hoja.append([])
    hoja.append(["Otros indicadores", "Cantidad", None, None, "Cómo se calcula"])
    extras = [
        ("Líneas con validaciones no realizadas", f"=COUNTIF({rango(OB)},\"*Validaciones no realizadas*\")",
         "SAP (si se pidió) o el historial de UUID no se pudieron consultar para esas líneas"),
        ("Líneas con discrepancia SAP", f"=COUNTIF({rango(S)},\"Sí\")",
         "Líneas del Detalle con DiscrepanciaSAP = Sí"),
        ("Líneas con impuestos a revisar", f"=COUNTIF({rango(R)},\"Revisar\")",
         "Variación de impuestos mayor a la tolerancia (ver hoja Cálculos)"),
        ("Líneas con redondeo en impuestos", f"=COUNTIF({rango(R)},\"Redondeo\")",
         "Variación de impuestos dentro de la tolerancia de redondeo (hoja Tasas)"),
        ("Variaciones de impuestos por redondeo", f"=SUMIF({rango(R)},\"Redondeo\",{rango(V)})",
         "Suma de 'Variación impuestos' de las líneas con Redondeo"),
        ("Variaciones de impuestos a revisar", f"=SUMIF({rango(R)},\"Revisar\",{rango(V)})",
         "Suma de 'Variación impuestos' de las líneas con Revisar"),
        ("Líneas con redondeo en importe", f"=COUNTIF({rango(RI)},\"Redondeo\")",
         "Diferencia entre TXT y factura dentro de la tolerancia de importe"),
        ("Neto de variaciones de importe", f"=SUM({rango(VI)})",
         "Suma de 'Variación importe' (factura − TXT). Positivo: las facturas suman más que el TXT"),
        ("Variaciones de importe sin signo", f"=SUMPRODUCT(ABS({rango(VI)}))",
         "Suma de las variaciones de importe sin importar si son a favor o en contra"),
        ("XML leídos", n_xml, "Archivos .xml encontrados en la carpeta XML (dato del programa)"),
        ("XML sin línea en TXT", "=MAX(COUNTA('XML sin TXT'!A:A)-1,0)", "Filas de la hoja 'XML sin TXT'"),
        ("XML que no se pudieron leer", "=MAX(COUNTA('Errores XML'!A:A)-1,0)", "Filas de la hoja 'Errores XML'"),
    ]
    for nombre, formula, explica in extras:
        hoja.append([nombre, formula, None, None, explica])

    # Formato
    azul = PatternFill("solid", fgColor="305496")
    for fila in (1, r_total + 3):
        for c in hoja[fila]:
            if c.value is not None:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = azul
    for fila in (r_total, r_total + 1):
        for c in hoja[fila]:
            c.font = Font(bold=True)
    for r in range(primera, r_total + 2):
        hoja[f"C{r}"].number_format = "#,##0.00"
        hoja[f"D{r}"].number_format = "0.0%"
    for r in range(primera, ultima + 1):
        estatus = hoja[f"A{r}"].value
        hoja[f"A{r}"].fill = PatternFill("solid", fgColor=COLORES.get(estatus, ROJO))
    for letra, ancho in zip("ABCDE", (38, 10, 16, 14, 62)):
        hoja.column_dimensions[letra].width = ancho
    hoja.freeze_panes = "A2"


# ------------------------------------------------------------------
# 4b. DESGLOSE DE IMPUESTOS CON FÓRMULAS DE EXCEL
#     Todo lo que se escribe aquí son fórmulas: quien abre el Excel da clic
#     en una celda y ve exactamente cómo se calculó (=C8*TasaIVA, etc.).
# ------------------------------------------------------------------
TASAS = [  # (nombre en Excel, concepto, valor, explicación)
    ("TasaIVA", "IVA", float(TASA_IVA), "Se cobra sobre la base (subtotal menos descuento)"),
    ("TasaISR_General", "Retención de ISR (honorarios)", float(TASA_ISR_GENERAL), "Personas físicas con honorarios, régimen 612"),
    ("TasaISR_RESICO", "Retención de ISR (RESICO)", float(TASA_ISR_RESICO), "Solo cuando la factura trae régimen 626"),
    ("ProporcionRetIVA", "Parte del IVA que se retiene", "=2/3", "Dos terceras partes del IVA (10.67% de la base)"),
    ("Tolerancia", "Variación que se considera OK", float(UMBRAL_OK),
     "Hasta esta cantidad la línea sale OK (centavos de redondeo)"),
    ("ToleranciaImpuestos", "Variación aceptada por redondeo", float(TOLERANCIA_IMPUESTOS),
     "Hasta esta cantidad sale 'Redondeo' (se acepta); arriba sale 'Revisar'"),
    ("ToleranciaImporte", "Diferencia de importe aceptada (TXT vs XML)", float(TOLERANCIA),
     "Diferencia máxima entre el TXT y la factura. Cambiarla aquí solo cambia las etiquetas; "
     "el cruce usa TOLERANCIA del programa"),
]

# Fórmula de la etiqueta según la variación (se usa en Detalle y en Cálculos)
ETIQUETA_VARIACION = '=IF({v}="","",IF({v}<=Tolerancia,"OK",IF({v}<=ToleranciaImpuestos,"Redondeo","Revisar")))'
ETIQUETA_IMPORTE = ('=IF({v}="","",IF(ABS({v})<=Tolerancia,"OK",'
                    'IF(ABS({v})<=ToleranciaImporte,"Redondeo","Revisar")))')

# Columnas del desglose en la hoja Detalle (van agrupadas: se abren con el botón +)
COLS_DESGLOSE = ["SubTotal", "Descuento", "Base", "IVA", "IVA calculado", "Dif IVA",
                 "RetISR", "RetISR calculado", "Dif RetISR", "RetIVA", "RetIVA calculado", "Dif RetIVA",
                 "Total", "Total calculado", "Dif Total"]

AZUL = PatternFill("solid", fgColor="305496")
GRIS = PatternFill("solid", fgColor="D9E1F2")
AMARILLO = PatternFill("solid", fgColor="FFF2CC")
FMT_PESOS = "#,##0.00;-#,##0.00;-"
LIGA = Font(color="0563C1", underline="single")


def ligar(celda, hoja, coordenada, texto):
    """Convierte la celda en un vínculo a otra hoja del mismo libro."""
    from openpyxl.worksheet.hyperlink import Hyperlink
    celda.value = texto
    celda.hyperlink = Hyperlink(ref=celda.coordinate, location=f"'{hoja}'!{coordenada}")
    celda.font = LIGA


def escribir_tasas(libro):
    """Hoja Tasas: los porcentajes en celdas amarillas, con nombre, para usarlos en las fórmulas."""
    from openpyxl.workbook.defined_name import DefinedName
    hoja = libro.create_sheet("Tasas")
    hoja.append(["Concepto", "Valor", "Nombre en fórmulas", "Explicación"])
    for nombre, concepto, valor, explica in TASAS:
        hoja.append([concepto, valor, nombre, explica])
        r = hoja.max_row
        hoja[f"B{r}"].fill = AMARILLO
        hoja[f"B{r}"].font = Font(color="0000FF", bold=True)
        hoja[f"B{r}"].number_format = "0.00" if nombre.startswith("Tolerancia") else "0.00%"
        libro.defined_names[nombre] = DefinedName(nombre, attr_text=f"Tasas!$B${r}")
    hoja.append([])
    hoja.append(["Las celdas amarillas se pueden cambiar para simular; las columnas de cálculo se recalculan solas."])
    hoja.append(["Ojo: el Estatus (Cruzado / Detenida) lo decide el programa con las tasas de su configuración."])
    for c in hoja[1]:
        c.font, c.fill = Font(bold=True, color="FFFFFF"), AZUL
    for letra, ancho in zip("ABCD", (34, 12, 20, 55)):
        hoja.column_dimensions[letra].width = ancho


def desglose_detalle(hoja, cols, detalle):
    """Llena las columnas calculadas del Detalle con fórmulas y las agrupa bajo un botón +."""
    L = {c: get_column_letter(cols.index(c) + 1) for c in cols}
    # Filas del Detalle que pagan el mismo XML (en pago quebrado son varias; se suman sus importes)
    por_xml = {}
    for k, fila in enumerate(detalle.itertuples(index=False), start=2):
        if isinstance(fila.ArchivoXML, str) and fila.ArchivoXML:
            por_xml.setdefault(fila.ArchivoXML, []).append(k)
    for k, fila in enumerate(detalle.itertuples(index=False), start=2):
        if pd.isna(fila.SubTotal):
            continue                                   # sin importes (sin XML o 2a línea de pago quebrado)
        f = lambda c: f"{L[c]}{k}"
        txt = "+".join(f"{L['Importe']}{j}" for j in por_xml[fila.ArchivoXML])
        formulas = {
            # Variación de importe = (Base + IVA de la factura) − importe del TXT. Positivo: la factura es mayor
            "Variación importe": f"=ROUND({f('Base')}+{f('IVA')}-({txt}),2)",
            "Revisión importe": ETIQUETA_IMPORTE.format(v=f("Variación importe")),
            "Base": f"={f('SubTotal')}-{f('Descuento')}",
            "IVA calculado": f'=IF({f("CodigoImpuesto")}="F2",ROUND({f("Base")}*TasaIVA,2),0)',
            "Dif IVA": f"={f('IVA')}-{f('IVA calculado')}",
            "RetISR calculado": f'=ROUND({f("Base")}*IF({f("RegimenEmisor")}="626",'
                                f'TasaISR_RESICO,TasaISR_General),2)',
            "Dif RetISR": f"={f('RetISR')}-{f('RetISR calculado')}",
            "RetIVA calculado": f"=ROUND({f('IVA calculado')}*ProporcionRetIVA,2)",
            "Dif RetIVA": f"={f('RetIVA')}-{f('RetIVA calculado')}",
            "Total calculado": f"={f('Base')}+{f('IVA calculado')}-{f('RetISR calculado')}"
                               f"-{f('RetIVA calculado')}",
            "Dif Total": f"={f('Total')}-{f('Total calculado')}",
            # Variación = suma de las diferencias (sin signo) de IVA, Ret. ISR y Ret. IVA
            "Variación impuestos": f"=ROUND(ABS({f('Dif IVA')})+ABS({f('Dif RetISR')})"
                                   f"+ABS({f('Dif RetIVA')}),2)",
            "Revisión impuestos": ETIQUETA_VARIACION.format(v=f("Variación impuestos")),
        }
        for col, formula in formulas.items():
            hoja[f(col)] = formula
        for col in COLS_DESGLOSE + ["Variación impuestos", "Variación importe"]:
            hoja[f(col)].number_format = FMT_PESOS

    # Colores de la etiqueta: verde OK, amarillo Redondeo, rojo Revisar
    from openpyxl.formatting.rule import CellIsRule
    for columna in ("Revisión impuestos", "Revisión importe"):
        rango = f"{L[columna]}2:{L[columna]}{hoja.max_row}"
        for texto, color in (("OK", "C6EFCE"), ("Redondeo", "FFEB9C"), ("Revisar", "FFC7CE")):
            hoja.conditional_formatting.add(rango, CellIsRule(operator="equal", formula=[f'"{texto}"'],
                                                              fill=PatternFill("solid", fgColor=color)))

    # Encabezados del desglose en gris claro para distinguirlos
    for col in COLS_DESGLOSE:
        celda = hoja[f"{L[col]}1"]
        celda.fill, celda.font = GRIS, Font(bold=True)
    # Anchos fijos (el autoajuste mediría el texto de la fórmula)
    anchas = ("Revisión impuestos", "Variación impuestos", "Revisión importe", "Variación importe")
    for col in COLS_DESGLOSE + ["Ver cálculo"] + list(anchas):
        hoja.column_dimensions[L[col]].width = 18 if col in anchas else 14
    # Agrupar: quedan ocultas y aparece un botón + arriba de "Revisión impuestos"
    for col in COLS_DESGLOSE:
        hoja.column_dimensions[L[col]].outline_level = 1
        hoja.column_dimensions[L[col]].hidden = True
    hoja.sheet_format.outlineLevelCol = 1
    hoja.sheet_properties.outlinePr.summaryRight = False
    return L


def escribir_calculos(libro, L, detalle):
    """Hoja Cálculos: una ficha por línea, en vertical y con explicación en palabras.
    Desde el Detalle se llega con el vínculo 'Ver cálculo' y se regresa con '← Volver'."""
    from openpyxl.formatting.rule import FormulaRule
    hoja = libro.create_sheet("Cálculos")
    det = libro["Detalle"]
    rojo = PatternFill("solid", fgColor="FFC7CE")
    r, fichas = 1, {}                                   # fichas: línea del TXT -> fila donde empieza

    # Líneas del TXT que pagan el mismo XML (pago quebrado): se suman en la ficha
    por_xml = {}
    for k, fila in enumerate(detalle.itertuples(index=False), start=2):
        if isinstance(fila.ArchivoXML, str) and fila.ArchivoXML:
            por_xml.setdefault(fila.ArchivoXML, []).append((k, fila.Linea))
    lineas_txt = lambda fila: [f"Detalle!{L['Importe']}{k}" for k, _ in por_xml[fila.ArchivoXML]]

    def explica_txt(fila):
        grupo = por_xml[fila.ArchivoXML]
        if len(grupo) == 1:
            return "Base + IVA de la factura, comparado con el importe que viene en el TXT"
        nums = ", ".join(str(n) for _, n in grupo)
        return f"Pago quebrado: se suman las líneas {nums} del TXT que pagan esta(s) factura(s)"
    for k, fila in enumerate(detalle.itertuples(index=False), start=2):
        if pd.isna(fila.SubTotal):
            continue
        d = lambda c: f"Detalle!{L[c]}{k}"             # referencia a la celda del Detalle
        s = r                                           # primera fila de la ficha

        hoja[f"A{s}"] = f"Línea {fila.Linea} del TXT  –  {fila.Nombre}"
        hoja[f"A{s}"].font = Font(bold=True, color="FFFFFF", size=12)
        for c in "ABCDE":
            hoja[f"{c}{s}"].fill = AZUL
        ligar(hoja[f"F{s}"], "Detalle", f"A{k}", "← Volver al Detalle")
        hoja[f"A{s+1}"] = (f'="RFC: "&{d("RFC")}&"     Régimen: "&{d("RegimenEmisor")}&" "&{d("NombreRegimen")}'
                           f'&"     Código TXT: "&{d("CodigoImpuesto")}&"     Estatus: "&{d("Estatus")}')
        hoja[f"A{s+1}"].font = Font(italic=True)
        fichas[fila.Linea] = s

        encab = ["Concepto", "Según la factura (XML)", "Debería ser", "Diferencia", "Cómo se calcula"]
        for c, texto in zip("ABCDE", encab):
            hoja[f"{c}{s+2}"] = texto
            hoja[f"{c}{s+2}"].font, hoja[f"{c}{s+2}"].fill = Font(bold=True), GRIS

        b = s + 3                                       # fila del Subtotal
        renglones = [
            ("Subtotal", f"={d('SubTotal')}", None,
             "Importe de los honorarios antes de impuestos"),
            ("(–) Descuento", f"={d('Descuento')}", None,
             "Descuento que trae la factura (casi siempre cero)"),
            ("Base", f"=B{b}-B{b+1}", None,
             "Subtotal menos descuento. Sobre esta cantidad se calculan los impuestos"),
            ("(+) IVA", f"={d('IVA')}", f'=IF({d("CodigoImpuesto")}="F2",ROUND(B{b+2}*TasaIVA,2),0)',
             "16% de la base, solo si el TXT dice F2. Si dice E0 no debe llevar IVA"),
            ("(–) Retención ISR", f"={d('RetISR')}",
             f'=ROUND(B{b+2}*IF({d("RegimenEmisor")}="626",TasaISR_RESICO,TasaISR_General),2)',
             "10% de la base. Si el régimen es 626 (RESICO), 1.25%"),
            ("(–) Retención IVA", f"={d('RetIVA')}", f"=ROUND(C{b+3}*ProporcionRetIVA,2)",
             "Dos terceras partes del IVA calculado"),
            ("Total de la factura", f"={d('Total')}", f"=B{b+2}+C{b+3}-C{b+4}-C{b+5}",
             "Base + IVA − retenciones. Es lo que realmente se le deposita"),
            ("Importe contra el TXT", f"=B{b+2}+B{b+3}", "=" + "+".join(d_txt for d_txt in lineas_txt(fila)),
             explica_txt(fila)),
        ]
        for i, (concepto, xml, calc, explica) in enumerate(renglones):
            rr = b + i
            hoja[f"A{rr}"], hoja[f"B{rr}"], hoja[f"E{rr}"] = concepto, xml, explica
            if calc:
                hoja[f"C{rr}"] = calc
                hoja[f"D{rr}"] = f"=B{rr}-C{rr}"
            for c in "BCD":
                hoja[f"{c}{rr}"].number_format = FMT_PESOS
            if concepto in ("Base", "Total de la factura"):
                hoja[f"A{rr}"].font = Font(bold=True)
        hoja[f"C{b+7}"].font = Font(color="008000")    # verde = viene de otra hoja (TXT)

        var = b + len(renglones)                        # variación de impuestos de la ficha
        hoja[f"A{var}"] = "Variación de impuestos"
        hoja[f"B{var}"] = f"=ROUND(ABS(D{b+3})+ABS(D{b+4})+ABS(D{b+5}),2)"
        hoja[f"B{var}"].number_format = FMT_PESOS
        hoja[f"E{var}"] = ("Suma de las diferencias de IVA, Ret. ISR y Ret. IVA (sin importar el signo). "
                           "Hasta 1.00 se acepta como redondeo")
        for c in "AB":
            hoja[f"{c}{var}"].font = Font(bold=True)

        res = var + 1
        hoja[f"A{res}"] = (f'=IF(B{var}<=Tolerancia,"✔ Los impuestos de esta factura cuadran",'
                           f'IF(B{var}<=ToleranciaImpuestos,"✔ Diferencia de redondeo aceptada: "&TEXT(B{var},"0.00"),'
                           f'"✖ Revisar: la diferencia pasa la tolerancia de "&TEXT(ToleranciaImpuestos,"0.00")))')
        hoja[f"A{res}"].font = Font(bold=True)

        ligar(det[f"{L['Ver cálculo']}{k}"], "Cálculos", f"A{s}", "Ver cálculo")
        r = res + 3

    # Las demás líneas de un pago quebrado también llevan a la ficha de su factura
    for grupo in por_xml.values():
        con_ficha = [n for _, n in grupo if n in fichas]
        for k, n in grupo:
            if con_ficha and n not in fichas:
                fichas[n] = fichas[con_ficha[0]]
                ligar(det[f"{L['Ver cálculo']}{k}"], "Cálculos", f"A{fichas[n]}", "Ver cálculo")

    # Diferencias: amarillo si es redondeo aceptado, rojo si pasa la tolerancia de impuestos
    hoja.conditional_formatting.add(
        f"D1:D{max(r, 2)}",
        FormulaRule(formula=["AND(ISNUMBER(D1),ABS(D1)>ToleranciaImpuestos)"], fill=rojo, stopIfTrue=True))
    hoja.conditional_formatting.add(
        f"D1:D{max(r, 2)}",
        FormulaRule(formula=["AND(ISNUMBER(D1),ABS(D1)>Tolerancia)"],
                    fill=PatternFill("solid", fgColor="FFEB9C")))
    hoja.conditional_formatting.add(
        f"A1:A{max(r, 2)}", FormulaRule(formula=['LEFT(A1,1)="✖"'], font=Font(bold=True, color="C00000")))
    for letra, ancho in zip("ABCDEF", (26, 22, 16, 16, 72, 22)):
        hoja.column_dimensions[letra].width = ancho
    hoja.sheet_view.showGridLines = False
    return fichas


def escribir_por_proveedor(libro, L, detalle, fichas):
    """Hoja Por proveedor: un renglón por proveedor con totales; con el botón + se abren sus líneas."""
    hoja = libro.create_sheet("Por proveedor")
    encab = ["Proveedor / Línea", "Nombre", "RFC", "Estatus", "Importe TXT", "Subtotal", "IVA",
             "Ret. ISR", "Ret. IVA", "Total", "Revisión impuestos", "Variación impuestos", "Variación importe", "Ver cálculo"]
    numericas = {"Importe TXT": "Importe", "Subtotal": "SubTotal", "IVA": "IVA",
                 "Ret. ISR": "RetISR", "Ret. IVA": "RetIVA", "Total": "Total"}
    hoja.append(encab)
    hoja.append(["TOTAL GENERAL"])                   # fórmulas abajo, cuando ya se sabe el rango

    det = detalle.reset_index(drop=True)
    det["FilaExcel"] = det.index + 2
    for id_prov, grupo in det.groupby("IDProveedor", sort=True):
        resumen = hoja.max_row + 1
        hoja.append([id_prov, grupo["Nombre"].iloc[0], grupo["RFC"].iloc[0],
                     f"{len(grupo)} línea(s)"])
        primera, ultima = resumen + 1, resumen + len(grupo)
        for fila in grupo.itertuples(index=False):
            k = fila.FilaExcel
            hoja.append([f"    Línea {fila.Linea}", fila.Folio, None, fila.Estatus])
            r = hoja.max_row
            for i, nombre in enumerate(encab[4:10], start=5):
                celda = hoja.cell(r, i, f"=Detalle!{L[numericas[nombre]]}{k}")
                celda.number_format = FMT_PESOS
            hoja.cell(r, 11, f"=Detalle!{L['Revisión impuestos']}{k}&\"\"")
            v = f"Detalle!{L['Variación impuestos']}{k}"
            hoja.cell(r, 12, f'=IF({v}="","",{v})').number_format = FMT_PESOS
            v = f"Detalle!{L['Variación importe']}{k}"
            hoja.cell(r, 13, f'=IF({v}="","",{v})').number_format = FMT_PESOS
            if fila.Linea in fichas:
                ligar(hoja.cell(r, 14), "Cálculos", f"A{fichas[fila.Linea]}", "Ver cálculo")
            hoja.row_dimensions[r].outline_level = 1
            hoja.row_dimensions[r].hidden = True
        for i in list(range(5, 11)) + [12, 13]:
            letra = get_column_letter(i)
            celda = hoja.cell(resumen, i, f"=SUBTOTAL(9,{letra}{primera}:{letra}{ultima})")
            celda.number_format = FMT_PESOS
        hoja.cell(resumen, 11, f'=IF(COUNTIF(K{primera}:K{ultima},"Revisar")>0,"Revisar",'
                                f'IF(COUNTIF(K{primera}:K{ultima},"Redondeo")>0,"Redondeo",'
                                f'IF(COUNTIF(K{primera}:K{ultima},"OK")>0,"OK","")))')
        for c in hoja[resumen]:
            c.font = Font(bold=True)
            c.fill = GRIS

    fin = hoja.max_row
    for i in list(range(5, 11)) + [12, 13]:
        letra = get_column_letter(i)
        celda = hoja.cell(2, i, f"=SUBTOTAL(9,{letra}3:{letra}{fin})")   # SUBTOTAL ignora los subtotales
        celda.number_format = FMT_PESOS
    for c in hoja[2]:
        c.font = Font(bold=True)
    for c in hoja[1]:
        c.font, c.fill = Font(bold=True, color="FFFFFF"), AZUL

    from openpyxl.formatting.rule import CellIsRule
    for texto, color in (("OK", "C6EFCE"), ("Redondeo", "FFEB9C"), ("Revisar", "FFC7CE")):
        hoja.conditional_formatting.add(f"K3:K{fin}", CellIsRule(operator="equal", formula=[f'"{texto}"'],
                                        fill=PatternFill("solid", fgColor=color)))
    hoja.sheet_format.outlineLevelRow = 1
    hoja.sheet_properties.outlinePr.summaryBelow = False   # el botón + queda en el renglón del proveedor
    hoja.freeze_panes = "A3"
    for letra, ancho in zip("ABCDEFGHIJKLMN", (20, 30, 16, 24, 14, 14, 13, 13, 13, 14, 18, 18, 18, 13)):
        hoja.column_dimensions[letra].width = ancho


def escribir_instrucciones(libro):
    hoja = libro.create_sheet("Cómo usar", 1)          # justo después del Resumen
    textos = [
        ("CÓMO LEER ESTE ARCHIVO", True),
        ("", False),
        ("1. Resumen: cuántas líneas hay en cada estatus y por cuánto dinero.", False),
        ("2. Por proveedor: un renglón por proveedor con sus totales.", False),
        ("     • Da clic en el botón  +  a la izquierda para ver sus líneas; en  –  para cerrarlas.", False),
        ("     • Los botones  1  y  2  de la esquina superior izquierda cierran o abren todo.", False),
        ("3. Detalle: una fila por línea del TXT.", False),
        ("     • Arriba de la columna 'Revisión impuestos' hay un botón  +  que muestra el desglose:", False),
        ("       subtotal, IVA, retenciones, y lo que 'debería ser' según las tasas.", False),
        ("     • 'Variación importe' = factura (subtotal − descuento + IVA) menos importe del TXT, en pesos.", False),
        ("       Positivo: la factura es mayor que el TXT. Negativo: el TXT es mayor.", False),
        ("     • 'Revisión importe': OK (hasta 0.05) · Redondeo (hasta 1.00, se acepta).", False),
        ("     • 'Revisión impuestos': OK (hasta 0.05) · Redondeo (hasta 1.00, se acepta) · "
         "Revisar (más de 1.00).", False),
        ("     • El vínculo 'Ver cálculo' lleva a la ficha de esa línea.", False),
        ("4. Cálculos: la ficha de cada línea, paso a paso y explicada en palabras.", False),
        ("5. Tasas: los porcentajes que usan las fórmulas (IVA 16%, ISR 10%, RESICO 1.25%...).", False),
        ("", False),
        ("QUÉ SIGNIFICA CADA ESTATUS", True),
        ("", False),
        ("• Cruzado: pasó TODAS las validaciones. No hay nada que revisar.", False),
        ("     Se validó: importe contra TXT, timbre, versión 4.0, datos del receptor (RFC, régimen, uso CFDI),", False),
        ("     periodo, tipo I, método y forma de pago (PUE con 03 o 99 / PPD con 99), moneda, total del XML, impuestos,", False),
        ("     códigos de retención en SAP (si se pidió) y que el UUID no se haya", False),
        ("     conciliado en otro mes ni venga repetido.", False),
        ("• Cruzado (pago quebrado): igual que Cruzado, pero varias líneas del TXT y/o varias facturas", False),
        ("     del mismo proveedor se cruzan como grupo (p. ej. 500 + 300 + 200 del TXT contra una factura", False),
        ("     de 1,000, o 600 + 400 contra dos facturas de 500).", False),
        ("• Cruzado con observaciones: cruzó, pero hay algo que revisar (ver Observaciones). Incluye las", False),
        ("     líneas en las que alguna validación NO se pudo hacer (SAP pedido pero cerrado, historial ilegible...).", False),
        ("• Detenida: no se debe pagar hasta corregir (impuestos, SAP, pago doble de un UUID, proveedor", False),
        ("     bloqueado...). Si el mismo XML viene copiado dos veces en la carpeta pero solo una línea lo", False),
        ("     paga, sale 'Cruzado con observaciones': Verificar con Responsables de Envío de XML.", False),
        ("• Proveedor extranjero: el catálogo trae un ID fiscal de otro país; no hay CFDI que cruzar.", False),
        ("• RFC inválido en catálogo: el RFC parece mexicano pero está mal escrito; corregir el catálogo.", False),
        ("• Rechazar: la factura no es válida (sin timbre, receptor, periodo, forma o método de pago", False),
        ("     incorrecto...). Hay que pedir otra.", False),
        ("", False),
        ("Colores: verde = OK · amarillo = redondeo aceptado · rojo = revisar. "
         "En la hoja Tasas, las celdas amarillas (tasas y tolerancias) se pueden cambiar.", False),
    ]
    for texto, negrita in textos:
        hoja.append([texto])
        hoja.cell(hoja.max_row, 1).font = Font(bold=negrita, size=14 if negrita else 12)
    hoja.column_dimensions["A"].width = 110
    hoja.sheet_view.showGridLines = False


def generar_reporte(rutas, anio, mes, resultado, sobrantes, errores_xml, n_xml, estados):
    detalle = pd.DataFrame(resultado)
    columnas = ["Linea", "Estatus", "Fecha", "IDProveedor", "Nombre", "RFC", "Folio",
                "CodigoImpuesto", "Importe", "Ver cálculo", "Revisión importe", "Variación importe",
                "Revisión impuestos", "Variación impuestos"] + COLS_DESGLOSE + [
                "RegimenEmisor", "NombreRegimen", "CodigosSAP", "RegimenSAP",
                "DiscrepanciaSAP", "FechaXML", "MetodoPago", "FormaPago", "UUID", "ArchivoXML",
                "CentroCosto", "Cuenta", "MotivoRechazo", "Observaciones"]
    detalle = detalle.reindex(columns=columnas)
    detalle["DiscrepanciaSAP"] = detalle["DiscrepanciaSAP"].map(lambda v: "Sí" if v is True else "No")
    detalle["_orden"] = detalle["Estatus"].map({e: i for i, e in enumerate(ORDEN_ESTATUS)})
    detalle = detalle.sort_values(["_orden", "IDProveedor"]).drop(columns="_orden")
    detalle = a_float(detalle)

    # Resumen para la pantalla (en Excel se escribe con fórmulas)
    resumen = (detalle.groupby("Estatus", sort=False)
               .agg(Lineas=("Linea", "count"), Importe=("Importe", "sum"))
               .reindex([e for e in ORDEN_ESTATUS if e in set(detalle["Estatus"])])
               .reset_index())
    total = pd.DataFrame([{"Estatus": "TOTAL TXT", "Lineas": resumen["Lineas"].sum(),
                           "Importe": resumen["Importe"].sum()}])
    extra = pd.DataFrame([
        {"Estatus": "Líneas con validaciones no realizadas",
         "Lineas": int(detalle["Observaciones"].fillna("").str.contains("Validaciones no realizadas").sum())},
        {"Estatus": "Líneas con discrepancia SAP", "Lineas": int((detalle["DiscrepanciaSAP"] == "Sí").sum())},
        {"Estatus": "XML leídos", "Lineas": n_xml},
        {"Estatus": "XML sin línea en TXT", "Lineas": len(sobrantes)},
        {"Estatus": "XML que no se pudieron leer", "Lineas": len(errores_xml)},
    ])
    resumen = pd.concat([resumen, total, extra], ignore_index=True)

    info = pd.DataFrame([
        {"Dato": "Periodo", "Valor": f"{MESES[mes - 1]} {anio}"},
        {"Dato": "Fecha de ejecución", "Valor": datetime.now().strftime("%d/%m/%Y %H:%M")},
        {"Dato": "Archivo TXT", "Valor": rutas["txt"].name},
        {"Dato": "SHA-256 del TXT", "Valor": huella(rutas["txt"])},
        {"Dato": "SHA-256 del catálogo", "Valor": huella(rutas["catalogo"])},
        {"Dato": "Control: líneas del TXT", "Valor": str(len(detalle))},
        {"Dato": "Control: suma de importes del TXT",
         "Valor": f"{sum((r['Importe'] for r in resultado), Decimal('0')):,.2f}"},
        {"Dato": "Control: XML leídos / usados / sin TXT / con error",
         "Valor": f"{n_xml} / {n_xml - len(errores_xml) - len(sobrantes)} / {len(sobrantes)} / {len(errores_xml)}"},
        {"Dato": "Catálogo", "Valor": f"{rutas['catalogo'].name} (modificado "
                 f"{datetime.fromtimestamp(rutas['catalogo'].stat().st_mtime):%d/%m/%Y})"},
        {"Dato": "RFC receptor esperado", "Valor": RFC_SOCIEDAD},
        {"Dato": "Nombre receptor esperado", "Valor": NOMBRE_SOCIEDAD or "No configurado (no se valida)"},
        {"Dato": "CP receptor esperado", "Valor": CP_SOCIEDAD or "No configurado (no se valida)"},
        {"Dato": "Régimen receptor / Uso CFDI", "Valor": f"{REGIMEN_RECEPTOR} / {', '.join(sorted(USOS_CFDI_OK))}"},
        {"Dato": "Métodos de pago aceptados", "Valor": ", ".join(sorted(METODOS_PAGO_OK))},
        {"Dato": "Formas de pago aceptadas",
         "Valor": ", ".join(f"{k} {FORMAS_PAGO.get(k, '')}".strip() for k in sorted(FORMAS_PAGO_OK))},
    ] + [{"Dato": nombre, "Valor": texto} for nombre, texto in estados.items()] + [
        {"Dato": "Tolerancia de importe (cruce TXT vs XML)", "Valor": str(TOLERANCIA)},
        {"Dato": "Tolerancia de impuestos (redondeo)", "Valor": str(TOLERANCIA_IMPUESTOS)},
        {"Dato": "Variación etiquetada como OK (hasta)", "Valor": str(UMBRAL_OK)},
    ])

    nombre = f"Conciliacion_{anio}-{mes:02d}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    destino = rutas["salida"] / nombre
    with pd.ExcelWriter(destino, engine="openpyxl") as w:
        w.book.create_sheet("Resumen")
        detalle.to_excel(w, sheet_name="Detalle", index=False)
        a_float(pd.DataFrame(sobrantes)).to_excel(w, sheet_name="XML sin TXT", index=False)
        pd.DataFrame(errores_xml, columns=["Archivo", "Error"]).to_excel(
            w, sheet_name="Errores XML", index=False)
        info.to_excel(w, sheet_name="Info", index=False)
        for hoja in w.book.worksheets:
            if hoja.title != "Resumen":
                formatear_hoja(hoja, colorear_estatus=(hoja.title == "Detalle"))
        hoja_det = w.book["Detalle"]
        hoja_det.auto_filter.ref = hoja_det.dimensions
        escribir_resumen(w.book["Resumen"], list(detalle.columns), len(detalle), n_xml)
        escribir_tasas(w.book)
        L = desglose_detalle(hoja_det, list(detalle.columns), detalle)
        fichas = escribir_calculos(w.book, L, detalle)
        escribir_por_proveedor(w.book, L, detalle, fichas)
        escribir_instrucciones(w.book)
        orden = ["Resumen", "Cómo usar", "Por proveedor", "Detalle", "Cálculos",
                 "XML sin TXT", "Errores XML", "Tasas", "Info"]
        w.book._sheets.sort(key=lambda h: orden.index(h.title) if h.title in orden else 99)
        w.book.active = 0
    return destino, resumen


# ------------------------------------------------------------------
# 5. SAP (códigos de retención por proveedor, transacción FK03)
# ------------------------------------------------------------------
NOMBRE_TABLA_LFBW = "tblSAPLCVI_FS_UI_VENDOR_CCTCTRL_LFBW"   # tabla de retenciones en BP
CASILLA_RETENCION = "chkRF02K-D0610"


def buscar_controles(nodo, condicion, encontrados=None, nivel=0):
    """Recorre la pantalla de SAP y devuelve los controles que cumplen la condición."""
    if encontrados is None:
        encontrados = []
    try:
        hijos = nodo.Children
    except Exception:
        return encontrados
    for i in range(hijos.Count):
        h = hijos.ElementAt(i)
        if condicion(h):
            encontrados.append(h)
        if nivel < 25:
            buscar_controles(h, condicion, encontrados, nivel + 1)
    return encontrados


def conectar_sap():
    import win32com.client
    return win32com.client.GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)


def cerrar_ventanas(session):
    """Cierra ventanas emergentes (wnd[1], wnd[2]...) que bloquean la pantalla principal."""
    for _ in range(5):
        try:
            session.findById("wnd[1]").close()
        except Exception:
            return


def enter(session):
    """Presiona Enter y confirma avisos amarillos (tipo W), que SAP detiene hasta otro Enter."""
    session.findById("wnd[0]").sendVKey(0)
    for _ in range(3):
        if session.findById("wnd[0]/sbar").MessageType != "W":
            break
        session.findById("wnd[0]").sendVKey(0)


def codigos_retencion(session, id_proveedor):
    paso = "abrir FK03"
    try:
        cerrar_ventanas(session)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nFK03"
        session.findById("wnd[0]").sendVKey(0)
        cerrar_ventanas(session)

        paso = "marcar solo la casilla de retenciones"
        usr = session.findById("wnd[0]/usr")
        for casilla in buscar_controles(usr, lambda c: c.Type == "GuiCheckBox"):
            casilla.Selected = casilla.Id.endswith(CASILLA_RETENCION)

        paso = "llenar proveedor y sociedad"
        session.findById("wnd[0]/usr/ctxtRF02K-LIFNR").text = id_proveedor
        session.findById("wnd[0]/usr/ctxtRF02K-BUKRS").text = SOCIEDAD_SAP
        enter(session)

        barra = session.findById("wnd[0]/sbar")
        if barra.MessageType in ("E", "A"):
            return [], barra.Text
        cerrar_ventanas(session)

        paso = "leer tabla de retenciones"
        tablas = buscar_controles(session.findById("wnd[0]/usr"),
                                  lambda c: c.Id.endswith(NOMBRE_TABLA_LFBW))
        if not tablas:
            return [], f"No apareció la tabla de retenciones ({session.findById('wnd[0]').Text})"
        tabla = tablas[0]
        codigos = []
        for fila in range(tabla.VisibleRowCount):
            texto = tabla.GetCell(fila, 2).Text.strip()
            if not texto:
                break
            codigos.append(texto)
        else:
            if tabla.VisibleRowCount and tabla.RowCount > tabla.VisibleRowCount:
                # Hay renglones fuera de la vista: podría faltar IH, IP o el de ISR. Mejor no validar
                # que validar con datos incompletos.
                return [], (f"La tabla de retenciones tiene {tabla.RowCount} renglones y solo se ven "
                            f"{tabla.VisibleRowCount}: revisar en FK03 a mano")
        return codigos, ""
    except Exception as e:
        detalle = e.args[2][2] if len(e.args) > 2 and e.args[2] else str(e)
        return [], f"Falló en el paso '{paso}': {detalle}"


def consultar_sap(ids, carpeta_salida, anio, mes, reusar=True):
    """Consulta cada proveedor una sola vez y guarda el resultado para no repetir.
    Con reusar=False se ignora lo guardado (útil si se corrigió algo en SAP)."""
    cache_ruta = carpeta_salida / f"CodigosSAP_{anio}-{mes:02d}.xlsx"
    resultado = {}
    if reusar and cache_ruta.exists():
        previo = pd.read_excel(cache_ruta, dtype=str).fillna("")
        for _, r in previo.iterrows():
            codigos = [c for c in r["Codigos"].split(",") if c]
            if not r["Error"] and r["IDProveedor"] in ids:
                resultado[r["IDProveedor"]] = (codigos, "")

    faltan = [i for i in ids if i not in resultado]
    no_consultados = 0
    if faltan:
        try:
            session = conectar_sap()
        except Exception:
            print("  SAP no está abierto o el scripting no está habilitado. Se omite la validación SAP.")
            if not resultado:
                return None, "No se pudo conectar a SAP"
            return resultado, (f"No se pudo conectar a SAP: se usaron {len(resultado)} consultas guardadas; "
                               f"{len(faltan)} proveedores sin validar")
        fallas_seguidas = 0
        for k, id_prov in enumerate(faltan, start=1):
            print(f"  SAP {k}/{len(faltan)}: proveedor {id_prov}   ", end="\r")
            resultado[id_prov] = codigos_retencion(session, id_prov)
            fallas_seguidas = fallas_seguidas + 1 if resultado[id_prov][1] else 0
            if fallas_seguidas == 3:
                print(f"\n  3 proveedores seguidos fallaron. Último error: {resultado[id_prov][1]}")
                print("  Se detiene la consulta a SAP para no seguir fallando.")
                no_consultados = len(faltan) - k
                break
        print()

    try:
        pd.DataFrame([{"IDProveedor": i, "Codigos": ",".join(c), "Error": e,
                       "Consulta": datetime.now().strftime("%d/%m/%Y %H:%M")}
                      for i, (c, e) in resultado.items()]).to_excel(cache_ruta, index=False)
    except PermissionError:
        print(f"  No se pudo guardar {cache_ruta.name} (¿está abierto?).")
    texto = f"{len(resultado)} proveedores consultados (guardado en {cache_ruta.name})"
    if no_consultados:
        texto += f"; {no_consultados} sin consultar por fallas"
    return resultado, texto


# ------------------------------------------------------------------
# PRINCIPAL
# ------------------------------------------------------------------
def si(pregunta, defecto):
    respuesta = input(f"{pregunta} (s/n) [{defecto}]: ").strip().lower() or defecto
    return respuesta in ("s", "si", "sí", "y", "yes")


def main():
    anio, mes = pedir_periodo()
    faltan_config = [n for n, v in (("NOMBRE_SOCIEDAD", NOMBRE_SOCIEDAD), ("CP_SOCIEDAD", CP_SOCIEDAD)) if not v]
    if faltan_config:
        print(f"\nAviso: {', '.join(faltan_config)} vacío(s) en la configuración; esa validación no se hace.")

    print("\nRevisando carpetas...")
    rutas = obtener_rutas(anio, mes)

    print(f"Leyendo catálogo ({rutas['catalogo'].name})...")
    catalogo = leer_catalogo(rutas["catalogo"])
    tipos = catalogo["TipoRFC"].value_counts()
    resumen_cat = (f"{len(catalogo):,} proveedores: {tipos.get('MX', 0):,} con RFC mexicano, "
                   f"{tipos.get('Extranjero', 0):,} extranjeros, {tipos.get('Inválido', 0):,} con RFC inválido, "
                   f"{tipos.get('Vacío', 0):,} sin RFC")
    print(f"  {resumen_cat}")

    print(f"Leyendo TXT ({rutas['txt'].name})...")
    lineas = leer_txt(rutas["txt"])
    print(f"  {len(lineas):,} líneas")

    print(f"Leyendo {len(rutas['xmls']):,} XML...")
    facturas, errores_xml = leer_xmls(rutas["xmls"], rutas["carpeta_xml"])
    print(f"  {len(facturas):,} leídos, {len(errores_xml)} con error")

    estados = {"Catálogo de proveedores": resumen_cat}

    historial, estados["Historial de UUID"] = cargar_historial()
    print(f"Historial de UUID: {estados['Historial de UUID']}")

    sap, estados["Validación SAP"] = None, "No se consultó"
    if si("¿Validar contra SAP? SAP debe estar abierto", "n"):
        en_catalogo = set(catalogo["IDProveedor"])
        ids = sorted({l["IDProveedor"] for l in lineas if l["IDProveedor"] in en_catalogo})
        cache = rutas["salida"] / f"CodigosSAP_{anio}-{mes:02d}.xlsx"
        reusar = not cache.exists() or si("Hay consultas SAP guardadas de este mes. ¿Reusarlas? "
                                          "(di 'n' si corregiste algo en SAP)", "s")
        print(f"Consultando {len(ids)} proveedores en SAP (FAVOR DE NO TOCAR EL MOUSE NI EL TECLADO)...")
        sap, estados["Validación SAP"] = consultar_sap(ids, rutas["salida"], anio, mes, reusar)

    print("Cruzando...")
    resultado, sobrantes = conciliar(lineas, catalogo, facturas, anio, mes, sap, historial)

    destino, resumen = generar_reporte(rutas, anio, mes, resultado, sobrantes,
                                       errores_xml, len(rutas["xmls"]), estados)
    print("\n" + resumen.to_string(index=False))
    print(f"\nReporte guardado en:\n  {destino}")

    if historial is not None:
        print(guardar_historial(resultado, anio, mes))
    else:
        print("El historial de UUID no se actualizó porque no se pudo leer.")


if __name__ == "__main__":
    try:
        main()
    except ErrorUsuario as e:
        print(f"\n*** NO SE PUDO COMPLETAR ***\n{e}")
    except Exception as e:
        print(f"\n*** ERROR INESPERADO ***\n{type(e).__name__}: {e}")
        print("Toma captura de este mensaje y compártela con quien mantiene el programa.")
    input("\nPresiona Enter para cerrar...")
