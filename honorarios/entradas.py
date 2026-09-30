"""Ubicación de carpetas del periodo y lectura de las tres entradas: catálogo, TXT y XML (CFDI)."""

import re
import xml.etree.ElementTree as ET
from datetime import datetime
from decimal import Decimal

from openpyxl import load_workbook
import pandas as pd

from .config import ARCHIVO_CATALOGO, COLUMNAS_TXT, CONFIG, MESES, NOMBRE_TABLA_CATALOGO, RAIZ
from .utilidades import (
    ErrorUsuario,
    clasificar_rfc,
    importe_txt,
    importe_xml,
    leer_texto,
    normalizar_activo,
    normalizar_id,
    normalizar_nombre,
)


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
    primera = True
    for n, linea in enumerate(lineas, start=1):
        if not linea.strip():
            continue
        campos = [c.strip() for c in linea.split("\t")]
        if primera:
            primera = False
            if not re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", campos[0]):
                continue                               # es un encabezado (aunque haya renglones vacíos antes)
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
