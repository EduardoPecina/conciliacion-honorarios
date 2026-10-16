"""TXT de pagos para SAP, generado desde el Excel de las líneas aprobadas (página Generar TXT).

El analista arma un Excel solo con las líneas listas para pagar, con las primeras 10 columnas del TXT en su orden
(A a J): Fecha, IDProveedor, Folio, Monto, X, Tax, 1, Centro de Costo, Cuenta Contable, PM. La hoja "TXT para SAP" del
Excel de la conciliación ya trae esas columnas: basta con dejar ahí las líneas que se van a pagar. La Referencia (la
columna 11 del TXT) no se captura: la arma el portal con el folio nuevo y la quincena de la fecha.

El portal NO edita nada: lee el Excel, revisa cada celda y, si algo no sirve, dice en qué fila y columna para que se
corrija en el Excel. Solo pone el Folio único y la Referencia, con "Formatear folios" (ver folios.py).

El TXT sale con el formato del TXT original: 11 columnas separadas por tabulador, sin encabezado, saltos de línea de
Windows (CRLF), codificación ANSI. Lo que Excel cambia al abrir un TXT se regresa a como va en el TXT: la fecha a
DD.MM.AAAA, el monto con 2 decimales y sin comas (9096 -> 9096.00), la columna 1 con sus ceros (1 -> 001)."""

import re
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

from . import folios
from .config import CODIGOS_IMPUESTO_TXT
from .entradas import es_encabezado_txt
from .utilidades import ErrorUsuario, importe_txt, normalizar_id

# Las 11 columnas del TXT, en su orden
COLUMNAS = ["Fecha", "IDProveedor", "Folio", "Monto", "X", "Tax", "1", "Centro de Costo", "Cuenta Contable",
            "PM", "Referencia"]
COLUMNAS_EXCEL = COLUMNAS[:-1]                 # las que se capturan en el Excel (A a J); la Referencia la arma el portal
LETRAS = dict(zip(COLUMNAS_EXCEL, "ABCDEFGHIJ"))
# Columnas que en el TXT de SAP llevan siempre el mismo valor
FIJOS = {"X": "X", "1": "001", "PM": "T"}
# Referencia = folio nuevo + separador + quincena: "26185113K7Q3MXP4_Q18-2026" (en el TXT original venía
# "A1893      Q18-2026": el folio, espacios y la quincena)
SEPARADOR_REFERENCIA = "_"
HOJA_PARA_TXT = "TXT para SAP"                 # hoja del Excel de la conciliación con las columnas del TXT
SALTO = "\r\n"
CODIFICACION = "cp1252"                        # ANSI de Windows
PROHIBIDOS = re.compile(r"[\t\r\n]")


def lineas_del_txt(datos):
    """bytes del TXT original -> {número de línea del archivo: [11 valores tal cual]} (sin encabezado ni renglones
    vacíos). Los números de línea son los mismos que usa la conciliación."""
    for codificacion in ("utf-8-sig", "latin-1"):
        try:
            texto = datos.decode(codificacion)
            break
        except UnicodeDecodeError:
            continue
    lineas, primera = {}, True
    for n, linea in enumerate(texto.splitlines(), start=1):
        if not linea.strip():
            continue
        campos = linea.split("\t")
        if primera:
            primera = False
            if es_encabezado_txt(campos):
                continue
        lineas[n] = (campos + [""] * len(COLUMNAS))[:len(COLUMNAS)]
    return lineas


# ------------------------------------------------------------------ Excel de las líneas aprobadas
def _texto(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))                         # 145362.0 -> 145362
    if isinstance(v, (datetime, date)):
        return v.strftime("%d.%m.%Y")
    return str(v).strip()


def _es_numero(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _fecha(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%d.%m.%Y"), None
    if _es_numero(v) and 20000 < v < 80000:        # fecha guardada como número de serie de Excel
        return from_excel(v).strftime("%d.%m.%Y"), None
    texto = _texto(v)
    for formato in ("%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(texto, formato).strftime("%d.%m.%Y"), None
        except ValueError:
            continue
    return texto, f"La fecha '{texto}' no es válida (formato DD.MM.AAAA, por ejemplo 29.09.2026)"


def _monto(v):
    if _es_numero(v):
        try:
            valor = Decimal(str(v))
        except InvalidOperation:
            valor = None
        if valor is not None and abs(valor * 100 - (valor * 100).to_integral_value()) > Decimal("0.000001"):
            return _texto(v), f"El monto {v} tiene más de 2 decimales"
    else:
        valor = importe_txt(_texto(v)) if _texto(v) else None
    if valor is None:
        return _texto(v), f"El monto '{_texto(v)}' no es válido (por ejemplo 9096.00)"
    if valor <= 0:
        return _texto(v), "El monto debe ser mayor que cero"
    return f"{valor.quantize(Decimal('0.01')):.2f}", None


def _valor(columna, v):
    """Celda del Excel -> (texto como va en el TXT, problema o None)."""
    if columna == "Fecha":
        return _fecha(v)
    if columna == "Monto":
        return _monto(v)
    if columna == "1":
        texto = _texto(v)
        if texto.isdigit() and int(texto) == 1:
            return FIJOS["1"], None
        return texto, f"Debe ser 1 (siempre lleva {FIJOS['1']})"
    if columna in ("X", "PM"):
        texto = _texto(v).upper()
        return texto, None if texto == FIJOS[columna] else f"Debe ser {FIJOS[columna]} (siempre lleva {FIJOS[columna]})"
    if columna == "Tax":
        texto = _texto(v).upper()
        return texto, None if texto in CODIGOS_IMPUESTO_TXT else f"'{texto}' no es válido: debe ser E0 (sin IVA) o F2 (con IVA)"
    # Una cuenta como 50E03158 Excel la puede leer como número en notación científica: ya no es la cuenta (las
    # cuentas de SAP miden hasta 10 caracteres)
    if columna == "Cuenta Contable" and _es_numero(v) and (not float(v).is_integer() or abs(v) >= 1e10):
        return _texto(v), ("Excel convirtió la cuenta en número (por ejemplo 5E+31): pon la columna como Texto "
                           "y vuelve a escribirla")
    texto = _texto(v)
    return texto, None if texto else f"Falta {'el' if columna in ('IDProveedor', 'Folio') else 'la'} {columna}"


def leer_excel(datos):
    """bytes del Excel -> (nombre de la hoja leída, líneas, problemas).
    líneas: [{"Fila": fila del Excel, "Fecha": ..., ..., "PM": ..., "Referencia": ""}] con los valores como van en el
    TXT (la Referencia se arma al formatear los folios).
    problemas: [(fila, columna, texto)]; si hay alguno, el TXT no se genera. Se lee la hoja "TXT para SAP" si el
    libro la tiene (el Excel de la conciliación) y si no, la hoja que estaba abierta al guardar. Solo las columnas A
    a J (lo que haya a la derecha no se lee); un encabezado en la primera fila y las filas vacías se saltan."""
    try:
        libro = load_workbook(BytesIO(datos), data_only=True)
    except Exception:                                          # noqa: BLE001 - cualquier archivo que no es Excel
        raise ErrorUsuario("No se pudo abrir el archivo como Excel. Guárdalo como \"Libro de Excel (.xlsx)\" y "
                           "vuelve a subirlo.")
    hoja = libro[HOJA_PARA_TXT] if HOJA_PARA_TXT in libro.sheetnames else libro.active
    lineas, problemas, primera = [], [], True
    for celdas in hoja.iter_rows(min_row=1, max_row=hoja.max_row, min_col=1, max_col=len(COLUMNAS_EXCEL)):
        valores = [c.value for c in celdas]
        if all(v is None or str(v).strip() == "" for v in valores):
            continue
        fila = celdas[0].row
        if primera:
            primera = False
            if isinstance(valores[0], str) and es_encabezado_txt([valores[0]]):
                continue
        linea = {"Fila": fila}
        for columna, v in zip(COLUMNAS_EXCEL, valores):
            linea[columna], problema = _valor(columna, v)
            if problema is None and PROHIBIDOS.search(linea[columna]):
                problema = "Tiene un tabulador o un salto de línea dentro de la celda"
            if problema is None:
                try:
                    linea[columna].encode(CODIFICACION)
                except UnicodeEncodeError:
                    problema = "Tiene un carácter que el TXT no admite (¿símbolo raro?)"
            if problema:
                problemas.append((fila, columna, problema))
        lineas.append({**linea, "Referencia": ""})
    if not lineas:
        raise ErrorUsuario(f"La hoja «{hoja.title}» no tiene líneas. Las columnas deben ir de la A a la J, en el "
                           "orden del TXT: " + ", ".join(COLUMNAS_EXCEL) + ".")
    return hoja.title, lineas, problemas


def misma_factura_repetida(lineas):
    """Filas con el mismo IDProveedor, folio, monto y fecha: [[fila, fila, ...]]. Puede ser un pago quebrado en partes
    iguales, o una línea copiada dos veces (se pagaría doble): lo decide el analista."""
    grupos = defaultdict(list)
    for linea in lineas:
        grupos[folios.clave_de(linea)].append(linea["Fila"])
    return [filas for filas in grupos.values() if len(filas) > 1]


def referencia(folio, fecha):
    """Referencia del TXT: el folio nuevo y la quincena de la fecha, unidos con guion bajo y sin espacios:
    "26185113K7Q3MXP4_Q18-2026" (como la del TXT original, "A1893      Q18-2026", con el folio nuevo)."""
    return f"{folio}{SEPARADOR_REFERENCIA}{folios.texto_quincena(fecha)}"


def quincenas(lineas):
    """{"Q18-2026": líneas} según la fecha de cada línea."""
    return Counter(folios.texto_quincena(l["Fecha"]) for l in lineas)


def formatear(lineas, registrados):
    """Pone a cada línea su folio único (folios.py) y su Referencia (folio nuevo + quincena).
    -> (líneas formateadas, problemas). Cada línea formateada guarda también FolioOriginal, Clave y Previo (el
    registro si ese folio ya se había generado antes)."""
    salida = []
    for linea, (folio, clave, previo) in zip(lineas, folios.formatear_folios(lineas, registrados)):
        salida.append({**linea, "Folio": folio, "Referencia": referencia(folio, linea["Fecha"]),
                       "FolioOriginal": linea["Folio"], "Clave": clave, "Previo": previo})
    # Un folio ya formateado se conserva: si el Excel trae dos veces la misma línea ya formateada, chocarían
    por_folio = defaultdict(list)
    for linea in salida:
        por_folio[(normalizar_id(linea["IDProveedor"]), linea["Folio"])].append(linea["Fila"])
    problemas = [(fila, "Folio", f"El folio {folio} está repetido para el mismo proveedor (filas "
                                 f"{', '.join(map(str, filas))}): SAP lo rechazaría. ¿Se copió la línea dos veces?")
                 for (_, folio), filas in por_folio.items() if len(filas) > 1 for fila in filas]
    return salida, problemas


def escribir(lineas):
    """Líneas -> bytes del TXT para SAP: 11 columnas separadas por tabulador, sin encabezado, CRLF, ANSI."""
    texto = "".join("\t".join(linea[c] for c in COLUMNAS) + SALTO for linea in lineas)
    return texto.encode(CODIFICACION)


def registros_de(lineas, archivo):
    """Lo que se guarda en HistorialFolios.csv al descargar el TXT (rastreo del folio nuevo al original)."""
    return [{"Folio": l["Folio"], "Clave": l["Clave"], "IDProveedor": normalizar_id(l["IDProveedor"]),
             "Periodo": folios.texto_quincena(l["Fecha"]),
             "Linea": l["Fila"], "FolioOriginal": l["FolioOriginal"], "Monto": l["Monto"], "Fecha": l["Fecha"],
             "Archivo": archivo}
            for l in lineas]


def periodo(lineas):
    """(año, mes, "Q18-2026") que más se repite en las líneas: para nombrar el TXT y guardar su copia."""
    meses = Counter(datetime.strptime(l["Fecha"], "%d.%m.%Y").strftime("%Y-%m") for l in lineas)
    anio, mes = map(int, meses.most_common(1)[0][0].split("-"))
    return anio, mes, quincenas(lineas).most_common(1)[0][0]


def _libro_en_bytes(libro):
    salida = BytesIO()
    libro.save(salida)
    return salida.getvalue()


def plantilla():
    """Excel vacío con las columnas A a J (en formato texto, para que Excel no cambie lo que se escriba)."""
    libro = Workbook()
    hoja = libro.active
    hoja.title = HOJA_PARA_TXT
    hoja.append(COLUMNAS_EXCEL)
    for i, columna in enumerate(COLUMNAS_EXCEL, start=1):
        letra = get_column_letter(i)
        hoja.column_dimensions[letra].width = 16
        hoja.column_dimensions[letra].number_format = "@"
        hoja.cell(1, i).font = Font(bold=True, color="FFFFFF")
        hoja.cell(1, i).fill = PatternFill("solid", fgColor="305496")
    hoja.freeze_panes = "A2"
    return _libro_en_bytes(libro)


def control_de_folios(lineas, nombres):
    """Excel de control del analista: de qué folio original salió cada folio nuevo."""
    libro = Workbook()
    hoja = libro.active
    hoja.title = "Control de folios"
    encabezados = ["Fila del Excel", "Médico", "IDProveedor", "Folio original", "Folio nuevo", "Monto", "Fecha",
                   "Tax", "Referencia", "Folio generado antes"]
    hoja.append(encabezados)
    for l in lineas:
        previo = l.get("Previo")
        hoja.append([l["Fila"], nombres.get(normalizar_id(l["IDProveedor"]), ""), l["IDProveedor"],
                     l["FolioOriginal"], l["Folio"], float(l["Monto"]), l["Fecha"], l["Tax"],
                     l["Referencia"],
                     f"Sí, el {previo['Registrado']}" if previo else ""])
        hoja.cell(hoja.max_row, 6).number_format = "#,##0.00"
    for i, encabezado in enumerate(encabezados, start=1):
        hoja.cell(1, i).font = Font(bold=True, color="FFFFFF")
        hoja.cell(1, i).fill = PatternFill("solid", fgColor="305496")
        hoja.column_dimensions[get_column_letter(i)].width = 34 if encabezado == "Médico" else 20
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    return _libro_en_bytes(libro)
