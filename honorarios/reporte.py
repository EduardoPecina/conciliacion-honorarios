"""Generación del Excel de salida: Resumen con fórmulas, Detalle, desglose de impuestos, hojas auxiliares."""

from datetime import datetime
from decimal import Decimal

from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd

from .config import (
    CP_SOCIEDAD,
    FORMAS_PAGO,
    FORMAS_PAGO_OK,
    MESES,
    METODOS_PAGO_OK,
    NOMBRE_SOCIEDAD,
    REGIMEN_RECEPTOR,
    RFC_SOCIEDAD,
    TASA_ISR_GENERAL,
    TASA_ISR_RESICO,
    TASA_IVA,
    TOLERANCIA,
    TOLERANCIA_IMPUESTOS,
    UMBRAL_OK,
    USOS_CFDI_OK,
)
from .utilidades import huella


# ------------------------------------------------------------------
# 4. REPORTE
# ------------------------------------------------------------------
ORDEN_ESTATUS = ["Cruzado", "Cruzado (pago quebrado)",
                 "Cruzado con observaciones", "Detenida", "Rechazar", "Diferencia de importe", "Sin XML",
                 "Proveedor extranjero", "RFC inválido en catálogo",
                 "Proveedor sin RFC en catálogo", "ID no está en catálogo", "Línea sin IDProveedor"]
# Quién tiene que actuar en cada estatus (misma tabla del README). La usa el portal para la lista "Qué hacer".
RESPONSABLE_POR_ESTATUS = {
    "Cruzado": "Nadie: se puede pagar",
    "Cruzado (pago quebrado)": "Nadie: se puede pagar",
    "Cruzado con observaciones": "Revisor (ver Observaciones)",
    "Detenida": "Customer Services",
    "Rechazar": "Médico: refacturar",
    "Diferencia de importe": "Revisor",
    "Sin XML": "Responsables de envío de XML",
    "Proveedor extranjero": "Revisor (factura del extranjero)",
    "RFC inválido en catálogo": "Administración del catálogo",
    "Proveedor sin RFC en catálogo": "Administración del catálogo",
    "ID no está en catálogo": "Administración del catálogo",
    "Línea sin IDProveedor": "Administración del catálogo",
}
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


def como_texto(celdas):
    """Un texto que empieza con "=" (p. ej. el nombre del emisor de un XML) openpyxl lo guarda como
    fórmula y Excel lo ejecutaría. Se fuerza a texto: el contenido se ve igual, pero no se ejecuta."""
    for celda in celdas:
        if celda.data_type == "f":
            celda.data_type = "s"


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
    OB, ES, EF = col("Observaciones"), col("EstadoSAT"), col("EFOS")
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
         "El SAT, SAP o el historial de UUID no se pudieron consultar para esas líneas"),
        ("Líneas con factura cancelada en el SAT", f"=COUNTIF({rango(ES)},\"*Cancelado*\")",
         "Líneas del Detalle con EstadoSAT = Cancelado"),
        ("Líneas con alerta EFOS (69-B)", f"=COUNTIF({rango(EF)},\"Revisar*\")",
         "El SAT indica que el emisor podría estar en la lista de operaciones simuladas (columna EFOS)"),
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
                           f"Hasta {TOLERANCIA_IMPUESTOS} se acepta como redondeo")
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
        como_texto(hoja[hoja.max_row][:3])
        primera, ultima = resumen + 1, resumen + len(grupo)
        for fila in grupo.itertuples(index=False):
            k = fila.FilaExcel
            hoja.append([f"    Línea {fila.Linea}", fila.Folio, None, fila.Estatus])
            r = hoja.max_row
            como_texto([hoja.cell(r, 2)])
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
        (f"     • 'Revisión importe': OK (hasta {UMBRAL_OK}) · Redondeo (hasta {TOLERANCIA}, se acepta).", False),
        (f"     • 'Revisión impuestos': OK (hasta {UMBRAL_OK}) · Redondeo (hasta {TOLERANCIA_IMPUESTOS}, se acepta) · "
         f"Revisar (más de {TOLERANCIA_IMPUESTOS}).", False),
        ("     • El vínculo 'Ver cálculo' lleva a la ficha de esa línea.", False),
        ("4. Cálculos: la ficha de cada línea, paso a paso y explicada en palabras.", False),
        ("5. Tasas: los porcentajes que usan las fórmulas (IVA 16%, ISR 10%, RESICO 1.25%...).", False),
        ("6. TXT para SAP: las líneas del TXT con sus 11 columnas (A a K) y, de apoyo, Estatus, Médico y Línea.", False),
        ("     • Deja solo las líneas listas para pagar (filtra por Estatus y borra las demás) y sube este Excel", False),
        ("       en la página Generar TXT del portal: ahí se ponen los folios únicos y se descarga el TXT para SAP.", False),
        ("", False),
        ("QUÉ SIGNIFICA CADA ESTATUS", True),
        ("", False),
        ("• Cruzado: pasó TODAS las validaciones. No hay nada que revisar.", False),
        ("     Se validó: importe contra TXT, timbre, versión 4.0, datos del receptor (RFC, régimen, uso CFDI),", False),
        ("     periodo, tipo I, método y forma de pago (PUE con 03 o 99 / PPD con 99), moneda, total del XML, impuestos,", False),
        ("     códigos de retención en SAP, que la factura siga vigente en el SAT", False),
        ("     y que el UUID no se haya", False),
        ("     conciliado en otro mes ni venga repetido.", False),
        ("• Cruzado (pago quebrado): igual que Cruzado, pero varias líneas del TXT y/o varias facturas", False),
        ("     del mismo proveedor se cruzan como grupo (p. ej. 500 + 300 + 200 del TXT contra una factura", False),
        ("     de 1,000, o 600 + 400 contra dos facturas de 500).", False),
        ("• Cruzado con observaciones: cruzó, pero hay algo que revisar (ver Observaciones), p. ej. alerta EFOS. Incluye las", False),
        ("     líneas en las que alguna validación NO se pudo hacer (SAP pedido pero cerrado, historial ilegible...).", False),
        ("• Detenida: no se debe pagar hasta corregir (impuestos, SAP, pago doble de un UUID...).", False),
        ("     Si el mismo XML viene copiado dos veces en la carpeta pero solo una línea lo", False),
        ("     paga, sale 'Cruzado con observaciones': Verificar con Responsables de Envío de XML.", False),
        ("• Proveedor extranjero: el catálogo trae un ID fiscal de otro país; no hay CFDI que cruzar.", False),
        ("• RFC inválido en catálogo: el RFC parece mexicano pero está mal escrito; corregir el catálogo.", False),
        ("• Rechazar: la factura no es válida (sin timbre, cancelada en el SAT, receptor, periodo, forma o", False),
        ("     método de pago incorrecto...). Hay que pedir otra.", False),
        ("", False),
        ("Colores: verde = OK · amarillo = redondeo aceptado · rojo = revisar. "
         "En la hoja Tasas, las celdas amarillas (tasas y tolerancias) se pueden cambiar.", False),
    ]
    for texto, negrita in textos:
        hoja.append([texto])
        hoja.cell(hoja.max_row, 1).font = Font(bold=negrita, size=14 if negrita else 12)
    hoja.column_dimensions["A"].width = 110
    hoja.sheet_view.showGridLines = False


def hoja_para_txt(resultado, ruta_txt):
    """Hoja "TXT para SAP": las líneas del TXT original con sus 11 columnas tal cual (A a K, como texto) y, de apoyo,
    Estatus, Médico y Línea (L a N). El analista deja aquí solo las líneas listas para pagar y sube el Excel en
    Generar TXT (que lee solo de la A a la K)."""
    from .txt_sap import COLUMNAS, lineas_del_txt
    por_linea = {f["Linea"]: f for f in resultado}
    orden = {e: i for i, e in enumerate(ORDEN_ESTATUS)}
    filas = [valores + [por_linea.get(n, {}).get("Estatus", ""), por_linea.get(n, {}).get("Nombre", ""), n]
             for n, valores in lineas_del_txt(ruta_txt.read_bytes()).items()]
    filas.sort(key=lambda f: (orden.get(f[11], 99), f[1], f[13]))
    return pd.DataFrame(filas, columns=COLUMNAS + ["Estatus", "Médico", "Línea"])


def generar_reporte(rutas, anio, mes, resultado, sobrantes, errores_xml, n_xml, estados):
    detalle = pd.DataFrame(resultado)
    columnas = ["Linea", "Estatus", "Fecha", "IDProveedor", "Nombre", "RFC", "Folio",
                "CodigoImpuesto", "Importe", "Ver cálculo", "Revisión importe", "Variación importe",
                "Revisión impuestos", "Variación impuestos"] + COLS_DESGLOSE + [
                "RegimenEmisor", "NombreRegimen", "CodigosSAP", "RegimenSAP",
                "DiscrepanciaSAP", "FechaXML", "MetodoPago", "FormaPago", "UUID", "EstadoSAT", "EFOS", "ArchivoXML",
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
        {"Estatus": "Líneas con factura cancelada en el SAT",
         "Lineas": int(detalle["EstadoSAT"].fillna("").str.contains("Cancelado").sum())},
        {"Estatus": "Líneas con alerta EFOS (69-B)",
         "Lineas": int(detalle["EFOS"].fillna("").str.startswith("Revisar").sum())},
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
        para_txt = hoja_para_txt(resultado, rutas["txt"])
        para_txt.to_excel(w, sheet_name="TXT para SAP", index=False)
        for hoja in w.book.worksheets:
            if hoja.title != "Resumen":
                # Hasta aquí estas hojas solo tienen datos (las fórmulas del programa se agregan después)
                como_texto(c for fila in hoja.iter_rows() for c in fila)
                formatear_hoja(hoja, colorear_estatus=hoja.title in ("Detalle", "TXT para SAP"))
        hoja_txt = w.book["TXT para SAP"]
        hoja_txt.auto_filter.ref = hoja_txt.dimensions
        for fila in hoja_txt.iter_rows(min_row=2, max_col=11):
            for celda in fila:
                celda.number_format = "@"                  # texto: Excel no le quita los ceros ni los decimales
        hoja_det = w.book["Detalle"]
        hoja_det.auto_filter.ref = hoja_det.dimensions
        escribir_resumen(w.book["Resumen"], list(detalle.columns), len(detalle), n_xml)
        escribir_tasas(w.book)
        L = desglose_detalle(hoja_det, list(detalle.columns), detalle)
        fichas = escribir_calculos(w.book, L, detalle)
        escribir_por_proveedor(w.book, L, detalle, fichas)
        escribir_instrucciones(w.book)
        orden = ["Resumen", "Cómo usar", "Por proveedor", "Detalle", "TXT para SAP", "Cálculos",
                 "XML sin TXT", "Errores XML", "Tasas", "Info"]
        w.book._sheets.sort(key=lambda h: orden.index(h.title) if h.title in orden else 99)
        w.book.active = 0
    return destino, resumen
