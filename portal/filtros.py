"""Filtros y orden de la tabla de Edición TXT. Solo cambian lo que se VE: el TXT descargado siempre lleva todas
las líneas. Funciones puras (sin Streamlit) para poder probarlas."""

import pandas as pd

from honorarios.utilidades import importe_txt

CRUZADOS_LIMPIOS = ["Cruzado", "Cruzado (pago quebrado)"]
# Del más grave al menos grave (para "Gravedad del estatus")
GRAVEDAD = ["Formato incorrecto", "Rechazar", "Detenida", "Diferencia de importe", "Sin XML", "Línea sin IDProveedor",
            "ID no está en catálogo", "Proveedor sin RFC en catálogo", "RFC inválido en catálogo",
            "Proveedor extranjero", "Cruzado con observaciones", "Cruzado (pago quebrado)", "Cruzado", "—"]
CON = {
    "Problemas en el TXT": lambda fila: fila["Problemas del TXT"] != "",
    "Posible ID equivocado": lambda fila: fila["Sugerencia"] != "",
    "Editadas": lambda fila: fila["_editada"],
    "Folio sin formatear": lambda fila: fila["_sin_folio"],
}
ORDENES = ["Orden del archivo", "Gravedad del estatus", "Médico (A-Z)", "Monto (mayor a menor)", "Problemas primero"]
MODOS = ["Mostrar solo", "Ocultar"]
TAX = ["Ambos", "E0", "F2"]
VACIOS = {"f_modo": "Mostrar solo", "f_estatus": [], "f_con": [], "f_tax": "Ambos", "f_min": None, "f_max": None,
          "f_buscar": "", "f_orden": "Orden del archivo"}
PENDIENTES = {**VACIOS, "f_modo": "Ocultar", "f_estatus": list(CRUZADOS_LIMPIOS), "f_orden": "Gravedad del estatus"}
PARA_CORREGIR = {**VACIOS, "f_con": ["Problemas en el TXT", "Posible ID equivocado"]}


def monto(texto):
    valor = importe_txt(texto)
    return None if valor is None else float(valor)


def aplicar(tabla, f):
    """tabla: DataFrame de Edición TXT (con las columnas auxiliares _editada y _sin_folio).
    f: valores de los filtros (claves de VACIOS).
    Devuelve (vista ordenada, activos): activos = [(descripción, valores para quitar SOLO ese filtro)], en el
    orden en que se aplican. El portal muestra cada uno con su botón para quitarlo."""
    vista, activos = tabla, []
    if f["f_estatus"]:
        dentro = vista["Estatus"].isin(f["f_estatus"])
        vista = vista[dentro if f["f_modo"] == "Mostrar solo" else ~dentro]
        activos.append((f"{'Solo' if f['f_modo'] == 'Mostrar solo' else 'Ocultar'}: {', '.join(f['f_estatus'])}",
                        {"f_estatus": [], "f_modo": VACIOS["f_modo"]}))
    if f["f_con"]:
        vista = vista[vista.apply(lambda fila: any(CON[c](fila) for c in f["f_con"]), axis=1)] if len(vista) else vista
        activos.append(("Con: " + " o ".join(f["f_con"]), {"f_con": []}))
    if f["f_tax"] != "Ambos":
        vista = vista[vista["Tax"].str.strip().str.upper() == f["f_tax"]]
        activos.append((f"Tax {f['f_tax']}", {"f_tax": VACIOS["f_tax"]}))
    montos = pd.to_numeric(vista["Monto"].map(monto), errors="coerce")       # inválidos: NaN (no pasan el filtro)
    if f["f_min"] is not None:
        vista, montos = vista[montos >= f["f_min"]], montos[montos >= f["f_min"]]
        activos.append((f"Monto ≥ ${f['f_min']:,.2f}", {"f_min": None}))
    if f["f_max"] is not None:
        vista = vista[montos <= f["f_max"]]
        activos.append((f"Monto ≤ ${f['f_max']:,.2f}", {"f_max": None}))
    if f["f_buscar"].strip():
        texto = f["f_buscar"].strip().lower()
        columnas = ["Médico", "IDProveedor", "Folio", "Referencia"]
        vista = vista[vista[columnas].apply(lambda fila: any(texto in str(v).lower() for v in fila), axis=1)] \
            if len(vista) else vista
        activos.append((f"Búsqueda: \"{f['f_buscar'].strip()}\"", {"f_buscar": ""}))
    if f["f_orden"] != VACIOS["f_orden"]:
        activos.append((f"Orden: {f['f_orden']}", {"f_orden": VACIOS["f_orden"]}))
    return ordenar(vista, f["f_orden"]), activos


def ordenar(vista, orden):
    if orden == "Gravedad del estatus":
        clave = vista["Estatus"].map(lambda e: GRAVEDAD.index(e) if e in GRAVEDAD else len(GRAVEDAD))
    elif orden == "Médico (A-Z)":
        clave = vista["Médico"].str.upper()
    elif orden == "Monto (mayor a menor)":
        clave = vista["Monto"].map(lambda t: -monto(t) if monto(t) is not None else float("inf"))   # inválidos al final
    elif orden == "Problemas primero":
        clave = vista.apply(lambda f: 0 if f["Problemas del TXT"] else (1 if f["Sugerencia"] else 2), axis=1)
    else:
        return vista.sort_values("Línea")
    return vista.assign(_clave=clave).sort_values(["_clave", "Línea"]).drop(columns="_clave")
