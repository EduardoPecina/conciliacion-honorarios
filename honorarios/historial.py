"""Historial de UUID ya conciliados en meses anteriores (evita pagar dos veces la misma factura)."""

from datetime import datetime

import pandas as pd

from .config import ARCHIVO_HISTORIAL, CONFIG
from .utilidades import guardar_seguro


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
