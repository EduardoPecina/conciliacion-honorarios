"""Página Historial: UUID ya conciliados (los que se registran al confirmar una conciliación)."""

import pandas as pd
import streamlit as st

from honorarios.config import ARCHIVO_HISTORIAL, CONFIG

st.title("Historial de UUID conciliados")
st.caption("Facturas que ya se pagaron en algún periodo. Si una vuelve a aparecer en otro mes, "
           "la conciliación la detiene como posible pago doble.")

ruta = CONFIG / ARCHIVO_HISTORIAL
if not ruta.exists():
    st.info("Todavía no hay historial: se crea al confirmar la primera conciliación.")
    st.stop()
try:
    historial = pd.read_csv(ruta, dtype=str, encoding="utf-8-sig").fillna("")
except Exception as e:                                           # noqa: BLE001 - se muestra al usuario
    st.error(f"No se pudo leer {ARCHIVO_HISTORIAL}: {e}")
    st.stop()

periodos = sorted(historial["Periodo"].unique(), reverse=True)
c = st.columns(2)
c[0].metric("UUID registrados", f"{historial['UUID'].nunique():,}")
c[1].metric("Periodos", len(periodos))

periodo = st.selectbox("Periodo", ["Todos"] + periodos)
buscar = st.text_input("Buscar UUID o IDProveedor").strip().upper()
vista = historial if periodo == "Todos" else historial[historial["Periodo"] == periodo]
if buscar:
    vista = vista[(vista["UUID"].str.upper() + " " + vista["IDProveedor"]).str.contains(buscar, regex=False)]
st.dataframe(vista, hide_index=True, use_container_width=True)
