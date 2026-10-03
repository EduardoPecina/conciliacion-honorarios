"""Trabajo de varios días: el portal guarda todo lo del mes en la carpeta del mes y ofrece continuar al volver.
La lógica está en honorarios/servicio.py (guardar_trabajo, cargar_trabajo); aquí solo se muestra."""

from pathlib import Path

import streamlit as st

from honorarios.servicio import cargar_trabajo, guardar_trabajo, trabajos_guardados
from honorarios.utilidades import ErrorUsuario

ss = st.session_state


def guardar():
    """Guarda el trabajo actual (si hay una conciliación en curso)."""
    if ss.get("conciliacion") is not None and ss.get("salida") and ss.get("txt_original"):
        guardar_trabajo(ss.conciliacion, ss.salida, ss.txt_original)


def continuar(ruta):
    try:
        conciliacion, txt_original, carpeta = cargar_trabajo(ruta)
    except ErrorUsuario as e:
        ss.error_trabajo = str(e)
        return
    ss.conciliacion, ss.txt_original, ss.salida = conciliacion, txt_original, carpeta
    ss.anio_sel, ss.mes_sel = conciliacion.anio, conciliacion.mes        # el selector muestra el mes recuperado
    ss.excel = Path(conciliacion.destino).read_bytes() if Path(conciliacion.destino).exists() else None
    ss.excel_version = None if conciliacion.editada or ss.excel is None else conciliacion.version
    ss.pop("etxt", None)                                 # Edición TXT se vuelve a armar (con su borrador)


def ofrecer_continuar():
    """Si no hay nada abierto y hay trabajo guardado, ofrece continuarlo."""
    if ss.get("error_trabajo"):
        st.error(ss.pop("error_trabajo"))
    if ss.get("conciliacion") is not None:
        return
    trabajos = trabajos_guardados()
    if not trabajos:
        return
    with st.container(border=True):
        st.markdown("**Continuar donde te quedaste**")
        for ruta, nombre, fecha in trabajos:
            st.button(f"Continuar {nombre} (última modificación: {fecha})", key=f"continuar_{ruta.stem}",
                      on_click=continuar, args=(ruta,))
        st.caption("Se recupera todo: el resultado, lo consultado al SAT y a SAP, y tus cambios en Edición TXT. "
                   "Para empezar de nuevo, sube los archivos abajo y procesa.")
