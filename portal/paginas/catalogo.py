"""Página Catálogo: consulta del catálogo de proveedores. Por ahora se lee del Excel de _Configuracion;
cuando exista la base de datos, las altas y cambios se harán aquí."""

import streamlit as st

from honorarios.config import ARCHIVO_CATALOGO, CONFIG, SOCIEDAD_SAP
from honorarios.entradas import leer_catalogo
from honorarios.utilidades import ErrorUsuario

st.title(f"Catálogo de Proveedores {SOCIEDAD_SAP}")
st.info(f"Por ahora el catálogo se edita en el Excel **{ARCHIVO_CATALOGO}** (carpeta _Configuracion). "
        "Esta página solo lo consulta.")

ruta = CONFIG / ARCHIVO_CATALOGO
if not ruta.exists():
    st.error(f"No se encontró el catálogo en {ruta}")
    st.stop()
try:
    catalogo = leer_catalogo(ruta)
except ErrorUsuario as e:
    st.error(str(e))
    st.stop()

tipos = catalogo["TipoRFC"].value_counts()
c = st.columns(4)
c[0].metric("Proveedores", f"{len(catalogo):,}")
c[1].metric("RFC mexicano", int(tipos.get("MX", 0)))
c[2].metric("Extranjeros", int(tipos.get("Extranjero", 0)))
c[3].metric("Sin RFC o RFC inválido", int(tipos.get("Vacío", 0) + tipos.get("Inválido", 0)))

buscar = st.text_input("Buscar por ID, nombre o RFC").strip().upper()
vista = catalogo
if buscar:
    texto = (catalogo["IDProveedor"] + " " + catalogo["Nombre"].str.upper() + " " + catalogo["RFC"])
    vista = catalogo[texto.str.contains(buscar, regex=False)]
# Todo como texto: el Excel puede traer columnas con números y texto mezclados (p. ej. la columna histórica Activo)
st.dataframe(vista.fillna("").astype(str), hide_index=True, width="stretch")
