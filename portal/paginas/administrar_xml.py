"""Página Administrar carpeta de XML: quitar de un ZIP los XML que no deben estar (facturas de otro periodo, copias
repetidas...) y descargar un ZIP nuevo sin ellos. El ZIP original no se toca.

Sin reglas de negocio: solo archivos (portal/archivos.py). Qué quitar lo decide el usuario, con el Excel de la
conciliación (columna ArchivoXML de las hojas "XML sin TXT" o Detalle)."""

import hashlib
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from honorarios.utilidades import ErrorUsuario
from portal.archivos import buscar_en_zip, leer_zip, zip_sin

st.title("Administrar carpeta de XML")
st.caption("Quita del ZIP los XML que no deben estar (facturas de otro periodo, copias repetidas, facturas que no "
           "son de esta quincena…) y descarga un ZIP nuevo sin ellos. **El ZIP original no se modifica.**")

subido = st.file_uploader("ZIP con los XML", type=["zip"])
pegado = st.text_area(
    "XML que quieres quitar (uno por renglón)", height=160,
    placeholder="Facturas Q18\\Septiembre\\A1872.xml\nA1890.xml\n6F9619FF-8B86-D011-B42D-00C04FC964FF",
    help="Puedes pegar la columna ArchivoXML del Excel de la conciliación (hoja XML sin TXT o Detalle) tal cual, "
         "solo el nombre del archivo, o el UUID. El encabezado de la columna y los renglones vacíos se ignoran.")
con_pdf = st.checkbox("Quitar también el PDF (u otro archivo) que tenga el mismo nombre", value=True,
                      help="Por ejemplo, al quitar A1872.xml también se quita A1872.pdf de la misma carpeta.")
if subido is None:
    st.info("Sube el ZIP y pega los XML que quieres quitar. En el Excel de la conciliación, la columna **ArchivoXML** "
            "dice la ruta de cada XML; la hoja **XML sin TXT** trae los que ninguna línea usó y, en **Qué pasó**, "
            "por qué.")
    st.stop()


@st.cache_data(show_spinner="Leyendo el ZIP...", max_entries=3)
def contenido_del_zip(nombre, huella, _datos):
    """Los XML del ZIP con su UUID, RFC, fecha y total (se lee una sola vez por ZIP; la huella distingue un ZIP de
    otro sin volver a revisar todo su contenido en cada clic)."""
    return leer_zip(nombre, _datos)[1]


datos = subido.getvalue()
try:
    xmls = contenido_del_zip(subido.name, hashlib.sha256(datos).hexdigest(), datos)
except ErrorUsuario as e:
    st.error(str(e))
    st.stop()

c = st.columns(3)
c[0].metric("XML en el ZIP", f"{len(xmls):,}")
pedidos = [r for r in pegado.splitlines() if r.strip()]
encontrados, problemas = buscar_en_zip(xmls, pedidos)
c[1].metric("A quitar", f"{len(encontrados):,}")
c[2].metric("Quedan", f"{len(xmls) - len(encontrados):,}")

if problemas:
    st.warning(f"{len(problemas)} renglón(es) no se encontraron en el ZIP:")
    st.dataframe(pd.DataFrame(problemas, columns=["Lo que pegaste", "Por qué"]), hide_index=True, width="stretch")
if not encontrados:
    if pedidos:
        st.info("Ninguno de los XML que pegaste está en este ZIP: no hay nada que quitar.")
    with st.expander(f"Ver los {len(xmls):,} XML del ZIP"):
        st.dataframe(pd.DataFrame(xmls).drop(columns="Ruta"), hide_index=True, width="stretch")
    st.stop()

st.subheader("Se van a quitar")
st.dataframe(pd.DataFrame(encontrados)[["ArchivoXML", "UUID", "RFCEmisor", "NombreEmisor", "FechaXML", "Total"]],
             hide_index=True, width="stretch")

archivo = leer_zip(subido.name, datos)[0]
nuevo, quitados = zip_sin(archivo, encontrados, con_pdf)
otros = [r for r in quitados if not r.lower().endswith(".xml")]
if otros:
    st.caption(f"También se quitan {len(otros)} archivo(s) con el mismo nombre: " + ", ".join(otros[:10])
               + ("…" if len(otros) > 10 else ""))

nombre_nuevo = f"{Path(subido.name).stem}_depurado.zip"
st.download_button(f"Descargar {nombre_nuevo}", nuevo, file_name=nombre_nuevo, mime="application/zip",
                   type="primary", width="stretch")
bitacora = BytesIO()
pd.DataFrame([{**{k: x[k] for k in ("ArchivoXML", "UUID", "RFCEmisor", "NombreEmisor", "FechaXML", "Total")},
               "ZIP original": subido.name} for x in encontrados]).to_excel(bitacora, index=False)
st.download_button("Descargar la lista de lo que se quitó (Excel)", bitacora.getvalue(),
                   file_name=f"{Path(subido.name).stem}_quitados.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   help="Para tu control, o para mandársela a quien armó la carpeta de XML.")
st.caption(f"El ZIP nuevo tiene las mismas carpetas y archivos que el original, sin los {len(quitados)} de arriba. "
           "Úsalo en Conciliador Automático en lugar del original.")
