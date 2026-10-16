"""Página Generar TXT: del Excel con las líneas aprobadas al TXT para SAP.

El portal NO edita nada aquí: lee y revisa el Excel con honorarios/txt_sap.py y pone los folios con
honorarios/folios.py. Si una celda no sirve, se dice en qué fila y columna, y se corrige en el Excel."""

import hashlib
from datetime import datetime

import pandas as pd
import streamlit as st

from honorarios import folios, txt_sap
from honorarios.config import ARCHIVO_CATALOGO, CONFIG, MESES, REPORTES
from honorarios.entradas import leer_catalogo
from honorarios.utilidades import ErrorUsuario, normalizar_id

ss = st.session_state

st.title("Generar TXT")
st.caption("Del Excel con las líneas **listas para pagar** al TXT para SAP. El portal no edita nada: revisa cada "
           "celda y, si algo no sirve, te dice en qué fila y columna corregirlo en tu Excel.")

with st.expander("¿Cómo armo el Excel?"):
    st.markdown(
        "1. Descarga el Excel de la conciliación en **Conciliador Automático** y abre la hoja **TXT para SAP**: trae "
        "las líneas del TXT con sus 11 columnas y, de apoyo, el Estatus y el Médico.\n"
        "2. Deja solo las líneas listas para pagar (filtra por Estatus y borra las demás). Lo que sigue en revisión "
        "se queda fuera y se paga después, cuando esté listo.\n"
        "3. Guarda y sube el Excel aquí. Se leen solo las columnas **A a K**, en este orden: "
        + ", ".join(f"**{c}**" for c in txt_sap.COLUMNAS) + ". El encabezado es opcional.\n\n"
        "Si armas un Excel nuevo, usa la plantilla (sus columnas ya están como texto: Excel no les quita los ceros).")
    st.download_button("Descargar plantilla vacía", txt_sap.plantilla(), file_name="Plantilla_TXT_SAP.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@st.cache_data(show_spinner=False)
def nombres_del_catalogo(modificado):
    """{IDProveedor: Nombre} del catálogo (solo para mostrar el médico). modificado: para releerlo si cambia."""
    try:
        catalogo = leer_catalogo(CONFIG / ARCHIVO_CATALOGO)
    except (ErrorUsuario, OSError):
        return None
    return dict(zip(catalogo["IDProveedor"], catalogo["Nombre"]))


ruta_catalogo = CONFIG / ARCHIVO_CATALOGO
nombres = nombres_del_catalogo(ruta_catalogo.stat().st_mtime if ruta_catalogo.exists() else 0)

subido = st.file_uploader("Excel con las líneas aprobadas", type=["xlsx"])
if subido is None:
    st.stop()
datos = subido.getvalue()
huella = hashlib.sha256(datos).hexdigest()
if ss.get("gtxt", {}).get("huella") != huella:              # Excel nuevo: se empieza de cero
    try:
        hoja, lineas, problemas = txt_sap.leer_excel(datos)
    except ErrorUsuario as e:
        st.error(str(e))
        st.stop()
    ss.gtxt = {"huella": huella, "archivo": subido.name, "hoja": hoja, "lineas": lineas, "problemas": problemas,
               "formateadas": None}
g = ss.gtxt
if ss.get("gtxt_aviso"):
    st.success(ss.pop("gtxt_aviso"))


def medico(id_proveedor):
    return (nombres or {}).get(normalizar_id(id_proveedor), "")


def tabla_de_problemas(problemas):
    return pd.DataFrame([{"Fila": f, "Columna": f"{txt_sap.LETRAS[c]} · {c}", "Qué corregir": t}
                         for f, c, t in sorted(problemas, key=lambda p: (p[0], txt_sap.COLUMNAS.index(p[1])))])


st.caption(f"Se leyó la hoja **«{g['hoja']}»** de {g['archivo']}.")
if g["problemas"]:
    filas = sorted({f for f, _, _ in g["problemas"]})
    st.error(f"No se puede generar el TXT: {len(filas)} fila(s) del Excel tienen celdas que SAP rechazaría. "
             "Corrígelas en tu Excel, guárdalo y vuelve a subirlo.")
    st.dataframe(tabla_de_problemas(g["problemas"]), hide_index=True, width="stretch")
    st.stop()

# ------------------------------------------------------------------ Revisión
lineas = g["formateadas"] or g["lineas"]
total = sum(float(l["Monto"]) for l in lineas)
c = st.columns(4)
c[0].metric("Líneas", f"{len(lineas):,}")
c[1].metric("Monto total", f"${total:,.2f}")
c[2].metric("Proveedores", len({normalizar_id(l["IDProveedor"]) for l in lineas}))
c[3].metric("Folios", "Formateados" if g["formateadas"] else "Por formatear")
st.caption("Compara las líneas y el monto total con tu Excel antes de descargar.")

repetidas = txt_sap.misma_factura_repetida(g["lineas"])
if repetidas:
    st.warning("**Posible pago doble:** estas filas tienen el mismo IDProveedor, folio, monto y fecha: "
               + "; ".join("filas " + ", ".join(map(str, filas)) for filas in repetidas)
               + ". Si no es un pago quebrado en partes iguales, quita las de más en tu Excel y vuelve a subirlo.")
if nombres is None:
    st.info("No se encontró el catálogo de proveedores en _Configuracion: la columna Médico queda vacía.")
else:
    sin_nombre = sorted({l["IDProveedor"] for l in lineas if not medico(l["IDProveedor"])})
    if sin_nombre:
        st.warning(f"{len(sin_nombre)} IDProveedor no están en el catálogo ({', '.join(sin_nombre[:10])}"
                   f"{'…' if len(sin_nombre) > 10 else ''}): revisa que el ID sea el correcto. SAP le paga a la "
                   "cuenta del ID que va en el TXT.")


def formatear():
    formateadas, problemas = txt_sap.formatear(g["lineas"], folios.cargar_folios())
    if problemas:
        g["problemas"] = problemas
    else:
        g["formateadas"] = formateadas


st.button("Formatear folios", type="secondary" if g["formateadas"] else "primary", on_click=formatear,
          disabled=bool(g["formateadas"]),
          help="Pone a cada línea un folio único de 16 caracteres que se puede rastrear: año y quincena, los últimos "
               "4 del IDProveedor y 8 caracteres calculados con la factura (misma factura = mismo folio; nunca se "
               "repite). La Referencia cambia igual: folio nuevo y el resto unidos con guion bajo "
               "(\"A1893      Q18-2026\" -> \"26185113K7Q3MXP4_Q18-2026\").")

# Fila, Médico y Folio original son de apoyo; después, las 11 columnas en el orden en que van en el TXT
vista = pd.DataFrame([{"Fila": l["Fila"], "Médico": medico(l["IDProveedor"]),
                       **({"Folio original": l["FolioOriginal"]} if g["formateadas"] else {}),
                       **{col: l[col] for col in txt_sap.COLUMNAS}} for l in lineas])
amarillo = "background-color: rgba(234, 179, 8, 0.28)"
st.dataframe(vista.style.map(lambda _: amarillo, subset=["Folio", "Referencia"]) if g["formateadas"] else vista,
             hide_index=True, width="stretch",
             column_config={"Fila": st.column_config.NumberColumn("Fila", help="Fila de tu Excel", format="%d"),
                            "Folio": st.column_config.TextColumn("Folio", width="medium"),
                            "Referencia": st.column_config.TextColumn("Referencia", width="large")})
st.caption("Fila, Médico" + (" y Folio original son" if g["formateadas"] else " son") + " de apoyo: no van en el TXT. "
           + ("🟨 Folio y Referencia ya formateados." if g["formateadas"]
              else "Formatea los folios para poder descargar el TXT."))

# ------------------------------------------------------------------ Descargar
if not g["formateadas"]:
    st.stop()
formateadas = g["formateadas"]
previas = [l for l in formateadas if l["Previo"]]
if previas:
    with st.expander(f"{len(previas)} línea(s) ya tenían este folio generado antes", expanded=False):
        st.caption("Misma factura = mismo folio. Si esas líneas ya se cargaron a SAP, SAP las rechaza por folio "
                   "repetido: no se pagan dos veces. Si no se cargaron (por ejemplo, la carga falló), no pasa nada.")
        st.dataframe(pd.DataFrame([{"Fila": l["Fila"], "Médico": medico(l["IDProveedor"]), "Folio": l["Folio"],
                                    "Generado": l["Previo"].get("Registrado", ""),
                                    "Archivo": l["Previo"].get("Archivo", "")} for l in previas]),
                     hide_index=True, width="stretch")

anio, mes, quincena = txt_sap.periodo(formateadas)
contenido = txt_sap.escribir(formateadas)
nombre_txt = f"TXT_SAP_{quincena}.txt"
control = txt_sap.control_de_folios(formateadas, nombres or {})


def registrar_descarga():
    """Al descargar: los folios quedan en HistorialFolios.csv (rastreo) y una copia del TXT y del control de folios
    en la carpeta del mes."""
    folios.registrar_folios(txt_sap.registros_de(formateadas, g["archivo"]))
    carpeta = REPORTES / str(anio) / MESES[mes - 1]
    marca = datetime.now().strftime("%Y%m%d_%H%M")
    try:
        carpeta.mkdir(parents=True, exist_ok=True)
        (carpeta / f"TXT_SAP_{quincena}_{marca}.txt").write_bytes(contenido)
        (carpeta / f"ControlFolios_{quincena}_{marca}.xlsx").write_bytes(control)
        ss.gtxt_aviso = (f"Se guardó una copia del TXT y del control de folios en Reportes\\{anio}\\"
                         f"{MESES[mes - 1]}.")
    except OSError:
        ss.gtxt_aviso = "Los folios quedaron registrados (no se pudo guardar la copia del TXT en Reportes)."


st.subheader("Descargar TXT para SAP")
st.download_button("Descargar TXT", contenido, file_name=nombre_txt, mime="text/plain", type="primary",
                   width="stretch", on_click=registrar_descarga)
st.caption(f"**{len(formateadas):,} líneas** por **${total:,.2f}**. Mismo formato que el TXT original: 11 columnas "
           "separadas por tabulador, sin encabezados. Al descargar, los folios quedan registrados en "
           "_Configuracion\\HistorialFolios.csv y se guarda una copia en la carpeta del mes.")
st.download_button("Descargar control de folios (Excel)", control,
                   file_name=f"ControlFolios_{quincena}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   help="Para tu control: de qué folio original salió cada folio nuevo, con médico, monto y fecha.")
