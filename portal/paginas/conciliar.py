"""Página Conciliador Automático: subir TXT y XML -> Procesar -> resumen -> Descargar Excel -> Confirmar."""

import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from honorarios.config import ARCHIVO_FOLIOS, CONFIG, MESES, REPORTES
from honorarios.entradas import NOMBRES_MES, rutas_en
from honorarios.estado_sat import alerta_efos, describir
from honorarios.reporte import ORDEN_ESTATUS, RESPONSABLE_POR_ESTATUS
from honorarios.servicio import (AYUDA_REUSAR_SAP, actualizar_sap, actualizar_sat, agregar_pendientes,
                                 conciliar_periodo, dias_desde_consulta, es_cruzada, fecha_consultas_sap,
                                 lineas_pendientes, regenerar_reporte, registrar_historial)
from honorarios.utilidades import ErrorUsuario, fecha_actual
from portal.archivos import preparar_carpeta, preparar_xmls
from portal.estilo import aviso_al_salir, ruta_logo
from portal.trabajo import guardar, ofrecer_continuar

CRUZADOS = ("Cruzado", "Cruzado (pago quebrado)")



@st.cache_data(ttl=3600, show_spinner=False)
def fecha_de_hoy():
    """Fecha de internet (una consulta por hora, no en cada clic)."""
    return fecha_actual()


if ruta_logo():
    st.image(str(ruta_logo()), width=220)
st.title("Conciliación de Honorarios")
ofrecer_continuar()

# ------------------------------------------------------------------ 1. Periodo y archivos
# Solo el mes actual y los anteriores: un periodo que todavía no llega no se puede conciliar. Se ofrece también
# el año anterior (en enero se concilia diciembre).
hoy, fuente_fecha = fecha_de_hoy()
col_anio, col_mes = st.columns(2)
anios = [hoy.year - 1, hoy.year]
if st.session_state.get("anio_sel") not in anios:                     # por omisión: el mes vigente
    st.session_state["anio_sel"] = hoy.year
anio = col_anio.selectbox("Año", anios, key="anio_sel")
meses = list(range(1, (hoy.month if anio == hoy.year else 12) + 1))
if st.session_state.get("mes_sel") not in meses:
    st.session_state["mes_sel"] = meses[-1]
mes = col_mes.selectbox("Mes", meses, key="mes_sel", format_func=lambda m: NOMBRES_MES[m - 1].capitalize())
st.caption(f"Hoy es {hoy.day} de {NOMBRES_MES[hoy.month - 1]} de {hoy.year} (fecha de {fuente_fecha}). "
           "Solo se pueden conciliar el mes actual y los anteriores.")
mes_confirmado = True
if (anio, mes) != (hoy.year, hoy.month):
    nombre_periodo = f"{NOMBRES_MES[mes - 1]} {anio}"
    st.warning(f"**{nombre_periodo.capitalize()} no es el mes vigente** "
               f"({NOMBRES_MES[hoy.month - 1]} {hoy.year}). La conciliación se hace una vez al mes: revisa que "
               "sea el mes correcto.")
    mes_confirmado = st.checkbox(f"Sí, quiero conciliar {nombre_periodo}")

st.info("Recuerda: **el TXT, los XML y el mes elegido deben coincidir** para que el resultado sea correcto.")
txt = st.file_uploader("1. TXT de pagos", type=["txt"])
xmls = st.file_uploader("2. Facturas XML (arrastra todos los archivos a la vez, o un .zip)",
                        type=["xml", "zip"], accept_multiple_files=True)
if xmls:
    st.caption(f"{len(xmls):,} archivo(s) cargado(s)")

salida = REPORTES / str(anio) / MESES[mes - 1]              # la carpeta del mes se crea sola al procesar
with st.container(border=True):
    st.markdown("**Antes de procesar** (la validación contra SAP es obligatoria)\n\n"
                "✅ SAP abierto en el sistema productivo  \n"
                "✅ Sesión de SAP iniciada  \n"
                "✅ Acceso a la transacción **FK03** (lo otorga IT)  \n"
                "✅ Scripting de SAP GUI habilitado (lo configura IT)")
    st.caption("Durante la validación, el portal usa la **primera ventana de SAP** que abriste: guarda lo que "
               "tengas ahí. No uses SAP hasta que termine; puedes seguir trabajando en otros programas. "
               "Si tienes abiertas varias conexiones (por ejemplo, Productivo y Calidad), deja solo Productivo.")
reusar_sap = False
fecha_sap = fecha_consultas_sap({"salida": salida}, anio, mes)
if fecha_sap:
    dia, hora = fecha_sap.split()
    reusar_sap = st.checkbox(f"Reusar lo consultado a SAP el {dia} a las {hora} (más rápido)", value=False,
                             help=AYUDA_REUSAR_SAP)

procesar = st.button("Procesar", type="primary", disabled=not (txt and xmls and mes_confirmado), width="stretch")

# ------------------------------------------------------------------ 2. Proceso
if procesar:
    st.session_state.pop("conciliacion", None)
    st.session_state.pop("resumen_pendientes", None)
    aviso = st.empty()                                     # pregunta antes de salir de la página (no bloquea)
    with aviso:
        st.html(aviso_al_salir(True), unsafe_allow_javascript=True)
    nota = st.info("⏳ **Conciliación en proceso.** Puedes hacer scroll y usar otros programas (excepto SAP). "
                   "Si cambias de página o cierras esta pestaña, el proceso se detiene.")
    with st.status("Procesando...", expanded=True) as estado:
        barra = st.progress(0.0, text="Consulta al SAT")
        barra_sap = st.progress(0.0, text="Consulta a SAP")
        try:
            with tempfile.TemporaryDirectory() as tmp:
                n = preparar_carpeta(tmp, (txt.name, txt.getvalue()), [(x.name, x.getvalue()) for x in xmls])
                st.write(f"{n:,} XML listos")
                rutas = rutas_en(Path(tmp))
                rutas["salida"] = salida
                salida.mkdir(parents=True, exist_ok=True)
                conciliacion = conciliar_periodo(
                    rutas, anio, mes, avisar=st.write, reusar_sap=reusar_sap,
                    progreso_sat=lambda hechas, total: barra.progress(hechas / total,
                                                                      text=f"Consulta al SAT: {hechas}/{total}"),
                    progreso_sap=lambda hechas, total, id_prov: barra_sap.progress(
                        hechas / total, text=f"Consulta a SAP: {hechas}/{total} (proveedor {id_prov})"),
                    hoy=hoy)
            st.session_state["conciliacion"] = conciliacion
            st.session_state["txt_original"] = (txt.name, txt.getvalue())      # para regenerar el Excel
            st.session_state["salida"] = salida
            st.session_state["excel_version"] = conciliacion.version
            guardar()                                                          # para continuar otro día
            st.session_state["excel"] = Path(conciliacion.destino).read_bytes()
            estado.update(label="Listo", state="complete", expanded=False)
        except ErrorUsuario as e:
            estado.update(label="No se pudo completar", state="error")
            st.error(str(e))
        except Exception as e:                                   # noqa: BLE001 - se muestra al usuario
            estado.update(label="Error inesperado", state="error")
            st.error(f"Error inesperado: {type(e).__name__}: {e}\n\n"
                     "Toma captura de este mensaje y compártela con quien mantiene el programa.")
    with aviso:                                            # terminó (bien o con error): ya no se pregunta al salir
        st.html(aviso_al_salir(False), unsafe_allow_javascript=True)
    nota.empty()
    if "conciliacion" in st.session_state:
        st.rerun()                                         # carga limpia: muestra el resultado y apaga el aviso

if not procesar:
    # Fuera del proceso nunca se pregunta al salir. Se reafirma en cada carga: el aviso vive en la página principal
    # y el aviso que lo apaga al terminar no siempre alcanza a ejecutarse.
    st.html(aviso_al_salir(False), unsafe_allow_javascript=True)

# ------------------------------------------------------------------ 3. Resultado
conciliacion = st.session_state.get("conciliacion")
if conciliacion is None:
    st.stop()

periodo = f"{NOMBRES_MES[conciliacion.mes - 1]} {conciliacion.anio}"
st.header(f"Resultado · {periodo}")
for aviso in conciliacion.avisos:
    st.warning(aviso)

# Consultas al SAT y a SAP: si el trabajo se retoma días después, pudieron cambiar
dias = dias_desde_consulta(conciliacion)
with st.expander(f"Consultas al SAT y a SAP: {conciliacion.consultado or 'sin fecha'}", expanded=dias >= 1):
    if dias >= 1:
        st.warning(f"Las consultas tienen {dias} día(s). Desde entonces una factura pudo cancelarse en el SAT o "
                   "Customer Services pudo corregir algo en SAP: actualízalas antes de confirmar o de armar el "
                   "Excel para el TXT.")
    col_sat, col_sap = st.columns(2)
    actualizar = None
    if col_sat.button("Actualizar SAT", width="stretch", help="Vuelve a consultar el estatus de todas las facturas "
                                                            "en el SAT (≈1 minuto) y vuelve a cruzar."):
        actualizar = "SAT"
    if col_sap.button("Actualizar SAP", width="stretch", help="Vuelve a consultar en SAP a todos los proveedores "
                                                            "(abre SAP con tu sesión antes) y vuelve a cruzar."):
        actualizar = "SAP"
    if actualizar:
        with st.status(f"Consultando {actualizar}...", expanded=True) as estado:
            barra = st.progress(0.0)
            if actualizar == "SAT":
                nueva, texto = actualizar_sat(conciliacion, avisar=st.write,
                                              progreso=lambda h, t: barra.progress(h / t, text=f"SAT: {h}/{t}"))
            else:
                nueva, texto = actualizar_sap(conciliacion, st.session_state["salida"], avisar=st.write,
                                              progreso=lambda h, t, i: barra.progress(h / t, text=f"SAP: {h}/{t} ({i})"))
            estado.update(label=texto, state="complete" if nueva is not conciliacion else "error")
        if nueva is not conciliacion:
            st.session_state["conciliacion"] = conciliacion = nueva
            guardar()
            st.rerun()

detalle = pd.DataFrame(conciliacion.resultado)
por_estatus = detalle.groupby("Estatus")["Importe"].agg(["count", "sum"])
lineas = lambda *estatus: int(sum(por_estatus["count"].get(e, 0) for e in estatus))

c = st.columns(5)
c[0].metric("Líneas del TXT", f"{len(detalle):,}")
c[1].metric("✅ Cruzado", lineas(*CRUZADOS))
c[2].metric("🟡 Con observaciones", lineas("Cruzado con observaciones"))
c[3].metric("🟠 Detenida", lineas("Detenida"))
c[4].metric("🔴 Rechazar", lineas("Rechazar"))

consultas = list(conciliacion.estados_sat.values())
sat = pd.Series([describir(x) for x in consultas]).value_counts()
c = st.columns(5)
c[0].metric("SAT: vigentes", int(sat.get("Vigente", 0) + sat.get("Vigente (cancelación en proceso)", 0)))
c[1].metric("SAT: canceladas", int(sat.get("Cancelado", 0)))
c[2].metric("SAT: no encontradas", int(sat.get("No encontrado", 0)))
c[3].metric("SAT: sin respuesta", int(sat.get("Sin respuesta", 0)))
c[4].metric("Alerta EFOS (69-B)", sum(1 for x in consultas if alerta_efos(x)))

observaciones = detalle.get("Observaciones", pd.Series(dtype=str)).fillna("")
c = st.columns(5)
c[0].metric("SAP: líneas con discrepancia", int(detalle.get("DiscrepanciaSAP", pd.Series(dtype=bool)).fillna(False).sum()))
c[1].metric("SAP: líneas sin validar", int(observaciones.str.contains("códigos de retención en SAP").sum()))
st.caption(f"SAP: {conciliacion.estados.get('Validación SAP', '')}")

st.subheader("Qué hacer")
pendientes = [{"Estatus": e, "Líneas": int(por_estatus.at[e, "count"]),
               "Importe": float(por_estatus.at[e, "sum"]), "Quién actúa": RESPONSABLE_POR_ESTATUS[e]}
              for e in ORDEN_ESTATUS if e in por_estatus.index and e not in CRUZADOS]
if pendientes:
    st.dataframe(pd.DataFrame(pendientes), hide_index=True, width="stretch",
                 column_config={"Importe": st.column_config.NumberColumn(format="$%,.2f")})
else:
    st.success("Todas las líneas cruzaron sin observaciones.")
if conciliacion.sobrantes:
    st.caption(f"{len(conciliacion.sobrantes):,} XML sin línea en el TXT (hoja 'XML sin TXT' del Excel). "
               f"{len(conciliacion.errores_xml):,} XML no se pudieron leer (hoja 'Errores XML').")



def txt_generado_despues_del_excel():
    """Se descargó un TXT en Generar TXT después de hacer el Excel: su columna "TXT generado" ya no está al día."""
    registro, excel = CONFIG / ARCHIVO_FOLIOS, Path(conciliacion.destino)
    return registro.exists() and excel.exists() and registro.stat().st_mtime > excel.stat().st_mtime


if (st.session_state.get("excel_version") != conciliacion.version or not st.session_state.get("excel")
        or txt_generado_despues_del_excel()):
    # Hubo cambios (actualizar SAT o SAP, facturas pendientes, un TXT generado): Excel nuevo con lo actual
    with st.spinner("Actualizando el Excel..."):
        conciliacion = regenerar_reporte(conciliacion, st.session_state["salida"], st.session_state["txt_original"])
        st.session_state["conciliacion"] = conciliacion
        st.session_state["excel"] = Path(conciliacion.destino).read_bytes()
        st.session_state["excel_version"] = conciliacion.version
        guardar()
st.download_button("Descargar Excel", st.session_state["excel"], file_name=Path(conciliacion.destino).name,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   type="primary", width="stretch")
st.info("**Para cargar a SAP:** en el Excel, hoja **TXT para SAP**, deja solo las líneas listas para pagar y "
        "súbelo en **Generar TXT**: ahí se formatean los folios y se descarga el TXT.")
st.page_link("paginas/generar_txt.py", label="Ir a Generar TXT", icon="📄")

# ------------------------------------------------------------------ 4. Facturas pendientes (segunda vuelta)
st.divider()
st.subheader("Facturas pendientes: refacturas y XML que faltaban")
resumen = st.session_state.get("resumen_pendientes")
if resumen:
    cruzaron = [c for c in resumen["cambios"] if c["Ahora"].startswith("Cruzado")]
    (st.success if cruzaron else st.info)(
        f"Se agregaron **{resumen['nuevas']} XML nuevos**: **{len(cruzaron)} línea(s) pasaron a Cruzado**."
        + (f" {resumen['sin_linea']} XML no encontraron línea pendiente (hoja 'XML sin TXT' del Excel)."
           if resumen["sin_linea"] else "")
        + (f" {len(resumen['ya_estaban'])} ya estaban en el mes (mismo UUID) y se ignoraron."
           if resumen["ya_estaban"] else "")
        + (f" {len(resumen['errores'])} no se pudieron leer (hoja 'Errores XML')." if resumen["errores"] else ""))
    if resumen["cambios"]:
        st.dataframe(pd.DataFrame(resumen["cambios"]), hide_index=True, width="stretch",
                     column_config={"Importe": st.column_config.NumberColumn(format="$%,.2f")})
        st.caption("El Excel de arriba ya está actualizado: en la hoja TXT para SAP, toma las líneas Cruzado que no "
                   "tienen TXT generado y súbelas en Generar TXT.")
sin_cruzar = lineas_pendientes(conciliacion)
if not sin_cruzar:
    st.success("No hay líneas pendientes: todas cruzaron.")
else:
    ya_cruzadas = sum(1 for f in conciliacion.resultado if es_cruzada(f))
    st.write(f"Las **{ya_cruzadas:,} líneas ya cruzadas no se tocan**. Sube **solo los XML nuevos** (refacturas de "
             f"las rechazadas o los XML que no habían llegado) y el portal los busca entre las **{len(sin_cruzar):,} "
             "líneas pendientes**, con las mismas reglas.")
    st.caption("Solo se consultan al SAT los XML nuevos (se necesita internet; SAP no). Una refactura puede venir "
               "fechada hasta el mes siguiente: queda Cruzado con observaciones. Un XML que ya estaba en el mes "
               "(mismo UUID) se ignora.")
    st.session_state.setdefault("pendientes_ver", 0)
    nuevos = st.file_uploader("XML nuevos (sueltos o en .zip)", type=["xml", "zip"], accept_multiple_files=True,
                              key=f"pendientes_{st.session_state['pendientes_ver']}")
    if st.button("Buscar en las líneas pendientes", type="primary", disabled=not nuevos):
        with st.status("Buscando en las líneas pendientes...", expanded=True) as estado:
            barra = st.progress(0.0, text="Consulta al SAT")
            try:
                with tempfile.TemporaryDirectory() as tmp:
                    # La subcarpeta queda en el nombre del XML (hoja Detalle): se sabe en qué vuelta llegó
                    carpeta = Path(tmp) / f"Pendientes {datetime.now():%d-%m-%Y %H%M}"
                    preparar_xmls(carpeta, [(x.name, x.getvalue()) for x in nuevos])
                    rutas_xml = sorted(r for r in carpeta.rglob("*") if r.suffix.lower() == ".xml")
                    nueva, resumen = agregar_pendientes(
                        conciliacion, rutas_xml, Path(tmp), avisar=st.write,
                        progreso_sat=lambda hechas, total: barra.progress(hechas / total,
                                                                          text=f"Consulta al SAT: {hechas}/{total}"))
                st.session_state["conciliacion"] = nueva
                st.session_state["resumen_pendientes"] = resumen
                st.session_state["pendientes_ver"] += 1          # el cargador queda vacío para la siguiente vuelta
                guardar()
                estado.update(label="Listo", state="complete", expanded=False)
            except ErrorUsuario as e:
                estado.update(label="No se agregó nada", state="error")
                st.error(str(e))
            except Exception as e:                               # noqa: BLE001 - se muestra al usuario
                estado.update(label="Error inesperado", state="error")
                st.error(f"Error inesperado: {type(e).__name__}: {e}\n\n"
                         "Toma captura de este mensaje y compártela con quien mantiene el programa.")
        if st.session_state.get("conciliacion") is not conciliacion:
            st.rerun()                                           # Excel y números con lo nuevo

st.divider()
st.subheader("Confirmar conciliación")
st.write(f"Registra en el historial los UUID que cruzaron en {periodo}. Así, si una de estas facturas vuelve "
         "a aparecer en otro mes, el programa la detiene como posible pago doble. "
         "Si ya se había confirmado este periodo, se reemplaza.")
if conciliacion.sin_consultar:
    st.warning(f"No se puede confirmar: en esta corrida no se pudo consultar {' ni '.join(conciliacion.sin_consultar)}. "
               "Vuelve a procesar con SAP abierto y conexión a internet.")
if st.button("Confirmar y registrar en el historial", disabled=not conciliacion.se_puede_registrar):
    mensaje = registrar_historial(conciliacion)
    (st.warning if "no se actualizó" in mensaje.lower() else st.success)(mensaje)
