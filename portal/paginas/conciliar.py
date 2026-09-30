"""Página Conciliar: subir TXT y XML -> Procesar -> resumen -> Descargar Excel -> Confirmar."""

import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from honorarios.config import MESES, RAIZ
from honorarios.entradas import NOMBRES_MES, rutas_en
from honorarios.estado_sat import alerta_efos, describir
from honorarios.reporte import ORDEN_ESTATUS, RESPONSABLE_POR_ESTATUS
from honorarios.servicio import conciliar_periodo, registrar_historial, ruta_cache_sap
from honorarios.utilidades import ErrorUsuario
from portal.archivos import preparar_carpeta

CRUZADOS = ("Cruzado", "Cruzado (pago quebrado)")

st.title("Conciliación de honorarios")

# ------------------------------------------------------------------ 1. Periodo y archivos
hoy = datetime.now()
col_anio, col_mes = st.columns(2)
anio = col_anio.selectbox("Año", list(range(hoy.year - 2, hoy.year + 1)), index=2)
mes = col_mes.selectbox("Mes", list(range(1, 13)), index=hoy.month - 1,
                        format_func=lambda m: NOMBRES_MES[m - 1].capitalize())

txt = st.file_uploader("1. TXT de pagos", type=["txt"])
xmls = st.file_uploader("2. Facturas XML (arrastra todos los archivos a la vez, o un .zip)",
                        type=["xml", "zip"], accept_multiple_files=True)
if xmls:
    st.caption(f"{len(xmls):,} archivo(s) cargado(s)")

salida = RAIZ / str(anio) / MESES[mes - 1] / "Salida"      # el Excel se guarda donde lo deja el .bat
st.info("**Antes de procesar, abre SAP e inicia sesión.** La validación contra SAP es obligatoria. "
        "Mientras se consulta SAP, **no toques el mouse ni el teclado**.")
reusar_sap = True
if ruta_cache_sap({"salida": salida}, anio, mes).exists():
    reusar_sap = st.checkbox("Reusar las consultas a SAP guardadas de este mes "
                             "(desmárcalo si corregiste algo en SAP)", value=True)

procesar = st.button("Procesar", type="primary", disabled=not (txt and xmls), use_container_width=True)

# ------------------------------------------------------------------ 2. Proceso
if procesar:
    st.session_state.pop("conciliacion", None)
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
                        hechas / total, text=f"Consulta a SAP: {hechas}/{total} (proveedor {id_prov})"))
            st.session_state["conciliacion"] = conciliacion
            st.session_state["excel"] = Path(conciliacion.destino).read_bytes()
            estado.update(label="Listo", state="complete", expanded=False)
        except ErrorUsuario as e:
            estado.update(label="No se pudo completar", state="error")
            st.error(str(e))
        except Exception as e:                                   # noqa: BLE001 - se muestra al usuario
            estado.update(label="Error inesperado", state="error")
            st.error(f"Error inesperado: {type(e).__name__}: {e}\n\n"
                     "Toma captura de este mensaje y compártela con quien mantiene el programa.")

# ------------------------------------------------------------------ 3. Resultado
conciliacion = st.session_state.get("conciliacion")
if conciliacion is None:
    st.stop()

periodo = f"{NOMBRES_MES[conciliacion.mes - 1]} {conciliacion.anio}"
st.header(f"Resultado · {periodo}")
for aviso in conciliacion.avisos:
    st.warning(aviso)

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
    st.dataframe(pd.DataFrame(pendientes), hide_index=True, use_container_width=True,
                 column_config={"Importe": st.column_config.NumberColumn(format="$%,.2f")})
else:
    st.success("Todas las líneas cruzaron sin observaciones.")
if conciliacion.sobrantes:
    st.caption(f"{len(conciliacion.sobrantes):,} XML sin línea en el TXT (hoja 'XML sin TXT' del Excel). "
               f"{len(conciliacion.errores_xml):,} XML no se pudieron leer (hoja 'Errores XML').")

st.download_button("Descargar Excel", st.session_state["excel"], file_name=Path(conciliacion.destino).name,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   type="primary", use_container_width=True)

st.divider()
st.subheader("Confirmar conciliación")
st.write(f"Registra en el historial los UUID que cruzaron en {periodo}. Así, si una de estas facturas vuelve "
         "a aparecer en otro mes, el programa la detiene como posible pago doble. "
         "Si ya se había confirmado este periodo, se reemplaza.")
if conciliacion.sin_consultar:
    st.warning(f"No se puede confirmar: en esta corrida no se pudo consultar {' ni '.join(conciliacion.sin_consultar)}. "
               "Vuelve a procesar con SAP abierto y conexión a internet.")
if st.button("Confirmar y registrar en el historial", disabled=not conciliacion.se_puede_registrar):
    st.success(registrar_historial(conciliacion))
