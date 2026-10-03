"""Página Edición TXT: revisar, corregir y descargar el TXT que se carga a SAP.

El portal NO decide nada aquí: lee y escribe el TXT con honorarios/txt_sap.py, vuelve a cruzar con
honorarios/servicio.py (mismas reglas que la conciliación) y genera folios con honorarios/folios.py."""

import getpass
import hashlib
import json
from dataclasses import replace
from datetime import datetime

import pandas as pd
import streamlit as st

from honorarios import folios, txt_sap
from honorarios.config import ARCHIVO_CATALOGO, CONFIG
from honorarios.entradas import NOMBRES_MES, leer_catalogo
from honorarios.servicio import consultar_sap_faltantes, ids_sin_consultar_en_sap, revalidar
from honorarios.utilidades import ErrorUsuario, guardar_seguro, normalizar_id
from portal.trabajo import guardar, ofrecer_continuar

ss = st.session_state
ROJO, AMARILLO = "background-color: rgba(220, 38, 38, 0.20)", "background-color: rgba(234, 179, 8, 0.28)"
AUXILIARES = ["Línea", "Estatus", "Médico"]
NOTAS = ["Problemas del TXT", "Sugerencia", "Motivo del estatus"]

st.title("Edición TXT")
st.caption("Revisa y corrige el TXT antes de cargarlo a SAP. Las columnas **Línea, Estatus, Médico** y las de la "
           "derecha son solo de apoyo: no van en el TXT descargado.")

# ------------------------------------------------------------------ 1. ¿Qué TXT?
conciliacion = ss.get("conciliacion")
original = ss.get("txt_original")
if conciliacion is None or original is None:
    ofrecer_continuar()
    st.info("Para ver el **estatus de cada línea**, primero procesa el mes en **Conciliador Automático**. "
            "Mientras tanto, puedes subir un TXT aquí para revisar y corregir solo su formato.")
    subido = st.file_uploader("TXT de pagos", type=["txt"])
    if subido is None:
        st.stop()
    conciliacion, original, salida = None, (subido.name, subido.getvalue()), None
else:
    salida = ss.get("salida")
nombre_txt, datos = original
huella = hashlib.sha256(datos).hexdigest()
periodo = f"{conciliacion.anio}-{conciliacion.mes:02d}" if conciliacion else None
borrador = salida / f"EdicionTXT_{periodo}.json" if salida else None


@st.cache_data(show_spinner=False)
def catalogo_de_configuracion():
    try:
        return leer_catalogo(CONFIG / ARCHIVO_CATALOGO)
    except (ErrorUsuario, OSError):
        return None


catalogo = conciliacion.entradas["catalogo"] if conciliacion else catalogo_de_configuracion()
nombres = dict(zip(catalogo["IDProveedor"], catalogo["Nombre"])) if catalogo is not None else {}
ids_catalogo = set(nombres) if catalogo is not None else None


def guardar_borrador():
    if borrador is None:
        return
    e = ss.etxt
    contenido = {"huella": huella, "guardado": datetime.now().strftime("%d/%m/%Y %H:%M"),
                 "lineas": {str(n): v for n, v in e["lineas"].items()}, "bitacora": e["bitacora"],
                 "claves": {str(n): c for n, c in e["claves"].items()}}
    try:
        borrador.parent.mkdir(parents=True, exist_ok=True)
        guardar_seguro(borrador, lambda tmp: tmp.write_text(json.dumps(contenido, ensure_ascii=False),
                                                            encoding="utf-8"))
    except OSError:
        pass                                            # el borrador es una ayuda: si falla, se sigue trabajando


def volver_a_cruzar():
    """Mismas reglas que la conciliación, con lo ya consultado al SAT y SAP (segundos)."""
    if ss.get("conciliacion") is None or "txt_original" not in ss:
        return
    lineas, _ = txt_sap.lineas_para_el_motor(ss.etxt["lineas"])
    ss.conciliacion = replace(revalidar(ss.conciliacion, lineas), cambios=list(ss.etxt["bitacora"]),
                              txt_editado=txt_sap.escribir(ss.etxt["txt"], ss.etxt["lineas"]))
    guardar()                                           # el trabajo del mes queda guardado para otro día


if ss.get("etxt", {}).get("huella") != huella:          # TXT nuevo: se empieza de cero (o del borrador)
    txt = txt_sap.leer(datos)
    ss.etxt = {"huella": huella, "txt": txt, "lineas": {n: list(v) for n, v in txt.lineas.items()},
               "bitacora": [], "claves": {}, "version": 0}
    if borrador and borrador.exists():
        try:
            guardado = json.loads(borrador.read_text(encoding="utf-8"))
            if guardado.get("huella") == huella and guardado.get("bitacora"):
                ss.etxt.update(lineas={int(n): v for n, v in guardado["lineas"].items()},
                               bitacora=guardado["bitacora"],
                               claves={int(n): c for n, c in guardado.get("claves", {}).items()})
                ss.etxt_aviso = (f"Se recuperaron tus {len(guardado['bitacora'])} cambios guardados el "
                                 f"{guardado['guardado']}.")
                volver_a_cruzar()
        except (OSError, ValueError, KeyError):
            pass
e = ss.etxt
txt = e["txt"]
if ss.get("etxt_aviso"):
    st.success(ss.pop("etxt_aviso"))


def registrar_cambio(n, columna, nuevo, motivo="Edición manual"):
    i = txt_sap.COLUMNAS.index(columna)
    anterior = e["lineas"][n][i]
    if nuevo.strip() == anterior.strip():
        return False
    e["lineas"][n][i] = nuevo
    e["bitacora"].append({"Linea": n, "Columna": columna, "Anterior": anterior, "Nuevo": nuevo, "Motivo": motivo,
                          "Usuario": getpass.getuser(), "Fecha": datetime.now().strftime("%d/%m/%Y %H:%M:%S")})
    return True


def despues_de_cambiar():
    volver_a_cruzar()
    guardar_borrador()
    e["version"] += 1                                   # tabla nueva: el editor no repite los cambios ya aplicados


def aplicar_ediciones():
    cambios = ss.get(f"editor_{e['version']}", {}).get("edited_rows", {})
    hubo = False
    for posicion, columnas in cambios.items():
        n = ss.etxt_vista[int(posicion)]
        for columna, valor in columnas.items():
            if columna in txt_sap.EDITABLES:
                hubo |= registrar_cambio(n, columna, "" if valor is None else str(valor).strip())
    if hubo:
        despues_de_cambiar()


def formatear_folios():
    resultado = {f["Linea"]: f for f in ss.conciliacion.resultado} if ss.get("conciliacion") else {}
    lineas = [(n, tuple(u.strip() for u in str(resultado.get(n, {}).get("UUID", "")).split(",") if u.strip()))
              for n in e["lineas"] if n not in txt.estructura]
    actuales = {n: v[2] for n, v in e["lineas"].items()}
    for n, (folio, clave) in folios.asignar_folios(lineas, folios.cargar_folios(), actuales).items():
        registrar_cambio(n, "Folio", folio, "Formatear folios")
        e["claves"][n] = clave
    despues_de_cambiar()


def aplicar_sugerencia(n, id_nuevo):
    if registrar_cambio(n, "IDProveedor", id_nuevo, "Sugerencia del portal"):
        despues_de_cambiar()


def deshacer_todo():
    e["lineas"] = {n: list(v) for n, v in txt.lineas.items()}
    e["bitacora"], e["claves"] = [], {}
    if borrador and borrador.exists():
        borrador.unlink()
    volver_a_cruzar()
    e["version"] += 1


def registrar_descarga():
    """Al descargar: copia del TXT, bitácora de cambios y folios generados, en la carpeta del mes."""
    if salida is None:
        return
    marca = datetime.now().strftime("%Y%m%d_%H%M")
    salida.mkdir(parents=True, exist_ok=True)
    (salida / f"TXT_SAP_{periodo}_{marca}.txt").write_bytes(txt_sap.escribir(txt, e["lineas"]))
    bitacora = pd.DataFrame(e["bitacora"], columns=["Linea", "Columna", "Anterior", "Nuevo", "Motivo", "Usuario",
                                                     "Fecha"])
    guardar_seguro(salida / f"BitacoraTXT_{periodo}.csv",
                   lambda tmp: bitacora.to_csv(tmp, index=False, encoding="utf-8-sig"))
    folios.registrar_folios([
        {"Folio": e["lineas"][n][2], "Clave": clave, "IDProveedor": normalizar_id(e["lineas"][n][1]),
         "Periodo": periodo, "Linea": n}
        for n, clave in e["claves"].items() if e["lineas"][n][2] and folios.es_folio_generado(e["lineas"][n][2])])
    ss.etxt_aviso = f"Se guardó una copia del TXT y la bitácora en Reportes\\{conciliacion.anio}\\{salida.name}."


# ------------------------------------------------------------------ 2. Estado de cada línea
conciliacion = ss.get("conciliacion") if conciliacion else None
resultado = {f["Linea"]: f for f in conciliacion.resultado} if conciliacion else {}
_, errores_motor = txt_sap.lineas_para_el_motor(e["lineas"])
lista = txt_sap.problemas(txt, e["lineas"], ids_catalogo, periodo, folios.cargar_folios())
por_linea, celdas, bloquean = {}, {}, set()
for n, columna, texto, bloquea in lista:
    por_linea.setdefault(n, []).append(texto)
    celdas[(n, columna or "Fecha")] = ROJO
    if bloquea:
        bloquean.add(n)
sugerencias = (txt_sap.sugerencias_de_id(conciliacion.resultado, conciliacion.sobrantes, nombres)
               if conciliacion else {})
editadas = {(b["Linea"], b["Columna"]) for b in e["bitacora"]}

filas = []
for n, valores in sorted(e["lineas"].items()):
    r = resultado.get(n, {})
    estatus = ("Formato incorrecto" if n in errores_motor or n in txt.estructura
               else r.get("Estatus", "—") if conciliacion else "—")
    filas.append({"Línea": n, "Estatus": estatus, "Médico": nombres.get(normalizar_id(valores[1]), ""),
                  **dict(zip(txt_sap.COLUMNAS, valores)),
                  "Problemas del TXT": " · ".join(por_linea.get(n, [])),
                  "Sugerencia": sugerencias.get(n, ("", ""))[1],
                  "Motivo del estatus": r.get("MotivoRechazo") or r.get("Observaciones", "")})
tabla = pd.DataFrame(filas, columns=AUXILIARES + txt_sap.COLUMNAS + NOTAS)

c = st.columns(4)
c[0].metric("Líneas", f"{len(tabla):,}")
c[1].metric("Con problemas en el TXT", len(por_linea))
c[2].metric("Cambios hechos", len(e["bitacora"]))
c[3].metric("Folios por formatear", sum(1 for v in e["lineas"].values() if not folios.es_folio_generado(v[2])))
if conciliacion and conciliacion.editada:
    st.caption("El estatus ya considera tus cambios (se volvió a cruzar con lo consultado al SAT y a SAP).")

# ------------------------------------------------------------------ 3. Acciones
faltan_sap = ids_sin_consultar_en_sap(conciliacion) if conciliacion else []
if faltan_sap:
    st.warning(f"{len(faltan_sap)} ID(s) no se han consultado en SAP ({', '.join(faltan_sap[:10])}"
               f"{'…' if len(faltan_sap) > 10 else ''}): sus líneas no pueden quedar Cruzado hasta consultarlos. "
               "Abre SAP con tu sesión y da clic en el botón.")
    if st.button(f"Consultar en SAP {len(faltan_sap)} ID(s)"):
        with st.status("Consultando SAP...", expanded=True) as estado:
            barra = st.progress(0.0)
            nueva, texto = consultar_sap_faltantes(
                conciliacion, salida, progreso=lambda h, t, i: barra.progress(h / t, text=f"SAP: {h}/{t} ({i})"),
                avisar=st.write)
            estado.update(label=texto, state="complete")
        ss.conciliacion = replace(nueva, cambios=list(e["bitacora"]), txt_editado=txt_sap.escribir(txt, e["lineas"]))
        guardar()
        st.rerun()

b = st.columns([1, 1, 2])
editar = b[0].toggle("Editar", help="Activa la edición de IDProveedor, Monto, Tax, Centro de Costo, Cuenta "
                                    "Contable y Referencia. Cada cambio queda en la bitácora.")
b[1].button("Formatear folios", on_click=formatear_folios, disabled=not conciliacion,
            help="Pone a cada línea un folio único de 10 caracteres, calculado a partir del UUID de su factura "
                 "(misma factura = mismo folio; nunca se repite). Las líneas sin factura reciben uno aleatorio. "
                 + ("" if conciliacion else "Disponible después de procesar el mes en Conciliador Automático."))

f = st.columns([2, 1, 2])
estatus_elegidos = f[0].multiselect("Estatus", sorted(tabla["Estatus"].unique()), placeholder="Todos")
solo_problemas = f[1].checkbox("Solo con problemas")
buscar = f[2].text_input("Buscar (médico, ID, folio, referencia)")
vista = tabla
if estatus_elegidos:
    vista = vista[vista["Estatus"].isin(estatus_elegidos)]
if solo_problemas:
    vista = vista[(vista["Problemas del TXT"] != "") | (vista["Sugerencia"] != "")]
if buscar:
    texto = buscar.strip().lower()
    vista = vista[vista[["Médico", "IDProveedor", "Folio", "Referencia"]].apply(
        lambda fila: any(texto in str(v).lower() for v in fila), axis=1)]
ss.etxt_vista = vista["Línea"].tolist()

if editar:
    st.data_editor(
        vista, key=f"editor_{e['version']}", on_change=aplicar_ediciones, hide_index=True,
        width="stretch", num_rows="fixed",
        disabled=[c for c in vista.columns if c not in txt_sap.EDITABLES],
        column_config={c: st.column_config.TextColumn(c) for c in txt_sap.COLUMNAS})
    st.caption("Las celdas grises no se editan: Fecha, X, 1 y PM son campos fijos de SAP y el Folio se cambia con "
               "\"Formatear folios\". Al editar el ID, el médico y el estatus se actualizan solos.")
else:
    def colorear(fila):
        n = fila["Línea"]
        return [celdas.get((n, c)) or (AMARILLO if (n, c) in editadas else "") for c in fila.index]
    st.dataframe(vista.style.apply(colorear, axis=1), hide_index=True, width="stretch")
    st.caption("🟥 problema en esa celda · 🟨 celda editada. Activa **Editar** para corregir.")

pendientes = {n: s for n, s in sugerencias.items() if n in set(ss.etxt_vista)}
if pendientes:
    with st.expander(f"Sugerencias de ID ({len(pendientes)})", expanded=True):
        for n, (id_nuevo, texto) in sorted(pendientes.items()):
            col = st.columns([5, 1])
            col[0].markdown(f"**Línea {n}:** {texto}")
            col[1].button("Aplicar", key=f"sugerencia_{n}", on_click=aplicar_sugerencia, args=(n, id_nuevo))

# ------------------------------------------------------------------ 4. Descargar
st.subheader("Descargar TXT para SAP")
if bloquean:
    st.error(f"No se puede descargar todavía: {len(bloquean)} línea(s) tienen errores que SAP rechazaría "
             f"(líneas {', '.join(map(str, sorted(bloquean)[:15]))}). Corrígelas arriba.")
else:
    sin_folio = [n for n, v in e["lineas"].items() if not folios.es_folio_generado(v[2])]
    if sin_folio and conciliacion:
        st.warning(f"{len(sin_folio)} línea(s) conservan el folio original. Si SAP los rechaza por repetidos, "
                   "usa \"Formatear folios\" antes de descargar.")
    nombre = (f"TXT_SAP_{NOMBRES_MES[conciliacion.mes - 1]}_{conciliacion.anio}.txt" if conciliacion
              else nombre_txt.rsplit(".", 1)[0] + "_corregido.txt")
    st.download_button("Descargar TXT", txt_sap.escribir(txt, e["lineas"]), file_name=nombre, mime="text/plain",
                       type="primary", width="stretch", on_click=registrar_descarga)
    st.caption("Mismo formato que el TXT original: 11 columnas separadas por tabulador, sin encabezados. "
               + ("Al descargar se guarda una copia y la bitácora de cambios en la carpeta del mes."
                  if salida else "Sin conciliación no se guarda copia ni bitácora: procesa el mes para tenerlas."))

if conciliacion:
    st.page_link("paginas/conciliar.py", label="Ver el reporte actualizado (Excel con tus cambios) en Conciliador "
                 "Automático", icon="📊")

if e["bitacora"]:
    with st.expander(f"Bitácora de cambios ({len(e['bitacora'])})"):
        st.dataframe(pd.DataFrame(e["bitacora"]), hide_index=True, width="stretch")
        if st.checkbox("Quiero deshacer todos mis cambios y volver al TXT original"):
            st.button("Deshacer todos los cambios", on_click=deshacer_todo)
