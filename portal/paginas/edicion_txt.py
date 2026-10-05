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
from honorarios.config import ARCHIVO_CATALOGO, CODIGOS_IMPUESTO_TXT, CONFIG
from honorarios.entradas import NOMBRES_MES, leer_catalogo
from honorarios.servicio import consultar_sap_faltantes, ids_sin_consultar_en_sap, revalidar
from honorarios.utilidades import ErrorUsuario, guardar_seguro, importe_txt, normalizar_id
from portal import filtros
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
    ss.pop("filtros", None)                               # otro TXT: la tabla empieza sin filtros
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
        referencia = txt_sap.referencia_con_folio(e["lineas"][n][10], actuales[n], folio)
        registrar_cambio(n, "Folio", folio, "Formatear folios")
        registrar_cambio(n, "Referencia", referencia, "Formatear folios")
        e["claves"][n] = clave
    despues_de_cambiar()


@st.dialog("Editar línea", width="large")
def editar_linea(n, apoyo):
    """Ventana para corregir a mano IDProveedor, Tax y Monto de una línea. La sugerencia del portal es solo apoyo:
    la ventana abre con los valores ACTUALES y quien edita decide (SAP le paga a la cuenta del ID del TXT)."""
    valores = dict(zip(txt_sap.COLUMNAS, e["lineas"][n]))
    st.markdown(f"**Línea {n}** · Fecha {valores['Fecha']} · Folio {valores['Folio']}")
    st.info(f"**Apoyo del portal:** {apoyo}", icon="💡")
    id_nuevo = st.text_input("IDProveedor", value=valores["IDProveedor"].strip(), key=f"dlg_id_{n}")
    nombre = nombres.get(normalizar_id(id_nuevo))
    st.text_input("Nombre (del catálogo, según el ID)", value=nombre or "Este ID no está en el catálogo",
                  disabled=True, key=f"dlg_nombre_{n}_{id_nuevo}")
    col_tax, col_monto = st.columns(2)
    opciones = sorted(CODIGOS_IMPUESTO_TXT)
    actual_tax = valores["Tax"].strip().upper()
    tax = col_tax.selectbox("Tax", opciones, index=opciones.index(actual_tax) if actual_tax in opciones else 0,
                            key=f"dlg_tax_{n}", help="E0: sin IVA (médicos) · F2: con IVA (otros profesionales)")
    monto = col_monto.text_input("Monto a pagar", value=valores["Monto"].strip(), key=f"dlg_monto_{n}")
    monto_valido = importe_txt(monto) is not None
    if not monto_valido:
        st.error("El monto no es válido (formato 1234.56 o 1,234.56, máximo 2 decimales).")
    st.caption("SAP le deposita a la cuenta bancaria del ID que va en el TXT: revisa en SAP (FK03) a quién "
               "corresponde el ID antes de guardar. El cambio queda en la bitácora con tu usuario.")
    guardar_col, cancelar_col = st.columns(2)
    if guardar_col.button("Guardar cambios", type="primary", width="stretch", disabled=not monto_valido):
        hubo = False
        for columna, valor in (("IDProveedor", id_nuevo.strip()), ("Tax", tax), ("Monto", monto.strip())):
            hubo |= registrar_cambio(n, columna, valor, "Edición manual (ventana Editar línea)")
        if hubo:
            despues_de_cambiar()
        st.rerun()
    if cancelar_col.button("Cancelar", width="stretch"):
        st.rerun()


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
                  "Motivo del estatus": r.get("MotivoRechazo") or r.get("Observaciones", ""),
                  "_editada": any(linea == n for linea, _ in editadas),
                  "_sin_folio": not folios.es_folio_generado(valores[2])})
tabla = pd.DataFrame(filas, columns=AUXILIARES + txt_sap.COLUMNAS + NOTAS + ["_editada", "_sin_folio"])

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
editar = b[0].toggle("Editar", help="Activa la edición de IDProveedor, Monto, Tax (E0/F2), Centro de Costo y "
                                    "Cuenta Contable. Cada cambio queda en la bitácora.")
b[1].button("Formatear folios", on_click=formatear_folios, disabled=not conciliacion,
            help="Pone a cada línea un folio único de 10 caracteres, calculado a partir del UUID de su factura "
                 "(misma factura = mismo folio; nunca se repite). Las líneas sin factura reciben uno aleatorio. "
                 "La Referencia cambia igual: folio nuevo y el resto unidos con guion bajo, sin espacios "
                 "(\"A1893      Q18-2026\" -> \"7KQ2M9XHRT_Q18-2026\"). "
                 + ("" if conciliacion else "Disponible después de procesar el mes en Conciliador Automático."))

# Filtros y orden: solo cambian lo que se VE (el TXT descargado siempre lleva todas las líneas).
# Los valores viven en ss.filtros, no en los controles: cuando un botón los cambia (Pendientes de revisar, Para
# corregir, ✕ de un filtro, Quitar todos), sube ss.filtros_ver y los controles del panel se crean de nuevo con otra
# clave. Si se guardaran en los controles, uno que estuvo abierto y luego cerrado se quedaba con su valor viejo y,
# en la siguiente acción, lo volvía a poner (un filtro ya quitado "revivía").
ss.setdefault("filtros", {c: list(v) if isinstance(v, list) else v for c, v in filtros.VACIOS.items()})
ss.setdefault("filtros_ver", 0)


def poner_filtros(valores):
    ss.filtros.update({c: list(v) if isinstance(v, list) else v for c, v in valores.items()})
    ss.filtros_ver += 1                                  # controles nuevos con los valores nuevos


def tomar_filtro(clave):
    """Un control del panel cambió: su valor pasa a ss.filtros."""
    valor = ss[f"{clave}_{ss.filtros_ver}"]
    ss.filtros[clave] = list(valor) if isinstance(valor, list) else valor


def control(clave):
    return {"key": f"{clave}_{ss.filtros_ver}", "on_change": tomar_filtro, "args": (clave,)}


fi = ss.filtros
vista, activos = filtros.aplicar(tabla, fi)
f = st.columns([1.3, 1.3, 1.1, 2.3])
with f[0].popover(f"Filtros y orden{f' ({len(activos)})' if activos else ''}", icon=":material/filter_alt:",
                  width="stretch"):
    st.markdown("**Estatus**")
    st.segmented_control("Modo", filtros.MODOS, default=fi["f_modo"], required=True, label_visibility="collapsed",
                         **control("f_modo"))
    # Siempre TODOS los estatus (con cuántas líneas tiene cada uno): si la lista dependiera de lo que hay en la
    # tabla, un filtro con un estatus que no aparece ese mes quedaría activo pero invisible y no se podría quitar.
    conteo = tabla["Estatus"].value_counts()
    st.multiselect("Estatus", filtros.GRAVEDAD, default=fi["f_estatus"], placeholder="Todos",
                   label_visibility="collapsed", format_func=lambda e: f"{e} ({int(conteo.get(e, 0))})",
                   **control("f_estatus"))
    st.multiselect("Solo líneas con", list(filtros.CON), default=fi["f_con"], placeholder="Cualquiera",
                   help="Muestra las líneas que tengan al menos una de las opciones elegidas.", **control("f_con"))
    st.segmented_control("Tax", filtros.TAX, default=fi["f_tax"], required=True, **control("f_tax"))
    m = st.columns(2)
    m[0].number_input("Monto desde $", min_value=0.0, step=1000.0, value=fi["f_min"], placeholder="Sin mínimo",
                      **control("f_min"))
    m[1].number_input("Monto hasta $", min_value=0.0, step=1000.0, value=fi["f_max"], placeholder="Sin máximo",
                      **control("f_max"))
    st.selectbox("Ordenar por", filtros.ORDENES, index=filtros.ORDENES.index(fi["f_orden"]),
                 help="También puedes ordenar con clic en el encabezado de cualquier columna.", **control("f_orden"))
f[1].button("Pendientes de revisar", icon=":material/pending_actions:", width="stretch",
            on_click=poner_filtros, args=(filtros.PENDIENTES,),
            help="Oculta las líneas Cruzado y ordena de la más grave a la menos grave.")
f[2].button("Para corregir", icon=":material/build:", width="stretch", on_click=poner_filtros,
            args=(filtros.PARA_CORREGIR,), help="Solo líneas con problemas en el TXT o con posible ID equivocado.")
f[3].text_input("Buscar", value=fi["f_buscar"], placeholder="Buscar médico, ID, folio o referencia",
                label_visibility="collapsed", **control("f_buscar"))
# Filtros activos, siempre a la vista: cada uno con su ✕ para quitar solo ese, y "Quitar todos los filtros"
with st.container(horizontal=True, vertical_alignment="center", gap="small"):
    st.markdown(f"Mostrando **{len(vista):,} de {len(tabla):,}** líneas" if activos
                else f"Mostrando **todas las líneas ({len(tabla):,})**, sin filtros", width="content")
    for texto, quitar in activos:
        st.button(texto, icon=":material/close:", key="quitar_" + "_".join(sorted(quitar)), on_click=poner_filtros,
                  args=(quitar,), help="Quitar solo este filtro")
    st.button("Quitar todos los filtros", icon=":material/filter_alt_off:", key="quitar_todos", type="primary",
              on_click=poner_filtros, args=(filtros.VACIOS,), disabled=not activos)
vista = vista.drop(columns=["_editada", "_sin_folio"])
ss.etxt_vista = vista["Línea"].tolist()

if editar:
    st.data_editor(
        vista, key=f"editor_{e['version']}", on_change=aplicar_ediciones, hide_index=True,
        width="stretch", num_rows="fixed",
        disabled=[c for c in vista.columns if c not in txt_sap.EDITABLES],
        column_config={**{c: st.column_config.TextColumn(c) for c in txt_sap.COLUMNAS},
                       "Tax": st.column_config.SelectboxColumn(
                           "Tax", options=sorted(CODIGOS_IMPUESTO_TXT),
                           help="E0: sin IVA (médicos) · F2: con IVA (otros profesionales). Cámbialo si el médico "
                                "facturó con otro esquema de impuestos.")})
    st.caption("Las celdas grises no se editan: Fecha, X, 1 y PM son campos fijos de SAP; el Folio y la Referencia "
               "se cambian juntos con \"Formatear folios\". Al editar el ID, el médico y el estatus se actualizan "
               "solos.")
else:
    def colorear(fila):
        n = fila["Línea"]
        return [celdas.get((n, c)) or (AMARILLO if (n, c) in editadas else "") for c in fila.index]
    st.dataframe(vista.style.apply(colorear, axis=1), hide_index=True, width="stretch")
    st.caption("🟥 problema en esa celda · 🟨 celda editada. Activa **Editar** para corregir.")

pendientes = {n: s for n, s in sugerencias.items() if n in set(ss.etxt_vista)}
if pendientes:
    with st.expander(f"Posibles ID equivocados ({len(pendientes)}) · solo apoyo", expanded=True):
        st.caption("El portal no cambia nada solo. Revisa cada caso en SAP y, si hay que corregir, usa "
                   "**Editar línea**.")
        for n, (_, texto) in sorted(pendientes.items()):
            col = st.columns([5, 1])
            col[0].markdown(f"**Línea {n}:** {texto}")
            if col[1].button("Editar línea", key=f"sugerencia_{n}"):
                editar_linea(n, texto)

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
    st.caption(f"Lleva **todas las líneas ({len(e['lineas']):,})**, sin importar los filtros de la tabla. "
               "Mismo formato que el TXT original: 11 columnas separadas por tabulador, sin encabezados. "
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
