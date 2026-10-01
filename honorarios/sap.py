"""Consulta de códigos de retención en SAP GUI (transacción FK03 -> BP, tabla LFBW). Es obligatoria:
se hace en cada corrida con la sesión de SAP abierta en la computadora de quien concilia."""

from datetime import datetime

import pandas as pd

from .config import SOCIEDAD_SAP


# ------------------------------------------------------------------
# 5. SAP (códigos de retención por proveedor, transacción FK03)
# ------------------------------------------------------------------
NOMBRE_TABLA_LFBW = "tblSAPLCVI_FS_UI_VENDOR_CCTCTRL_LFBW"   # tabla de retenciones en BP
CASILLA_RETENCION = "chkRF02K-D0610"


def buscar_controles(nodo, condicion, encontrados=None, nivel=0):
    """Recorre la pantalla de SAP y devuelve los controles que cumplen la condición."""
    if encontrados is None:
        encontrados = []
    try:
        hijos = nodo.Children
    except Exception:
        return encontrados
    for i in range(hijos.Count):
        h = hijos.ElementAt(i)
        if condicion(h):
            encontrados.append(h)
        if nivel < 25:
            buscar_controles(h, condicion, encontrados, nivel + 1)
    return encontrados


def conectar_sap():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()     # necesario fuera del hilo principal (el portal corre en otro hilo)
    return win32com.client.GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)


def cerrar_ventanas(session):
    """Cierra ventanas emergentes (wnd[1], wnd[2]...) que bloquean la pantalla principal."""
    for _ in range(5):
        try:
            session.findById("wnd[1]").close()
        except Exception:
            return


def enter(session):
    """Presiona Enter y confirma avisos amarillos (tipo W), que SAP detiene hasta otro Enter."""
    session.findById("wnd[0]").sendVKey(0)
    for _ in range(3):
        if session.findById("wnd[0]/sbar").MessageType != "W":
            break
        session.findById("wnd[0]").sendVKey(0)


def codigos_retencion(session, id_proveedor):
    paso = "abrir FK03"
    try:
        cerrar_ventanas(session)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nFK03"
        session.findById("wnd[0]").sendVKey(0)
        cerrar_ventanas(session)

        paso = "marcar solo la casilla de retenciones"
        usr = session.findById("wnd[0]/usr")
        for casilla in buscar_controles(usr, lambda c: c.Type == "GuiCheckBox"):
            casilla.Selected = casilla.Id.endswith(CASILLA_RETENCION)

        paso = "llenar proveedor y sociedad"
        session.findById("wnd[0]/usr/ctxtRF02K-LIFNR").text = id_proveedor
        session.findById("wnd[0]/usr/ctxtRF02K-BUKRS").text = SOCIEDAD_SAP
        enter(session)

        barra = session.findById("wnd[0]/sbar")
        if barra.MessageType in ("E", "A"):
            return [], barra.Text
        cerrar_ventanas(session)

        paso = "leer tabla de retenciones"
        tablas = buscar_controles(session.findById("wnd[0]/usr"),
                                  lambda c: c.Id.endswith(NOMBRE_TABLA_LFBW))
        if not tablas:
            return [], f"No apareció la tabla de retenciones ({session.findById('wnd[0]').Text})"
        tabla = tablas[0]
        codigos = []
        for fila in range(tabla.VisibleRowCount):
            texto = tabla.GetCell(fila, 2).Text.strip()
            if not texto:
                break
            codigos.append(texto)
        else:
            if tabla.VisibleRowCount and tabla.RowCount > tabla.VisibleRowCount:
                # Hay renglones fuera de la vista: podría faltar IH, IP o el de ISR. Mejor no validar
                # que validar con datos incompletos.
                return [], (f"La tabla de retenciones tiene {tabla.RowCount} renglones y solo se ven "
                            f"{tabla.VisibleRowCount}: revisar en FK03 a mano")
        return codigos, ""
    except Exception as e:
        detalle = e.args[2][2] if len(e.args) > 2 and e.args[2] else str(e)
        return [], f"Falló en el paso '{paso}': {detalle}"


def progreso_en_consola(hechas, total, id_proveedor):
    print(f"  SAP {hechas}/{total}: proveedor {id_proveedor}   ", end="\n" if hechas == total else "\r")


FORMATO_CONSULTA = "%d/%m/%Y %H:%M"


def fecha_consultas_guardadas(cache_ruta):
    """Fecha de la consulta más antigua que se reusaría del archivo de consultas guardadas (texto "dd/mm/aaaa
    hh:mm"), o None si no hay archivo o no se puede leer. Es lo más viejo que se usaría al reusar."""
    try:
        previo = pd.read_excel(cache_ruta, dtype=str).fillna("")
        fechas = [datetime.strptime(f, FORMATO_CONSULTA) for f, e in zip(previo["Consulta"], previo["Error"])
                  if f and not e]
    except Exception:
        return None
    return min(fechas).strftime(FORMATO_CONSULTA) if fechas else None


def consultar_sap(ids, carpeta_salida, anio, mes, reusar=True, progreso=progreso_en_consola, avisar=print):
    """Consulta cada proveedor una sola vez y guarda el resultado para no repetir.
    Con reusar=False se ignora lo guardado (útil si se corrigió algo en SAP).
    Devuelve ({ID: (códigos, error)}, texto) o (None, texto) si no se pudo consultar nada."""
    cache_ruta = carpeta_salida / f"CodigosSAP_{anio}-{mes:02d}.xlsx"
    resultado = {}
    cuando = {}                                   # fecha de cada consulta: lo reusado conserva la suya
    if reusar and cache_ruta.exists():
        previo = pd.read_excel(cache_ruta, dtype=str).fillna("")
        for _, r in previo.iterrows():
            codigos = [c for c in r["Codigos"].split(",") if c]
            if not r["Error"] and r["IDProveedor"] in ids:
                resultado[r["IDProveedor"]] = (codigos, "")
                cuando[r["IDProveedor"]] = r.get("Consulta", "")

    faltan = [i for i in ids if i not in resultado]
    no_consultados = 0
    if faltan:
        try:
            session = conectar_sap()
        except Exception:
            avisar("SAP no está abierto (o el scripting de SAP GUI no está habilitado): ninguna línea quedará "
                   "'Cruzado'. Abre SAP, inicia sesión y vuelve a procesar.")
            if not resultado:
                return None, "No se pudo conectar a SAP"
            return resultado, (f"No se pudo conectar a SAP: se usaron {len(resultado)} consultas guardadas; "
                               f"{len(faltan)} proveedores sin validar")
        fallas_seguidas = 0
        for k, id_prov in enumerate(faltan, start=1):
            resultado[id_prov] = codigos_retencion(session, id_prov)
            cuando[id_prov] = datetime.now().strftime(FORMATO_CONSULTA)
            progreso(k, len(faltan), id_prov)
            fallas_seguidas = fallas_seguidas + 1 if resultado[id_prov][1] else 0
            if fallas_seguidas == 3:
                avisar(f"3 proveedores seguidos fallaron en SAP (último error: {resultado[id_prov][1]}). "
                       "Se detiene la consulta a SAP para no seguir fallando.")
                no_consultados = len(faltan) - k
                break

    try:
        pd.DataFrame([{"IDProveedor": i, "Codigos": ",".join(c), "Error": e,
                       "Consulta": cuando.get(i) or datetime.now().strftime(FORMATO_CONSULTA)}
                      for i, (c, e) in resultado.items()]).to_excel(cache_ruta, index=False)
    except PermissionError:
        avisar(f"No se pudo guardar {cache_ruta.name} (¿está abierto?).")
    texto = f"{len(resultado)} proveedores consultados (guardado en {cache_ruta.name})"
    if no_consultados:
        texto += f"; {no_consultados} sin consultar por fallas"
    return resultado, texto
