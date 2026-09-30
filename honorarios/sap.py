"""Consulta opcional de códigos de retención en SAP GUI (transacción FK03 -> BP, tabla LFBW)."""

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
    import win32com.client
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


def consultar_sap(ids, carpeta_salida, anio, mes, reusar=True):
    """Consulta cada proveedor una sola vez y guarda el resultado para no repetir.
    Con reusar=False se ignora lo guardado (útil si se corrigió algo en SAP)."""
    cache_ruta = carpeta_salida / f"CodigosSAP_{anio}-{mes:02d}.xlsx"
    resultado = {}
    if reusar and cache_ruta.exists():
        previo = pd.read_excel(cache_ruta, dtype=str).fillna("")
        for _, r in previo.iterrows():
            codigos = [c for c in r["Codigos"].split(",") if c]
            if not r["Error"] and r["IDProveedor"] in ids:
                resultado[r["IDProveedor"]] = (codigos, "")

    faltan = [i for i in ids if i not in resultado]
    no_consultados = 0
    if faltan:
        try:
            session = conectar_sap()
        except Exception:
            print("  SAP no está abierto o el scripting no está habilitado. Se omite la validación SAP.")
            if not resultado:
                return None, "No se pudo conectar a SAP"
            return resultado, (f"No se pudo conectar a SAP: se usaron {len(resultado)} consultas guardadas; "
                               f"{len(faltan)} proveedores sin validar")
        fallas_seguidas = 0
        for k, id_prov in enumerate(faltan, start=1):
            print(f"  SAP {k}/{len(faltan)}: proveedor {id_prov}   ", end="\r")
            resultado[id_prov] = codigos_retencion(session, id_prov)
            fallas_seguidas = fallas_seguidas + 1 if resultado[id_prov][1] else 0
            if fallas_seguidas == 3:
                print(f"\n  3 proveedores seguidos fallaron. Último error: {resultado[id_prov][1]}")
                print("  Se detiene la consulta a SAP para no seguir fallando.")
                no_consultados = len(faltan) - k
                break
        print()

    try:
        pd.DataFrame([{"IDProveedor": i, "Codigos": ",".join(c), "Error": e,
                       "Consulta": datetime.now().strftime("%d/%m/%Y %H:%M")}
                      for i, (c, e) in resultado.items()]).to_excel(cache_ruta, index=False)
    except PermissionError:
        print(f"  No se pudo guardar {cache_ruta.name} (¿está abierto?).")
    texto = f"{len(resultado)} proveedores consultados (guardado en {cache_ruta.name})"
    if no_consultados:
        texto += f"; {no_consultados} sin consultar por fallas"
    return resultado, texto
