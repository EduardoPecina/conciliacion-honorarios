"""Flujo principal del programa: pregunta el periodo, lee entradas, cruza y genera el reporte."""

from .config import CP_SOCIEDAD, NOMBRE_SOCIEDAD
from .entradas import leer_catalogo, leer_txt, leer_xmls, obtener_rutas, pedir_periodo
from .historial import cargar_historial, guardar_historial
from .cruce import conciliar
from .estado_sat import consultar_estados
from .reporte import generar_reporte
from .sap import consultar_sap


# ------------------------------------------------------------------
# PRINCIPAL
# ------------------------------------------------------------------
def si(pregunta, defecto):
    respuesta = input(f"{pregunta} (s/n) [{defecto}]: ").strip().lower() or defecto
    return respuesta in ("s", "si", "sí", "y", "yes")


def main():
    anio, mes = pedir_periodo()
    faltan_config = [n for n, v in (("NOMBRE_SOCIEDAD", NOMBRE_SOCIEDAD), ("CP_SOCIEDAD", CP_SOCIEDAD)) if not v]
    if faltan_config:
        print(f"\nAviso: {', '.join(faltan_config)} vacío(s) en la configuración; esa validación no se hace.")

    print("\nRevisando carpetas...")
    rutas = obtener_rutas(anio, mes)

    print(f"Leyendo catálogo ({rutas['catalogo'].name})...")
    catalogo = leer_catalogo(rutas["catalogo"])
    tipos = catalogo["TipoRFC"].value_counts()
    resumen_cat = (f"{len(catalogo):,} proveedores: {tipos.get('MX', 0):,} con RFC mexicano, "
                   f"{tipos.get('Extranjero', 0):,} extranjeros, {tipos.get('Inválido', 0):,} con RFC inválido, "
                   f"{tipos.get('Vacío', 0):,} sin RFC")
    print(f"  {resumen_cat}")

    print(f"Leyendo TXT ({rutas['txt'].name})...")
    lineas = leer_txt(rutas["txt"])
    print(f"  {len(lineas):,} líneas")

    print(f"Leyendo {len(rutas['xmls']):,} XML...")
    facturas, errores_xml = leer_xmls(rutas["xmls"], rutas["carpeta_xml"])
    print(f"  {len(facturas):,} leídos, {len(errores_xml)} con error")

    estados = {"Catálogo de proveedores": resumen_cat}

    historial, estados["Historial de UUID"] = cargar_historial()
    print(f"Historial de UUID: {estados['Historial de UUID']}")

    print(f"Consultando en el SAT {len({f['UUID'] for f in facturas if f['UUID']}):,} facturas (requiere internet)...")
    estados_sat, estados["Estatus en el SAT"] = consultar_estados(facturas)
    print(f"  {estados['Estatus en el SAT']}")

    sap, estados["Validación SAP"] = None, "No se consultó"
    if si("¿Validar contra SAP? SAP debe estar abierto", "n"):
        en_catalogo = set(catalogo["IDProveedor"])
        ids = sorted({l["IDProveedor"] for l in lineas if l["IDProveedor"] in en_catalogo})
        cache = rutas["salida"] / f"CodigosSAP_{anio}-{mes:02d}.xlsx"
        reusar = not cache.exists() or si("Hay consultas SAP guardadas de este mes. ¿Reusarlas? "
                                          "(di 'n' si corregiste algo en SAP)", "s")
        print(f"Consultando {len(ids)} proveedores en SAP (FAVOR DE NO TOCAR EL MOUSE NI EL TECLADO)...")
        sap, estados["Validación SAP"] = consultar_sap(ids, rutas["salida"], anio, mes, reusar)

    print("Cruzando...")
    resultado, sobrantes = conciliar(lineas, catalogo, facturas, anio, mes, sap, historial, estados_sat)

    destino, resumen = generar_reporte(rutas, anio, mes, resultado, sobrantes,
                                       errores_xml, len(rutas["xmls"]), estados)
    print("\n" + resumen.to_string(index=False))
    print(f"\nReporte guardado en:\n  {destino}")

    if historial is not None:
        print(guardar_historial(resultado, anio, mes))
    else:
        print("El historial de UUID no se actualizó porque no se pudo leer.")
