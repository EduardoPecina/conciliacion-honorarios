"""Programa de consola (.bat): pregunta el periodo, corre la conciliación y registra el historial.
La conciliación en sí está en servicio.py, la misma que usa el portal web."""

from .config import CP_SOCIEDAD, NOMBRE_SOCIEDAD
from .entradas import obtener_rutas, pedir_periodo
from .sap import consultar_sap
from .servicio import conciliar_periodo, registrar_historial


# ------------------------------------------------------------------
# PRINCIPAL
# ------------------------------------------------------------------
def si(pregunta, defecto):
    texto_defecto = "sí" if defecto == "s" else "no"
    respuesta = input(f"{pregunta} (s = sí, n = no; Enter = {texto_defecto}): ").strip().lower() or defecto
    return respuesta in ("s", "si", "sí", "y", "yes")


def main():
    anio, mes = pedir_periodo()
    faltan_config = [n for n, v in (("NOMBRE_SOCIEDAD", NOMBRE_SOCIEDAD), ("CP_SOCIEDAD", CP_SOCIEDAD)) if not v]
    if faltan_config:
        print(f"\nAviso: {', '.join(faltan_config)} vacío(s) en la configuración; esa validación no se hace.")

    print("\nRevisando carpetas...")
    rutas = obtener_rutas(anio, mes)

    def obtener_sap(catalogo, lineas):
        """SAP GUI solo existe en la computadora de quien corre el programa: se pregunta aquí."""
        if not si("¿Validar contra SAP? SAP debe estar abierto", "n"):
            return None
        en_catalogo = set(catalogo["IDProveedor"])
        ids = sorted({l["IDProveedor"] for l in lineas if l["IDProveedor"] in en_catalogo})
        cache = rutas["salida"] / f"CodigosSAP_{anio}-{mes:02d}.xlsx"
        reusar = not cache.exists() or si("Hay consultas SAP guardadas de este mes. ¿Reusarlas? "
                                          "(di 'n' si corregiste algo en SAP)", "s")
        print(f"Consultando {len(ids)} proveedores en SAP (FAVOR DE NO TOCAR EL MOUSE NI EL TECLADO)...")
        return consultar_sap(ids, rutas["salida"], anio, mes, reusar)

    conciliacion = conciliar_periodo(rutas, anio, mes, obtener_sap=obtener_sap)
    print("\n" + conciliacion.resumen.to_string(index=False))
    print(f"\nReporte guardado en:\n  {conciliacion.destino}")
    print(registrar_historial(conciliacion))
