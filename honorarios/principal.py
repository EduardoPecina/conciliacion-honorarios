"""Programa de consola (.bat): pregunta el periodo, pide abrir SAP, corre la conciliación y registra el historial.
La conciliación en sí está en servicio.py, la misma que usa el portal web."""

from .config import CP_SOCIEDAD, NOMBRE_SOCIEDAD
from .entradas import obtener_rutas, pedir_periodo
from .servicio import AYUDA_REUSAR_SAP, conciliar_periodo, fecha_consultas_sap, registrar_historial


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

    print("\nAntes de continuar (la validación contra SAP es obligatoria):")
    print("  - SAP abierto en el sistema productivo, con tu sesión iniciada")
    print("  - Acceso a la transacción FK03 y scripting de SAP GUI habilitado (los otorga IT)")
    print("Se usará la primera ventana de SAP que abriste: guarda lo que tengas ahí.")
    print("No uses SAP hasta que termine; puedes seguir trabajando en otros programas.")
    input("Cuando SAP esté listo, presiona Enter para continuar...")
    fecha_sap = fecha_consultas_sap(rutas, anio, mes)
    reusar_sap = False
    if fecha_sap:
        dia, hora = fecha_sap.split()
        print(f"\n{AYUDA_REUSAR_SAP}")
        reusar_sap = si(f"¿Reusar lo consultado a SAP el {dia} a las {hora}? (más rápido)", "n")

    conciliacion = conciliar_periodo(rutas, anio, mes, reusar_sap=reusar_sap)
    print("\n" + conciliacion.resumen.to_string(index=False))
    print(f"\nReporte guardado en:\n  {conciliacion.destino}")
    print(registrar_historial(conciliacion))
