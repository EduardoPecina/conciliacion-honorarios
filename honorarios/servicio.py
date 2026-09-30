"""Conciliación completa SIN preguntas en pantalla: la usan el programa de consola (.bat) y el portal web.

Aquí no hay reglas de negocio (esas viven en reglas.py y cruce.py): solo se encadenan los pasos
leer -> consultar SAT -> consultar SAP -> cruzar -> generar Excel, y se avisa el avance con funciones que
cada interfaz decide cómo mostrar (print en la consola, mensajes en el portal)."""

from dataclasses import dataclass, field

import pandas as pd

from .cruce import conciliar
from .entradas import leer_catalogo, leer_txt, leer_xmls
from .estado_sat import consultar_estados, progreso_en_consola
from .historial import cargar_historial, guardar_historial
from .reporte import generar_reporte
from .sap import consultar_sap, progreso_en_consola as progreso_sap_en_consola
from .utilidades import fecha_actual, validar_periodo


@dataclass
class Conciliacion:
    """Resultado de una corrida, listo para mostrarse o registrarse en el historial."""
    anio: int
    mes: int
    resultado: list                 # una fila por línea del TXT
    sobrantes: list                 # XML sin línea en el TXT
    errores_xml: list               # XML que no se pudieron leer
    estados: dict                   # estado de cada fuente (catálogo, historial, SAT, SAP) para la hoja Info
    estados_sat: dict               # UUID -> consulta al SAT
    destino: object                 # ruta del Excel generado
    resumen: pd.DataFrame           # líneas e importe por estatus (lo mismo que el Resumen del Excel)
    historial_legible: bool         # False: el historial no se pudo leer y no se debe registrar nada
    avisos: list = field(default_factory=list)
    sin_consultar: list = field(default_factory=list)   # "SAP" / "SAT" si no hubo conexión en esta corrida

    @property
    def se_puede_registrar(self):
        return self.historial_legible and not self.sin_consultar


def describir_catalogo(catalogo):
    tipos = catalogo["TipoRFC"].value_counts()
    return (f"{len(catalogo):,} proveedores: {tipos.get('MX', 0):,} con RFC mexicano, "
            f"{tipos.get('Extranjero', 0):,} extranjeros, {tipos.get('Inválido', 0):,} con RFC inválido, "
            f"{tipos.get('Vacío', 0):,} sin RFC")


def ruta_cache_sap(rutas, anio, mes):
    """Consultas a SAP guardadas del periodo: permiten no repetir la consulta si se vuelve a procesar."""
    return rutas["salida"] / f"CodigosSAP_{anio}-{mes:02d}.xlsx"


def conciliar_periodo(rutas, anio, mes, avisar=print, progreso_sat=progreso_en_consola,
                      progreso_sap=progreso_sap_en_consola, reusar_sap=True,
                      consultar_sat=consultar_estados, consultar_sap=consultar_sap, hoy=None):
    """Corre la conciliación con las rutas de entradas.rutas_en() / obtener_rutas().

    avisar(texto): muestra el avance.  progreso_sat(hechas, total) y progreso_sap(hechas, total, id): avance de
    las consultas. SAP se consulta siempre con la sesión de SAP GUI abierta en esta computadora; reusar_sap=False
    ignora las consultas guardadas del periodo (si se corrigió algo en SAP).
    NO registra el historial de UUID: eso lo decide quien llama, con registrar_historial().
    hoy: fecha de referencia (por omisión, la de internet); no se concilia un periodo que todavía no llega."""
    validar_periodo(anio, mes, hoy or fecha_actual()[0])
    avisos = []

    def aviso(texto):
        avisos.append(texto)
        avisar(texto)

    avisar(f"Leyendo catálogo ({rutas['catalogo'].name})...")
    catalogo = leer_catalogo(rutas["catalogo"])
    estados = {"Catálogo de proveedores": describir_catalogo(catalogo)}
    avisar(f"  {estados['Catálogo de proveedores']}")

    avisar(f"Leyendo TXT ({rutas['txt'].name})...")
    lineas = leer_txt(rutas["txt"])
    avisar(f"  {len(lineas):,} líneas")

    avisar(f"Leyendo {len(rutas['xmls']):,} XML...")
    facturas, errores_xml = leer_xmls(rutas["xmls"], rutas["carpeta_xml"])
    avisar(f"  {len(facturas):,} leídos, {len(errores_xml)} con error")

    historial, estados["Historial de UUID"] = cargar_historial()
    avisar(f"Historial de UUID: {estados['Historial de UUID']}")
    if historial is None:
        aviso("El historial de UUID no se pudo leer: no se detectarán pagos de meses anteriores "
              "y esta conciliación no se podrá registrar.")

    avisar(f"Consultando en el SAT {len({f['UUID'] for f in facturas if f['UUID']}):,} facturas...")
    estados_sat, estados["Estatus en el SAT"] = consultar_sat(facturas, progreso=progreso_sat, avisar=aviso)
    avisar(f"  {estados['Estatus en el SAT']}")

    en_catalogo = set(catalogo["IDProveedor"])
    ids = sorted({l["IDProveedor"] for l in lineas if l["IDProveedor"] in en_catalogo})
    avisar(f"Consultando {len(ids):,} proveedores en SAP (no uses SAP hasta que termine)...")
    sap, estados["Validación SAP"] = consultar_sap(ids, rutas["salida"], anio, mes, reusar_sap,
                                                   progreso=progreso_sap, avisar=aviso)
    avisar(f"  {estados['Validación SAP']}")

    avisar("Cruzando...")
    resultado, sobrantes = conciliar(lineas, catalogo, facturas, anio, mes, sap, historial, estados_sat)
    destino, resumen = generar_reporte(rutas, anio, mes, resultado, sobrantes,
                                       errores_xml, len(rutas["xmls"]), estados)
    sin_consultar = [nombre for nombre, falta in (
        ("SAP", sap is None),
        ("SAT", bool(estados_sat) and all(c.get("SinConexion") for c in estados_sat.values()))) if falta]
    return Conciliacion(anio, mes, resultado, sobrantes, errores_xml, estados, estados_sat,
                        destino, resumen, historial is not None, avisos, sin_consultar)


def registrar_historial(conciliacion):
    """Registra los UUID cruzados del periodo en el historial (reemplaza lo que hubiera de ese periodo)."""
    if not conciliacion.historial_legible:
        return "El historial de UUID no se actualizó porque no se pudo leer."
    if conciliacion.sin_consultar:
        return (f"El historial de UUID no se actualizó: no se pudo consultar {' ni '.join(conciliacion.sin_consultar)}. "
                "Vuelve a conciliar con conexión antes de registrar el periodo.")
    return guardar_historial(conciliacion.resultado, conciliacion.anio, conciliacion.mes)
