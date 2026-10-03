"""Conciliación completa SIN preguntas en pantalla: la usan el programa de consola (.bat) y el portal web.

Aquí no hay reglas de negocio (esas viven en reglas.py y cruce.py): solo se encadenan los pasos
leer -> consultar SAT -> consultar SAP -> cruzar -> generar Excel, y se avisa el avance con funciones que
cada interfaz decide cómo mostrar (print en la consola, mensajes en el portal)."""

import pickle
from dataclasses import dataclass, field, fields, replace
from datetime import datetime
from pathlib import Path

import pandas as pd

from . import __version__
from .config import REPORTES
from .cruce import conciliar
from .entradas import NOMBRES_MES, leer_catalogo, leer_txt, leer_xmls
from .estado_sat import consultar_estados, progreso_en_consola
from .historial import cargar_historial, guardar_historial
from .reporte import generar_reporte
from .sap import consultar_sap, fecha_consultas_guardadas, progreso_en_consola as progreso_sap_en_consola
from .utilidades import ErrorUsuario, fecha_actual, guardar_seguro, validar_periodo


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
    # Lo leído y consultado en esta corrida (catálogo, líneas, facturas, SAP, historial): Edición TXT vuelve a
    # cruzar con esto al editar, sin repetir las consultas al SAT ni a SAP.
    entradas: dict = field(default_factory=dict, repr=False)
    editada: bool = False           # True: el resultado ya incluye cambios hechos en Edición TXT
    cambios: list = field(default_factory=list)     # bitácora de Edición TXT
    version: int = 0                # sube con cada cambio: el portal sabe cuándo regenerar el Excel
    consultado: str = ""            # fecha y hora de las consultas al SAT y a SAP ("dd/mm/aaaa hh:mm")
    txt_editado: bytes = b""        # el TXT con los cambios de Edición TXT (para el Excel y su huella)

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


def fecha_consultas_sap(rutas, anio, mes):
    """Fecha de las consultas a SAP guardadas del periodo ("dd/mm/aaaa hh:mm"), o None si no hay que reusar."""
    ruta = ruta_cache_sap(rutas, anio, mes)
    return fecha_consultas_guardadas(ruta) if ruta.exists() else None


AYUDA_REUSAR_SAP = ("Ya procesaste este mes en esta computadora. Márcalo solo si desde entonces NO se cambió nada en "
                    "SAP. Si se corrigió algún proveedor (p. ej. se le agregó la retención IH o pasó de RESICO a "
                    "honorarios, RS↔RH), déjalo sin marcar para consultar todo de nuevo.")


def revisar_periodo_de_los_archivos(lineas, facturas, anio, mes):
    """Regla: el TXT, los XML y el periodo elegido deben coincidir. Si la mayoría de las fechas (del TXT y de los
    XML) son de otro mes, se detiene ANTES de consultar SAT y SAP: lo más probable es que se eligió el mes
    equivocado. Si no, todas las facturas saldrían "Rechazar" por estar fuera del periodo, lo que confunde."""
    periodo = f"{anio}-{mes:02d}"
    fechas = [l["_fecha"].strftime("%Y-%m") for l in lineas] + [f["FechaXML"][:7] for f in facturas if f["FechaXML"]]
    if not fechas or fechas.count(periodo) * 2 >= len(fechas):
        return
    detectado = max(set(fechas), key=fechas.count)
    porcentaje = fechas.count(detectado) * 100 // len(fechas)
    nombre = lambda p: f"{NOMBRES_MES[int(p[5:]) - 1]} {p[:4]}"
    raise ErrorUsuario(f"Los archivos parecen ser de {nombre(detectado)} ({porcentaje}% de las fechas del TXT y de "
                       f"los XML) y elegiste {nombre(periodo)}. El TXT, los XML y el mes elegido deben coincidir: "
                       "revisa el mes o los archivos. No se consultó al SAT ni a SAP.")


def conciliar_periodo(rutas, anio, mes, avisar=print, progreso_sat=progreso_en_consola,
                      progreso_sap=progreso_sap_en_consola, reusar_sap=False,
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
    revisar_periodo_de_los_archivos(lineas, facturas, anio, mes)

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
    entradas = {"catalogo": catalogo, "lineas": lineas, "facturas": facturas, "sap": sap, "historial": historial,
                "ruta_catalogo": rutas["catalogo"], "n_xml": len(rutas["xmls"])}
    return Conciliacion(anio, mes, resultado, sobrantes, errores_xml, estados, estados_sat,
                        destino, resumen, historial is not None, avisos, sin_consultar, entradas,
                        consultado=datetime.now().strftime(FORMATO_FECHA))


def revalidar(conciliacion, lineas):
    """Vuelve a cruzar con las líneas editadas, usando lo ya consultado al SAT y a SAP (tarda segundos).
    Las mismas reglas que la conciliación: nada se salta."""
    e = conciliacion.entradas
    resultado, sobrantes = conciliar(lineas, e["catalogo"], e["facturas"], conciliacion.anio, conciliacion.mes,
                                     e["sap"], e["historial"], conciliacion.estados_sat)
    return replace(conciliacion, resultado=resultado, sobrantes=sobrantes, editada=True,
                   entradas={**e, "lineas": lineas}, version=conciliacion.version + 1)


def ids_sin_consultar_en_sap(conciliacion):
    """IDs del TXT (editado) que están en el catálogo pero no se consultaron en SAP (p. ej. un ID corregido)."""
    e = conciliacion.entradas
    en_catalogo = set(e["catalogo"]["IDProveedor"])
    return sorted({l["IDProveedor"] for l in e["lineas"] if l["IDProveedor"] in en_catalogo} - set(e["sap"] or {}))


def consultar_sap_faltantes(conciliacion, carpeta, progreso=progreso_sap_en_consola, avisar=print):
    """Consulta en SAP solo los IDs que faltan (reusa lo ya consultado este mes) y vuelve a cruzar."""
    e = conciliacion.entradas
    ids = sorted(set(e["sap"] or {}) | set(ids_sin_consultar_en_sap(conciliacion)))
    sap, texto = consultar_sap(ids, carpeta, conciliacion.anio, conciliacion.mes, True, progreso=progreso,
                               avisar=avisar)
    if sap is None:
        return conciliacion, texto
    actualizada = replace(conciliacion, entradas={**e, "sap": {**(e["sap"] or {}), **sap}},
                          sin_consultar=[x for x in conciliacion.sin_consultar if x != "SAP"])
    return revalidar(actualizada, e["lineas"]), texto


def registrar_historial(conciliacion):
    """Registra los UUID cruzados del periodo en el historial (reemplaza lo que hubiera de ese periodo)."""
    if not conciliacion.historial_legible:
        return "El historial de UUID no se actualizó porque no se pudo leer."
    if conciliacion.sin_consultar:
        return (f"El historial de UUID no se actualizó: no se pudo consultar {' ni '.join(conciliacion.sin_consultar)}. "
                "Vuelve a conciliar con conexión antes de registrar el periodo.")
    return guardar_historial(conciliacion.resultado, conciliacion.anio, conciliacion.mes)


# ------------------------------------------------------------------ Trabajo de varios días
FORMATO_FECHA = "%d/%m/%Y %H:%M"


def ruta_trabajo(carpeta, anio, mes):
    return Path(carpeta) / f"Trabajo_{anio}-{mes:02d}.pkl"


def guardar_trabajo(conciliacion, carpeta, txt_original):
    """Guarda todo lo del mes (resultado, lo consultado al SAT y a SAP, el TXT original, la bitácora) para poder
    continuar otro día sin volver a procesar. Se guarda en la carpeta del mes; si falla, no detiene nada."""
    datos = {"programa": __version__, "guardado": datetime.now().strftime(FORMATO_FECHA), "txt": txt_original,
             "conciliacion": {f.name: getattr(conciliacion, f.name) for f in fields(conciliacion)}}
    destino = ruta_trabajo(carpeta, conciliacion.anio, conciliacion.mes)
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        guardar_seguro(destino, lambda tmp: tmp.write_bytes(pickle.dumps(datos, protocol=pickle.HIGHEST_PROTOCOL)))
    except OSError:
        return False
    return True


def trabajos_guardados(raiz=REPORTES, limite=3):
    """Trabajos guardados más recientes: [(ruta, "septiembre 2026", "03/10/2026 18:40")]."""
    trabajos = sorted(Path(raiz).glob("*/*/Trabajo_*.pkl"), key=lambda r: r.stat().st_mtime, reverse=True)
    salida = []
    for ruta in trabajos[:limite]:
        anio, mes = ruta.stem.split("_")[1].split("-")
        salida.append((ruta, f"{NOMBRES_MES[int(mes) - 1]} {anio}",
                       datetime.fromtimestamp(ruta.stat().st_mtime).strftime(FORMATO_FECHA)))
    return salida


def cargar_trabajo(ruta):
    """-> (conciliacion, txt_original, carpeta). Solo lee archivos que escribió este programa (carpeta Reportes).
    Si el trabajo es de una versión con otros campos, toma los que coinciden (los nuevos quedan con su valor
    por omisión). ErrorUsuario si no se puede recuperar."""
    try:
        datos = pickle.loads(Path(ruta).read_bytes())
        guardados = datos["conciliacion"]
        conocidos = {f.name for f in fields(Conciliacion)}
        conciliacion = Conciliacion(**{k: v for k, v in guardados.items() if k in conocidos})
    except Exception as e:                                        # noqa: BLE001 - archivo dañado o de otra versión
        raise ErrorUsuario(f"No se pudo recuperar el trabajo guardado ({type(e).__name__}). Vuelve a procesar el mes.")
    return conciliacion, datos["txt"], Path(ruta).parent


def actualizar_sat(conciliacion, progreso=progreso_en_consola, avisar=print, consultar_sat=consultar_estados):
    """Vuelve a consultar el SAT (p. ej. si pasaron días: una factura pudo cancelarse) y vuelve a cruzar."""
    estados_sat, texto = consultar_sat(conciliacion.entradas["facturas"], progreso=progreso, avisar=avisar)
    sin_conexion = bool(estados_sat) and all(c.get("SinConexion") for c in estados_sat.values())
    if sin_conexion:
        return conciliacion, texto
    actualizada = replace(conciliacion, estados_sat=estados_sat,
                          estados={**conciliacion.estados, "Estatus en el SAT": texto},
                          sin_consultar=[x for x in conciliacion.sin_consultar if x != "SAT"],
                          consultado=datetime.now().strftime(FORMATO_FECHA))
    return revalidar(actualizada, conciliacion.entradas["lineas"]), texto


def actualizar_sap(conciliacion, carpeta, progreso=progreso_sap_en_consola, avisar=print, consultar=None):
    """Vuelve a consultar en SAP a todos los proveedores (p. ej. si Customer Services corrigió algo) y vuelve a cruzar."""
    e = conciliacion.entradas
    en_catalogo = set(e["catalogo"]["IDProveedor"])
    ids = sorted({l["IDProveedor"] for l in e["lineas"] if l["IDProveedor"] in en_catalogo})
    sap, texto = (consultar or consultar_sap)(ids, carpeta, conciliacion.anio, conciliacion.mes, False,
                                               progreso=progreso, avisar=avisar)
    if sap is None:
        return conciliacion, texto
    actualizada = replace(conciliacion, entradas={**e, "sap": sap},
                          estados={**conciliacion.estados, "Validación SAP": texto},
                          sin_consultar=[x for x in conciliacion.sin_consultar if x != "SAP"],
                          consultado=datetime.now().strftime(FORMATO_FECHA))
    return revalidar(actualizada, e["lineas"]), texto


def dias_desde_consulta(conciliacion, hoy=None):
    try:
        return ((hoy or datetime.now()) - datetime.strptime(conciliacion.consultado, FORMATO_FECHA)).days
    except ValueError:
        return 0


def regenerar_reporte(conciliacion, carpeta, txt_editado):
    """Excel con el resultado actual (con las ediciones) y la hoja "Cambios TXT". El TXT editado se guarda junto
    al Excel: su nombre y su huella SHA-256 quedan en la hoja Info."""
    e = conciliacion.entradas
    carpeta = Path(carpeta)
    carpeta.mkdir(parents=True, exist_ok=True)
    ruta_txt = carpeta / f"TXT_editado_{conciliacion.anio}-{conciliacion.mes:02d}.txt"
    ruta_txt.write_bytes(txt_editado)
    rutas = {"txt": ruta_txt, "catalogo": e["ruta_catalogo"], "salida": carpeta}
    destino, resumen = generar_reporte(rutas, conciliacion.anio, conciliacion.mes, conciliacion.resultado,
                                       conciliacion.sobrantes, conciliacion.errores_xml, e["n_xml"],
                                       conciliacion.estados, cambios=conciliacion.cambios)
    return replace(conciliacion, destino=destino, resumen=resumen)
