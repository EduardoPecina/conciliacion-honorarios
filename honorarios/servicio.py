"""Conciliación completa SIN preguntas en pantalla: la usan el programa de consola (.bat) y el portal web.

Aquí no hay reglas de negocio (esas viven en reglas.py y cruce.py): solo se encadenan los pasos
leer -> consultar SAT -> consultar SAP -> cruzar -> generar Excel, y se avisa el avance con funciones que
cada interfaz decide cómo mostrar (print en la consola, mensajes en el portal)."""

import hashlib
import hmac
import os
import pickle
import secrets
import tempfile
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
    # Lo leído y consultado en esta corrida (catálogo, líneas, facturas, SAP, historial): Actualizar SAT o SAP vuelve
    # a cruzar con esto sin repetir la consulta que no se pidió.
    entradas: dict = field(default_factory=dict, repr=False)
    version: int = 0                # sube con cada cambio: el portal sabe cuándo regenerar el Excel
    consultado: str = ""            # fecha y hora de las consultas al SAT y a SAP ("dd/mm/aaaa hh:mm")
    agregadas: list = field(default_factory=list)   # vueltas de "Facturas pendientes" (fecha, XML, resultado)

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
    """Vuelve a cruzar usando lo ya consultado al SAT y a SAP (tarda segundos). Las mismas reglas que la
    conciliación: nada se salta."""
    e = conciliacion.entradas
    resultado, sobrantes = conciliar(lineas, e["catalogo"], e["facturas"], conciliacion.anio, conciliacion.mes,
                                     e["sap"], e["historial"], conciliacion.estados_sat)
    return replace(conciliacion, resultado=resultado, sobrantes=sobrantes, entradas={**e, "lineas": lineas},
                   version=conciliacion.version + 1)


# ------------------------------------------------------------------ Facturas pendientes (refacturas)
def es_cruzada(fila):
    """Cruzado, Cruzado (pago quebrado) o Cruzado con observaciones: ya está lista; no se vuelve a tocar."""
    return str(fila.get("Estatus", "")).startswith("Cruzado")


def lineas_pendientes(conciliacion):
    """Filas del resultado que todavía no cruzan (Sin XML, Rechazar, Diferencia de importe, Detenida...)."""
    return [f for f in conciliacion.resultado if not es_cruzada(f)]


def agregar_pendientes(conciliacion, rutas_xml, carpeta_xml, avisar=print, progreso_sat=progreso_en_consola,
                       consultar_sat=consultar_estados, ahora=None, carpeta=None,
                       progreso_sap=progreso_sap_en_consola, consultar_sap=consultar_sap):
    """Segunda vuelta del mes: refacturas, XML que no habían llegado y médicos dados de alta después.

    Las líneas que ya cruzaron (y sus facturas) NO se tocan. El catálogo se vuelve a leer: un médico dado de alta
    después de procesar ya cuenta. Los XML nuevos se consultan en el SAT (solo esos) y se buscan, con las mismas
    reglas, entre las líneas pendientes, junto con las facturas del mes que ninguna línea cruzada usa (si el médico
    mandó una factura nueva, el cruce prefiere la válida a la que se rechazó). Lo consultado a SAP se reusa; solo los
    proveedores que nunca se consultaron (los nuevos en el catálogo) se consultan, si se da la carpeta del mes y SAP
    está abierto. Una factura agregada puede venir fechada hasta el mes siguiente (reglas.py).
    Se puede correr sin XML nuevos: así cruzan las líneas cuyo médico se acaba de dar de alta en el catálogo.
    -> (conciliación actualizada, resumen). ErrorUsuario si no hay nada nuevo o no hay conexión con el SAT."""
    ahora = ahora or datetime.now()
    e = conciliacion.entradas
    nuevas, errores = leer_xmls(rutas_xml, carpeta_xml) if rutas_xml else ([], [])
    del_mes = {f["UUID"] for f in e["facturas"] if f["UUID"]}
    ya_estaban = [f["ArchivoXML"] for f in nuevas if f["UUID"] and f["UUID"] in del_mes]
    nuevas = [{**f, "Agregada": ahora.strftime("%d/%m/%Y")} for f in nuevas
              if not (f["UUID"] and f["UUID"] in del_mes)]

    try:
        catalogo = leer_catalogo(e["ruta_catalogo"])                     # el de hoy, con las altas nuevas
    except OSError:
        avisar("No se pudo leer el catálogo: se usa el que se leyó al procesar el mes.")
        catalogo = e["catalogo"]
    antes_en_catalogo = set(e["catalogo"]["IDProveedor"])
    fijas = [f for f in conciliacion.resultado if es_cruzada(f)]
    lineas_fijas = {f["Linea"] for f in fijas}
    pendientes = [l for l in e["lineas"] if l["Linea"] not in lineas_fijas]
    altas = sorted({l["IDProveedor"] for l in pendientes
                    if l["IDProveedor"] in set(catalogo["IDProveedor"]) - antes_en_catalogo})
    cambio_catalogo = not catalogo.equals(e["catalogo"])
    resumen = {"leidos": len(rutas_xml), "nuevas": len(nuevas), "ya_estaban": ya_estaban, "errores": errores,
               "cambios": [], "sin_linea": 0, "altas": altas, "sap": ""}
    if not nuevas and not cambio_catalogo:
        raise ErrorUsuario(("Ningún XML nuevo: " + (f"los {len(ya_estaban)} ya estaban en la conciliación del mes "
                                                      "(mismo UUID)" if ya_estaban else "no se pudo leer ninguno")
                            if rutas_xml else "No hay nada nuevo")
                           + ". Sube solo las refacturas o los XML que faltaban; sin XML, la búsqueda sirve cuando "
                             "se corrigió el catálogo.")

    estados_sat = conciliacion.estados_sat
    if nuevas:
        avisar(f"Consultando en el SAT {len({f['UUID'] for f in nuevas if f['UUID']}):,} facturas nuevas...")
        estados_nuevos, _ = consultar_sat(nuevas, progreso=progreso_sat, avisar=avisar)
        if estados_nuevos and all(c.get("SinConexion") for c in estados_nuevos.values()):
            raise ErrorUsuario("No hubo conexión con el SAT: no se agregó nada. Vuelve a intentarlo en unos minutos.")
        estados_sat = {**estados_sat, **estados_nuevos}

    # Proveedores de las líneas pendientes que nunca se consultaron en SAP (los nuevos en el catálogo)
    sap = e["sap"]
    sin_sap = sorted({l["IDProveedor"] for l in pendientes if l["IDProveedor"] in set(catalogo["IDProveedor"])}
                     - set(sap or {}))
    if sin_sap and carpeta is not None:
        avisar(f"Consultando en SAP {len(sin_sap)} proveedor(es) que no se habían consultado...")
        consultados, resumen["sap"] = consultar_sap(sorted(set(sap or {}) | set(sin_sap)), Path(carpeta),
                                                    conciliacion.anio, conciliacion.mes, True,
                                                    progreso=progreso_sap, avisar=avisar)
        if consultados is not None:
            sap = {**(sap or {}), **consultados}

    avisar("Buscando en las líneas pendientes...")
    usadas = {i for f in fijas for i in f.get("_grupo", ())}
    uuid_pagados = {e["facturas"][i]["UUID"] for i in usadas} - {""}
    facturas = e["facturas"] + nuevas
    # Ni las facturas de las líneas cruzadas ni sus copias (mismo UUID) pueden cruzar otra línea: sería pago doble
    disponibles = [i for i in range(len(facturas)) if i not in usadas and facturas[i]["UUID"] not in uuid_pagados]
    filas, sobrantes = conciliar(pendientes, catalogo, [facturas[i] for i in disponibles],
                                 conciliacion.anio, conciliacion.mes, sap, e["historial"], estados_sat)
    copias = [x for x in conciliacion.sobrantes if x["UUID"] and x["UUID"] in uuid_pagados]
    sobrantes = copias + sobrantes                       # siguen en "XML sin TXT" con su motivo
    for fila in filas:                                   # índices de las facturas disponibles -> de todas
        if "_grupo" in fila:
            fila["_grupo"] = tuple(sorted(disponibles[i] for i in fila["_grupo"]))

    antes = {f["Linea"]: f["Estatus"] for f in conciliacion.resultado}
    resumen["cambios"] = [{"Línea": f["Linea"], "Médico": f.get("Nombre", ""), "Importe": float(f["Importe"]),
                           "Antes": antes.get(f["Linea"], ""), "Ahora": f["Estatus"]}
                          for f in filas if f["Estatus"] != antes.get(f["Linea"])]
    nuevas_uuid = {f["UUID"] for f in nuevas}
    resumen["sin_linea"] = sum(1 for s_ in sobrantes if s_["UUID"] in nuevas_uuid)
    cruzaron = sum(1 for c in resumen["cambios"] if c["Ahora"].startswith("Cruzado"))
    vuelta = len(conciliacion.agregadas) + 2
    texto = (f"{len(nuevas)} XML agregados el {ahora:%d/%m/%Y %H:%M}"
             + (f" y {len(altas)} proveedor(es) nuevos en el catálogo" if altas else "")
             + f": {cruzaron} línea(s) pasaron a Cruzado, {resumen['sin_linea']} XML sin línea pendiente")
    return replace(conciliacion,
                   resultado=sorted(fijas + filas, key=lambda f: f["Linea"]), sobrantes=sobrantes,
                   errores_xml=conciliacion.errores_xml + errores, estados_sat=estados_sat,
                   estados={**conciliacion.estados, f"Facturas pendientes (vuelta {vuelta})": texto},
                   entradas={**e, "facturas": facturas, "n_xml": e["n_xml"] + len(rutas_xml), "catalogo": catalogo,
                             "sap": sap},
                   agregadas=conciliacion.agregadas + [{"Fecha": f"{ahora:%d/%m/%Y %H:%M}", "Resultado": texto}],
                   version=conciliacion.version + 1), resumen


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


# El trabajo guardado va FIRMADO (HMAC-SHA256) con una clave que solo existe en esta computadora, guardada fuera de
# OneDrive (%LOCALAPPDATA%\\Honorarios). Al continuar, si la firma no coincide (archivo modificado, puesto por otra
# persona o guardado en otra computadora) NO se abre: el formato pickle podría ejecutar código si alguien lo alterara.
ENCABEZADO_TRABAJO = b"HONORARIOS-TRABAJO-FIRMADO-1\n"


def ruta_clave():
    if os.environ.get("HONORARIOS_CLAVE"):                     # para pruebas
        return Path(os.environ["HONORARIOS_CLAVE"])
    base = os.environ.get("LOCALAPPDATA") or Path.home() / ".config"
    return Path(base) / "Honorarios" / "clave_trabajo.bin"


def clave_de_esta_computadora():
    """Clave secreta de 256 bits para firmar el trabajo guardado; se crea la primera vez."""
    ruta = ruta_clave()
    try:
        clave = ruta.read_bytes()
        if len(clave) == 32:
            return clave
    except OSError:
        pass
    clave = secrets.token_bytes(32)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_bytes(clave)
    try:
        os.chmod(ruta, 0o600)                                   # solo el usuario (en Windows, AppData ya es suyo)
    except OSError:
        pass
    return clave


def firmar(cuerpo):
    return ENCABEZADO_TRABAJO + hmac.new(clave_de_esta_computadora(), cuerpo, hashlib.sha256).digest() + cuerpo


def verificar(datos):
    """-> el cuerpo, solo si la firma es de esta computadora y el archivo no se modificó. ErrorUsuario si no."""
    if not datos.startswith(ENCABEZADO_TRABAJO):
        raise ErrorUsuario("El trabajo guardado es de una versión anterior del programa o no lo escribió este "
                           "programa, y por seguridad no se abre. Vuelve a procesar el mes.")
    firma = datos[len(ENCABEZADO_TRABAJO):len(ENCABEZADO_TRABAJO) + 32]
    cuerpo = datos[len(ENCABEZADO_TRABAJO) + 32:]
    esperada = hmac.new(clave_de_esta_computadora(), cuerpo, hashlib.sha256).digest()
    if not hmac.compare_digest(firma, esperada):
        raise ErrorUsuario("El trabajo guardado no se pudo verificar: se guardó en otra computadora o el archivo "
                           "se modificó. Por seguridad no se abre. Vuelve a procesar el mes.")
    return cuerpo


def guardar_trabajo(conciliacion, carpeta, txt_original):
    """Guarda todo lo del mes (resultado, lo consultado al SAT y a SAP, el TXT original) para poder
    continuar otro día sin volver a procesar. Se guarda en la carpeta del mes; si falla, no detiene nada."""
    datos = {"programa": __version__, "guardado": datetime.now().strftime(FORMATO_FECHA), "txt": txt_original,
             "conciliacion": {f.name: getattr(conciliacion, f.name) for f in fields(conciliacion)}}
    destino = ruta_trabajo(carpeta, conciliacion.anio, conciliacion.mes)
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        firmado = firmar(pickle.dumps(datos, protocol=pickle.HIGHEST_PROTOCOL))
        guardar_seguro(destino, lambda tmp: tmp.write_bytes(firmado))
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
    """-> (conciliacion, txt_original, carpeta). Solo abre archivos firmados por este programa en esta computadora
    (la firma se comprueba ANTES de leer el contenido). Si el trabajo es de una versión con otros campos, toma los
    que coinciden (los nuevos quedan con su valor por omisión). ErrorUsuario si no se puede recuperar."""
    try:
        cuerpo = verificar(Path(ruta).read_bytes())
    except OSError:
        raise ErrorUsuario("No se pudo leer el trabajo guardado. Vuelve a procesar el mes.")
    try:
        datos = pickle.loads(cuerpo)
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


def regenerar_reporte(conciliacion, carpeta, txt_original):
    """Excel con el resultado actual (p. ej. después de Actualizar SAT o SAP). txt_original: (nombre, bytes) del TXT
    del mes; como al procesar, su nombre y su huella SHA-256 quedan en la hoja Info y sus líneas en "TXT para SAP"."""
    e = conciliacion.entradas
    carpeta = Path(carpeta)
    carpeta.mkdir(parents=True, exist_ok=True)
    nombre, datos = txt_original
    with tempfile.TemporaryDirectory() as tmp:
        ruta_txt = Path(tmp) / Path(nombre).name
        ruta_txt.write_bytes(datos)
        rutas = {"txt": ruta_txt, "catalogo": e["ruta_catalogo"], "salida": carpeta}
        destino, resumen = generar_reporte(rutas, conciliacion.anio, conciliacion.mes, conciliacion.resultado,
                                           conciliacion.sobrantes, conciliacion.errores_xml, e["n_xml"],
                                           conciliacion.estados)
    return replace(conciliacion, destino=destino, resumen=resumen)
