"""Estatus del CFDI en el SAT (Vigente / Cancelado / No encontrado).

Usa el servicio web oficial ConsultaCFDIService, no el portal de verificación: el portal tiene captcha
porque es para personas; el servicio web es el que el SAT publica para consultas automatizadas, no pide
contraseña ni e.firma y responde lo mismo que el portal.

Por cada factura se envía solo la "expresión impresa" del CFDI (la del código QR): UUID, RFC emisor,
RFC receptor, total y los últimos 8 caracteres del sello. Nada más sale de la computadora."""

import time
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from xml.sax.saxutils import escape

from .config import CONSULTAS_SAT_SIMULTANEAS, TIEMPO_ESPERA_SAT, URL_ESTADO_SAT
from .utilidades import a_decimal, normalizar_nombre


# ------------------------------------------------------------------
# 6. ESTATUS DEL CFDI EN EL SAT (servicio web ConsultaCFDIService)
# ------------------------------------------------------------------
SOAP_ACTION = "http://tempuri.org/IConsultaCFDIService/Consulta"
SOBRE = ('<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" '
         'xmlns:tem="http://tempuri.org/"><soapenv:Header/><soapenv:Body><tem:Consulta>'
         '<tem:expresionImpresa>{}</tem:expresionImpresa></tem:Consulta></soapenv:Body></soapenv:Envelope>')
CAMPOS_RESPUESTA = ("CodigoEstatus", "Estado", "EsCancelable", "EstatusCancelacion", "ValidacionEFOS")
# ValidacionEFOS: si el emisor está en la lista de EFOS ("Empresas que Facturan Operaciones Simuladas",
# listado definitivo del artículo 69-B del CFF). Códigos según la "Documentación del Servicio de Consulta
# de CFDI" del SAT (versión 1.4):
#   100 = el RFC del emisor SÍ está en la lista de EFOS                      -> alerta
#   200 = el RFC del emisor NO está en la lista                               -> sin alerta
#   201 = ni el emisor ni los RFC "a cuenta de terceros" de la factura están  -> sin alerta
# Cualquier otro código (uno que el SAT agregue después) también es alerta: mejor revisar de más que pagar a un
# EFOS. La librería phpcfdi/sat-estado-cfdi, muy usada con este servicio, aplica el mismo criterio.
EFOS_EN_LISTA = "100"
EFOS_NO_LISTADO = {"200", "201"}
ESTADOS = {"VIGENTE": "Vigente", "CANCELADO": "Cancelado", "NOENCONTRADO": "No encontrado"}
REINTENTOS = 3
PAUSA_REINTENTO = 2              # segundos; se duplica en cada reintento


def total_expresion(texto):
    """Total como lo pide la expresión impresa del CFDI 4.0: sin ceros de más y con al menos un decimal
    (1000.00 -> 1000.0, 2010.01 -> 2010.01)."""
    valor = a_decimal(texto)
    if valor is None:
        return str(texto)
    t = f"{valor:.6f}".rstrip("0")
    return t + "0" if t.endswith(".") else t


def expresion(f):
    """Expresión del código QR del CFDI (Anexo 20), que es lo que recibe el servicio del SAT."""
    partes = [f"id={f['UUID']}", f"re={f['RFCEmisor']}", f"rr={f['RFCReceptor']}",
              f"tt={total_expresion(f['TotalTexto'])}"]
    if f.get("SelloFinal"):
        partes.append(f"fe={f['SelloFinal']}")
    return "https://verificacfdi.facturaelectronica.sat.gob.mx/default.aspx?" + "&".join(partes)


def enviar(expresion_impresa):
    """Consulta SOAP al SAT; devuelve el texto de la respuesta."""
    cuerpo = SOBRE.format(escape(expresion_impresa)).encode("utf-8")
    peticion = urllib.request.Request(URL_ESTADO_SAT, data=cuerpo, method="POST", headers={
        "Content-Type": "text/xml; charset=utf-8", "SOAPAction": SOAP_ACTION})
    with urllib.request.urlopen(peticion, timeout=TIEMPO_ESPERA_SAT) as respuesta:
        return respuesta.read().decode("utf-8")


def leer_respuesta(texto):
    """Extrae CodigoEstatus, Estado, EsCancelable y EstatusCancelacion de la respuesta SOAP."""
    datos = {}
    for el in ET.fromstring(texto).iter():
        nombre = el.tag.rsplit("}", 1)[-1]
        if nombre in CAMPOS_RESPUESTA:
            datos[nombre] = (el.text or "").strip()
    return datos


def consultar_uno(f, enviar=enviar):
    """Devuelve {"Estado", "EstatusCancelacion", "EsCancelable", "Error", "SinConexion"}.
    Estado vacío = no se obtuvo respuesta válida (el motivo va en Error). Reintenta las fallas de red."""
    error, sin_conexion = "", False
    for intento in range(REINTENTOS):
        if intento:
            time.sleep(PAUSA_REINTENTO * 2 ** (intento - 1))
        try:
            datos = leer_respuesta(enviar(expresion(f)))
        except urllib.error.HTTPError as e:
            error, sin_conexion = f"el SAT respondió con error HTTP {e.code}", False
            continue
        except (urllib.error.URLError, OSError) as e:          # sin internet, proxy, tiempo agotado
            error, sin_conexion = f"sin respuesta del SAT ({getattr(e, 'reason', e)})", True
            continue
        except ET.ParseError:
            error, sin_conexion = "la respuesta del SAT no se pudo leer", False
            continue
        estado = ESTADOS.get(normalizar_nombre(datos.get("Estado", "")))
        if estado:
            return {"Estado": estado, "EstatusCancelacion": datos.get("EstatusCancelacion", ""),
                    "EsCancelable": datos.get("EsCancelable", ""), "ValidacionEFOS": datos.get("ValidacionEFOS", ""),
                    "Error": "", "SinConexion": False}
        # Respuesta válida pero sin estado (p. ej. "N - 601: La expresión impresa proporcionada no es
        # válida"): reintentar no la cambia.
        error = f"el SAT respondió '{datos.get('CodigoEstatus') or 'sin estatus'}'"
        return {"Estado": "", "EstatusCancelacion": "", "EsCancelable": "", "ValidacionEFOS": "", "Error": error, "SinConexion": False}
    return {"Estado": "", "EstatusCancelacion": "", "EsCancelable": "", "ValidacionEFOS": "", "Error": error,
            "SinConexion": sin_conexion}


def en_proceso_de_cancelacion(consulta):
    """El emisor pidió cancelar y el SAT espera que el receptor acepte o rechace."""
    return normalizar_nombre(consulta.get("EstatusCancelacion", "")) == "ENPROCESO"


def alerta_efos(consulta):
    """Código de ValidacionEFOS que indica que el emisor podría estar en la lista 69-B, o "" si no hay alerta."""
    codigo = (consulta or {}).get("ValidacionEFOS", "").strip()
    return codigo if codigo and codigo not in EFOS_NO_LISTADO else ""


def describir_efos(consulta):
    """Texto corto para la columna EFOS del Excel."""
    if not consulta or not consulta.get("Estado"):
        return ""
    codigo = alerta_efos(consulta)
    if codigo == EFOS_EN_LISTA:
        return f"Revisar: en lista EFOS (código {codigo})"
    if codigo:
        return f"Revisar: código {codigo} no documentado por el SAT"
    return "No listado" if consulta.get("ValidacionEFOS") else ""


def describir(consulta):
    """Texto corto para la columna EstadoSAT del Excel."""
    if not consulta:
        return "No se consultó"
    if not consulta["Estado"]:
        return "Sin respuesta"
    if consulta["Estado"] == "Vigente" and en_proceso_de_cancelacion(consulta):
        return "Vigente (cancelación en proceso)"
    return consulta["Estado"]


def consultar_estados(facturas, enviar=enviar):
    """Consulta cada UUID una sola vez. Devuelve ({UUID: consulta}, texto para la hoja Info)."""
    unicas = {}
    for f in facturas:
        if f["UUID"]:
            unicas.setdefault(f["UUID"], f)
    if not unicas:
        return {}, "No hay facturas timbradas que consultar"
    pendientes = list(unicas.values())

    # La primera va sola: si no hay internet se sabe de inmediato, sin esperar un tiempo agotado por factura
    resultado = {pendientes[0]["UUID"]: consultar_uno(pendientes[0], enviar)}
    if resultado[pendientes[0]["UUID"]]["SinConexion"]:
        error = resultado[pendientes[0]["UUID"]]["Error"]
        print(f"  No se pudo conectar al SAT: {error}.\n  ATENCIÓN: ninguna línea quedará 'Cruzado'; "
              "vuelve a correr el programa con conexión antes de pagar.")
        for f in pendientes[1:]:
            resultado[f["UUID"]] = {"Estado": "", "EstatusCancelacion": "", "EsCancelable": "",
                                    "Error": f"no se pudo conectar al SAT ({error})", "SinConexion": True}
        return resultado, f"No se pudo conectar al SAT: {error}"

    with ThreadPoolExecutor(CONSULTAS_SAT_SIMULTANEAS) as pool:
        futuros = {pool.submit(consultar_uno, f, enviar): f["UUID"] for f in pendientes[1:]}
        for k, futuro in enumerate(as_completed(futuros), start=2):
            print(f"  SAT {k}/{len(pendientes)}   ", end="\r")
            resultado[futuros[futuro]] = futuro.result()
    print()

    conteo = {}
    for consulta in resultado.values():
        clave = describir(consulta)
        conteo[clave] = conteo.get(clave, 0) + 1
    texto = ", ".join(f"{n:,} {estado.lower()}" for estado, n in sorted(conteo.items()))
    efos = sum(1 for consulta in resultado.values() if alerta_efos(consulta))
    return resultado, f"{len(resultado):,} UUID consultados: {texto}; {efos:,} con alerta EFOS (69-B)"
