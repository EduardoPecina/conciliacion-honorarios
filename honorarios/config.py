"""Configuración y reglas de negocio ajustables: datos del receptor, tasas, tolerancias,
catálogos del SAT y de SAP. Es el ÚNICO archivo que normalmente hay que editar."""

import os
import re
from decimal import Decimal
from pathlib import Path


# ------------------------------------------------------------------
# CONFIGURACIÓN (lo único que se podría ajustar)
# ------------------------------------------------------------------
RFC_SOCIEDAD = "SIN070530EI5"            # RFC receptor obligatorio en los XML
SOCIEDAD_SAP = "TM26"

# Datos del receptor en CFDI 4.0. Si se dejan vacíos, esa validación no se hace (se avisa al iniciar).
NOMBRE_SOCIEDAD = "SERVICIOS INTEGRALES NOVA DE MONTERREY"   # Razón social de la Constancia, sin "S.A. de C.V."
CP_SOCIEDAD = ""                         # Código postal del domicilio fiscal, ej. "64000"
REGIMEN_RECEPTOR = "601"                 # General de Ley Personas Morales
USOS_CFDI_OK = {"G03"}                   # Gastos en general
VERSIONES_CFDI_OK = {"4.0"}

ARCHIVO_CATALOGO = "CatalogoProveedoresTM26.xlsx"
NOMBRE_TABLA_CATALOGO = "Catalogo"
TOLERANCIA = Decimal("1.00")             # diferencia máxima de importe aceptada (cruce TXT vs XML)
TOLERANCIA_IMPUESTOS = Decimal("1.00")   # variación de impuestos aceptada como redondeo
UMBRAL_OK = Decimal("0.05")              # hasta aquí la variación se etiqueta OK; arriba, "Redondeo"

# True: una diferencia de importe entre 0.05 y 1.00 deja la línea "Cruzado con observaciones".
DIFERENCIA_IMPORTE_ES_OBSERVACION = True
# UUID repetido en la carpeta XML (la misma factura copiada dos veces):
#   False -> si solo UNA línea del TXT la paga, queda "Cruzado con observaciones" (no hay pago doble);
#            se detiene solo si dos líneas distintas terminan pagando la misma factura.
#   True  -> siempre "Detenida".
UUID_REPETIDO_SIEMPRE_DETIENE = False
MSG_UUID_REPETIDO = "Verificar con Responsables de Envío de XML"
MSG_SAP_SIN_IH = ("Discrepancia en Sistema e Información Enviada por Customer Services NOVA, "
                  "revisar Constancia Fiscal y Alta en el Sistema")
# True: si la fecha de la línea del TXT no es del periodo, se anota como observación.
FECHA_TXT_DEBE_SER_DEL_PERIODO = True

# Tasas de impuestos (las usa el programa para decidir el Estatus y el Excel para mostrar el cálculo)
TASA_IVA = Decimal("0.16")
TASA_ISR_GENERAL = Decimal("0.10")       # honorarios, régimen 612
TASA_ISR_RESICO = Decimal("0.0125")      # régimen 626

# Claves SAT de producto/servicio que se consideran servicios profesionales (primeros 2 dígitos)
PREFIJOS_CLAVE_SERVICIOS = ("80", "81", "82", "83", "84", "85", "86", "93")

ARCHIVO_HISTORIAL = "HistorialUUID.csv"

MESES = ["01 Ene", "02 Feb", "03 Mar", "04 Abr", "05 May", "06 Jun",
         "07 Jul", "08 Ago", "09 Sep", "10 Oct", "11 Nov", "12 Dic"]

COLUMNAS_TXT = ["Fecha", "IDProveedor", "Folio", "Importe", "FijoX",
                "CodigoImpuesto", "Fijo1", "CentroCosto", "Cuenta", "PM", "Referencia"]
CODIGOS_IMPUESTO_TXT = {"E0", "F2"}      # E0 sin IVA, F2 con IVA

METODOS_PAGO_OK = {"PUE", "PPD"}         # métodos de pago aceptados
FORMAS_PAGO_OK = {"03", "99"}            # transferencia electrónica / por definir

FORMAS_PAGO = {"01": "Efectivo", "02": "Cheque nominativo", "03": "Transferencia electrónica",
               "04": "Tarjeta de crédito", "28": "Tarjeta de débito", "99": "Por definir"}

REGIMENES = {
    "601": "General de Ley Personas Morales", "603": "Personas Morales con Fines no Lucrativos",
    "605": "Sueldos y Salarios", "606": "Arrendamiento", "607": "Enajenación o Adquisición de Bienes",
    "608": "Demás ingresos", "610": "Residentes en el Extranjero",
    "612": "Personas Físicas con Actividades Empresariales y Profesionales",
    "614": "Ingresos por intereses", "616": "Sin obligaciones fiscales",
    "621": "Incorporación Fiscal", "625": "Plataformas Tecnológicas",
    "626": "Régimen Simplificado de Confianza (RESICO)",
}

# Códigos de retención en SAP (tabla LFBW)
ISR_SAP = {"RS": "RESICO", "RH": "No RESICO", "RM": "No RESICO", "RE": "Extranjero"}

# Validaciones que una línea DEBE pasar para quedar "Cruzado" (clave -> nombre para el usuario).
# "sap" solo se exige si en esa corrida se eligió validar contra SAP.
VALIDACIONES = {
    "cfdi": "reglas del CFDI",
    "impuestos": "recálculo de impuestos",
    "sap": "códigos de retención en SAP",
    "historial": "historial de UUID conciliados",
}

# Observaciones que no quitan el "Cruzado" (son informativas, no hay nada que revisar)
INFORMATIVAS = ("Pago quebrado", "Factura PPD")
# Observaciones que mencionan SAP pero no son discrepancias con SAP
NO_ES_DISCREPANCIA_SAP = ("Lleva IVA: falta validar", "Validaciones no realizadas", "No se pudo validar en SAP",
                          "ID fiscal extranjero")

# Carpeta "Honorarios" (la que contiene _Programa, _Configuracion y los años).
# Se puede cambiar con la variable de entorno HONORARIOS_RAIZ (lo usan las pruebas).
RAIZ = Path(os.environ.get("HONORARIOS_RAIZ") or Path(__file__).resolve().parents[2])
CONFIG = RAIZ / "_Configuracion"
PATRON_RFC = re.compile(r"[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}")
PARECE_RFC_MX = re.compile(r"[A-ZÑ&]{3,4}\d")          # empieza como RFC mexicano (3-4 letras y dígitos)
RFC_GENERICO_EXTRANJERO = "XEXX010101000"
