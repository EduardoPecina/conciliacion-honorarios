"""Reglas de negocio por factura: CFDI, receptor, pago, códigos SAP y recálculo de impuestos."""

from decimal import Decimal, ROUND_HALF_UP

from .config import (
    CP_SOCIEDAD,
    FORMAS_PAGO,
    FORMAS_PAGO_OK,
    ISR_SAP,
    METODOS_PAGO_OK,
    MSG_SAP_SIN_IH,
    MSG_UUID_REPETIDO,
    NOMBRE_SOCIEDAD,
    PATRON_RFC,
    PREFIJOS_CLAVE_SERVICIOS,
    REGIMENES,
    REGIMEN_RECEPTOR,
    RFC_SOCIEDAD,
    TASA_ISR_GENERAL,
    TASA_ISR_RESICO,
    TASA_IVA,
    TOLERANCIA_IMPUESTOS,
    UMBRAL_OK,
    USOS_CFDI_OK,
    UUID_REPETIDO_SIEMPRE_DETIENE,
    VERSIONES_CFDI_OK,
)
from .estado_sat import en_proceso_de_cancelacion
from .utilidades import a_decimal, normalizar_nombre


def validar_xml(f, ctx):
    """Reglas de negocio de una factura.
    Devuelve (motivos de rechazo, observaciones, motivos para detener, validaciones aprobadas)."""
    r, o, d, ok = [], [], [], set()

    # Timbre y versión
    if not f["UUID"]:
        r.append("El XML no tiene Timbre Fiscal Digital (no está timbrado)")
    if f["Version"] not in VERSIONES_CFDI_OK:
        r.append(f"CFDI versión {f['Version'] or 'desconocida'}, solo se acepta {', '.join(VERSIONES_CFDI_OK)}")

    # Receptor
    if f["RFCReceptor"] != RFC_SOCIEDAD:
        r.append(f"Facturado a {f['RFCReceptor'] or 'RFC vacío'}, debe ser {RFC_SOCIEDAD}")
    if REGIMEN_RECEPTOR and f["RegimenReceptor"] != REGIMEN_RECEPTOR:
        r.append(f"Régimen del receptor {f['RegimenReceptor'] or 'vacío'}, debe ser {REGIMEN_RECEPTOR} "
                 f"{REGIMENES.get(REGIMEN_RECEPTOR, '')}".strip())
    if CP_SOCIEDAD and f["CPReceptor"] != CP_SOCIEDAD:
        r.append(f"Código postal del receptor {f['CPReceptor'] or 'vacío'}, debe ser {CP_SOCIEDAD}")
    if NOMBRE_SOCIEDAD and normalizar_nombre(f["NombreReceptor"]) != normalizar_nombre(NOMBRE_SOCIEDAD):
        r.append(f"Nombre del receptor '{f['NombreReceptor']}', debe ser '{NOMBRE_SOCIEDAD}'")
    if f["UsoCFDI"] not in USOS_CFDI_OK:
        r.append(f"Uso de CFDI {f['UsoCFDI'] or 'vacío'}, debe ser {', '.join(sorted(USOS_CFDI_OK))}")

    # Emisor
    if not PATRON_RFC.fullmatch(f["RFCEmisor"]):
        r.append(f"RFC del emisor '{f['RFCEmisor']}' con formato inválido")
    elif len(f["RFCEmisor"]) != 13:
        d.append("El emisor es persona moral (RFC de 12 caracteres): las retenciones de honorarios "
                 "de persona física no aplican")

    # Periodo, tipo, método y forma de pago
    if f["FechaXML"][:7] != ctx["periodo"]:
        r.append(f"Factura del {f['FechaXML']}, fuera del periodo {ctx['mes']:02d}/{ctx['anio']}: "
                 "pedir refacturación")
    if f["TipoComprobante"] != "I":
        r.append(f"Tipo de comprobante {f['TipoComprobante']}, se esperaba I (Ingreso)")
    if f["MetodoPago"] not in METODOS_PAGO_OK:
        r.append(f"Método de pago {f['MetodoPago'] or 'vacío'}, debe ser PUE o PPD")
    if f["FormaPago"] not in FORMAS_PAGO_OK:
        nombre = FORMAS_PAGO.get(f["FormaPago"], "")
        r.append(f"Forma de pago {f['FormaPago'] or 'vacía'} {nombre}".strip()
                 + ", solo se acepta 03 Transferencia o 99 Por definir")
    if f["MetodoPago"] == "PPD" and f["FormaPago"] != "99":
        r.append(f"Método PPD con forma {f['FormaPago'] or 'vacía'}: en PPD la forma debe ser 99 Por definir")
    if f["MetodoPago"] == "PPD":
        o.append("Factura PPD: el doctor debe enviar el comprobante de pago (complemento de pago / REP) "
                 "después de recibir el pago")
    if f["Moneda"] != "MXN":
        d.append(f"Moneda {f['Moneda'] or 'vacía'}: los importes no son comparables contra el TXT en pesos")

    # Duplicados dentro del mes
    veces = ctx["uuids"].get(f["UUID"], 0)
    if f["UUID"] and veces > 1:
        texto = f"{MSG_UUID_REPETIDO}: el mismo UUID viene en {veces} archivos XML"
        (d if UUID_REPETIDO_SIEMPRE_DETIENE else o).append(texto)

    # Aritmética del propio XML
    if abs(f["DifTotalXML"]) > TOLERANCIA_IMPUESTOS:
        d.append(f"El Total del XML ({f['Total']:,.2f}) no cuadra con subtotal − descuento + impuestos − "
                 f"retenciones (diferencia {f['DifTotalXML']:,.2f})")
    if f["OtrosImpuestos"]:
        o.append(f"La factura trae otros impuestos (IEPS o locales) por {f['OtrosImpuestos']:,.2f}")

    # Conceptos
    if not f["ClavesProdServ"]:
        r.append("La factura no trae conceptos")
    elif abs(f["SumaConceptos"] - f["SubTotal"]) > TOLERANCIA_IMPUESTOS:
        d.append(f"Los conceptos suman {f['SumaConceptos']:,.2f} pero el SubTotal dice {f['SubTotal']:,.2f}: "
                 "XML alterado o mal emitido")
    malas = sorted({c for c in f["ClavesProdServ"] if not c.startswith(PREFIJOS_CLAVE_SERVICIOS)})
    if malas:
        o.append(f"Clave(s) de producto/servicio {', '.join(malas)} no corresponden a servicios profesionales")
    if f["IVA"] > 0 and f["ObjetosImp"] and all(x == "01" for x in f["ObjetosImp"]):
        o.append("Conceptos marcados 01 (no objeto de impuesto) pero la factura trae IVA")
    # Un concepto 02 con solo retención de ISR (honorarios E0, p. ej. médicos) es válido: no se avisa.
    # Si el TXT dice F2 y falta el IVA lo detiene el recálculo; si SAP tiene IH, la validación SAP.
    if f.get("Conceptos02SinImpuestos"):
        o.append(f"{f['Conceptos02SinImpuestos']} concepto(s) marcados 02 (objeto de impuesto) sin ningún "
                 "impuesto desglosado: XML mal emitido")
    ok.add("cfdi")

    # Estatus en el SAT (solo si en esta corrida se pidió consultarlo)
    if ctx.get("estados_sat") is not None and f["UUID"]:
        consulta = ctx["estados_sat"].get(f["UUID"])
        if consulta and consulta["Estado"]:
            if consulta["Estado"] == "Cancelado":
                r.append("Factura CANCELADA en el SAT: no se puede pagar, pedir la factura que la sustituye")
            elif consulta["Estado"] == "No encontrado":
                d.append("El SAT no encuentra este CFDI con el UUID, RFC y total del XML: puede ser un XML alterado "
                         "o apócrifo, o recién timbrado (puede tardar hasta 72 horas en aparecer); volver a consultar")
            elif en_proceso_de_cancelacion(consulta):
                d.append(f"El emisor pidió cancelar esta factura y el SAT espera que {RFC_SOCIEDAD} acepte o rechace "
                         "la cancelación en el portal del SAT (si no responde en 3 días hábiles, se cancela sola): "
                         "aclarar con el médico antes de pagar")
            ok.add("estado_sat")
        else:
            o.append("No se pudo consultar el estatus en el SAT: "
                     + (consulta["Error"] if consulta else "no se consultó este UUID"))

    # Historial de meses anteriores
    if ctx["historial"] is not None:
        previos = [x for x in ctx["historial"].get(f["UUID"], []) if x.get("Periodo") != ctx["periodo"]]
        if previos:
            p = previos[0]
            d.append(f"Este UUID ya se concilió en el periodo {p.get('Periodo')} (línea {p.get('Linea')} "
                     f"del TXT, importe {p.get('Importe')}): verificar que no se pague dos veces")
        ok.add("historial")
    return r, o, d, ok


def validar_sap(codigo_txt, regimen_xml, codigos, xml_con_iva):
    """Compara TXT (E0/F2), IVA y régimen del XML contra los códigos de retención de SAP.
    Devuelve (régimen según SAP, observaciones, motivos para detener la factura)."""
    obs, detener = [], []
    c = set(codigos)
    if (codigo_txt == "F2" or xml_con_iva) and "IH" not in c:
        origen = "el TXT dice F2" if codigo_txt == "F2" else "el XML trae IVA"
        detener.append(f"{MSG_SAP_SIN_IH} (lleva IVA porque {origen}, pero SAP no tiene IH)")
    if "IP" in c:
        detener.append("SAP tiene IP (Ret. IVA Pitex 100%): retendría todo el IVA, no corresponde a honorarios")
    if codigo_txt == "E0" and "IH" in c:
        obs.append("TXT dice E0 (sin IVA) pero SAP tiene IH (retención de IVA)")
    isr = c & set(ISR_SAP)
    if len(isr) != 1:
        detener.append(f"SAP con indicadores de ISR {sorted(isr) or 'ninguno'}, se esperaba uno: "
                       "no se sabe cuánto ISR retendrá")
        return "", obs, detener
    regimen_sap = ISR_SAP[isr.pop()]
    if regimen_sap == "Extranjero":
        detener.append("SAP marca al proveedor como Extranjero (RE) pero factura con CFDI nacional")
    for reg in [r.strip() for r in regimen_xml.split(",") if r.strip()]:
        if regimen_sap == "RESICO" and reg != "626":
            detener.append(f"SAP retendrá ISR RESICO 1.25% (RS) pero el XML trae régimen {reg}: "
                           "el pago no coincidirá con la factura")
        elif regimen_sap == "No RESICO" and reg == "626":
            detener.append("SAP retendrá ISR 10% (RH) pero el XML es RESICO 626: "
                           "el pago no coincidirá con la factura")
    return regimen_sap, obs, detener


def redondear(valor):
    """Redondea a 2 decimales igual que ROUND de Excel (0.005 sube)."""
    return valor.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def revisar_impuestos(fila):
    """Recalcula IVA y retenciones de una línea cruzada (mismas reglas que la hoja Cálculos).
    Devuelve el motivo para detenerla, o "" si la variación está dentro de la tolerancia."""
    base = fila["SubTotal"] - fila["Descuento"]
    con_iva = fila["CodigoImpuesto"] == "F2"
    regimenes = {r.strip() for r in fila.get("RegimenEmisor", "").split(",") if r.strip()}
    resico = regimenes == {"626"}
    iva_calc = redondear(base * TASA_IVA) if con_iva else Decimal("0")
    isr_calc = redondear(base * (TASA_ISR_RESICO if resico else TASA_ISR_GENERAL))
    riva_calc = redondear(iva_calc * 2 / 3)
    conceptos = [
        ("IVA", fila["IVA"], iva_calc, "16%" if con_iva else "el TXT dice E0: sin IVA"),
        ("Ret. ISR", fila["RetISR"], isr_calc, "RESICO 1.25%" if resico else "10%"),
        ("Ret. IVA", fila["RetIVA"], riva_calc, "2/3 del IVA" if con_iva else "sin IVA no se retiene"),
    ]
    variacion = sum(abs(xml - calc) for _, xml, calc, _ in conceptos)
    if variacion <= TOLERANCIA_IMPUESTOS:
        return ""
    partes = [f"{nombre} {xml:,.2f}, debería ser {calc:,.2f} ({regla})"
              for nombre, xml, calc, regla in conceptos if abs(xml - calc) > UMBRAL_OK]
    if fila["RetIVA"] > 0 and fila["IVA"] == 0:
        partes.append("retiene IVA sin cobrar IVA (¿capturó la retención de ISR como IVA?)")
    tasas = {a_decimal(t) for t in fila.get("_tasas_iva", set())} - {None}
    if any(t != TASA_IVA for t in tasas):
        partes.append("el XML usa tasa de IVA " + ", ".join(f"{t:.0%}" for t in sorted(tasas))
                      + " (¿tasa fronteriza 8%?)")
    if len(regimenes) > 1:
        partes.append(f"las facturas traen regímenes distintos ({', '.join(sorted(regimenes))})")
    return f"Impuestos no cuadran (variación {variacion:,.2f}): " + "; ".join(partes)
