"""Cruce TXT -> Catálogo -> XML: uno a uno y pagos quebrados (N líneas contra M facturas),
aplicación de reglas y asignación del Estatus final."""

from bisect import bisect_left
from decimal import Decimal
from itertools import combinations
from math import comb

from .config import (
    CODIGOS_IMPUESTO_TXT,
    DIFERENCIA_IMPORTE_ES_OBSERVACION,
    FECHA_TXT_DEBE_SER_DEL_PERIODO,
    FORMAS_PAGO,
    INFORMATIVAS,
    MSG_FACTURA_CANCELADA,
    MSG_UUID_REPETIDO,
    NO_ES_DISCREPANCIA_SAP,
    REGIMENES,
    TOLERANCIA,
    UMBRAL_OK,
    VALIDACIONES,
)
from .estado_sat import describir as describir_estado_sat
from .utilidades import clasificar_rfc
from .reglas import redondear, revisar_impuestos, validar_sap, validar_xml


# ------------------------------------------------------------------
# 3. CRUCE
# ------------------------------------------------------------------
# Pago quebrado = cualquier grupo de N líneas del TXT contra M facturas del MISMO proveedor
# (N ó M pueden ser 1, pero no ambos: eso es el cruce uno a uno).
MAX_PIEZAS_GRUPO = 10          # máximo de piezas (líneas + facturas) en un pago quebrado
MAX_COMBINACIONES = 400_000    # presupuesto de combinaciones por búsqueda (acota el tiempo)
MAX_CANDIDATOS = 200           # grupos alternativos que se guardan para detectar ambigüedad

CAMPOS_XML = ("SubTotal", "Descuento", "IVA", "RetISR", "RetIVA", "Total")


def a_centavos(valor):
    return int(redondear(valor) * 100)


def buscar_grupo(lineas, xmls, tolerancia, preferencia=lambda idxs: 0):
    """Busca el grupo más chico de líneas y facturas cuyas sumas coinciden dentro de la tolerancia.

    lineas, xmls: listas de (id, importe) de UN proveedor. Se prueban tamaños de grupo crecientes
    (3 piezas, luego 4...) y para cada tamaño todas las formas de repartirlo (1 línea + 2 XML,
    2 líneas + 1 XML...). Las sumas de un lado se ordenan y el otro lado se busca por bisección,
    así que cada combinación cuesta O(log n) y no O(n).

    Devuelve (ids_lineas, ids_xml, ambigua, truncada) o None si no hay grupo.
      ambigua  = hay otro grupo del mismo tamaño que comparte alguna pieza y tiene importes distintos
                 (si solo intercambia piezas del mismo importe, la elección no cambia el resultado).
      truncada = hubo repartos que no se exploraron por el presupuesto de combinaciones."""
    tol = a_centavos(tolerancia)
    L = [(i, v, a_centavos(v)) for i, v in lineas]
    X = [(i, v, a_centavos(v)) for i, v in xmls]
    truncada, trabajo = False, 0
    for k in range(3, min(MAX_PIEZAS_GRUPO, len(L) + len(X)) + 1):
        candidatos = []
        for s in range(1, k):
            t = k - s
            if s > len(L) or t > len(X):
                continue
            costo = comb(len(L), s) + comb(len(X), t)
            if trabajo + costo > MAX_COMBINACIONES:
                truncada = True
                continue
            trabajo += costo
            sumas_x = sorted((sum(c for *_, c in combo), combo) for combo in combinations(X, t))
            claves = [c for c, _ in sumas_x]
            for combo_l in combinations(L, s):
                suma_l = sum(c for *_, c in combo_l)
                # margen de 1 centavo por el redondeo a centavos; abajo se verifica con el importe exacto
                j = bisect_left(claves, suma_l - tol - 1)
                while j < len(claves) and claves[j] <= suma_l + tol + 1:
                    combo_x = sumas_x[j][1]
                    dif = sum(v for _, v, _ in combo_x) - sum(v for _, v, _ in combo_l)
                    if abs(dif) <= tolerancia:
                        ids_l, ids_x = tuple(i for i, *_ in combo_l), tuple(i for i, *_ in combo_x)
                        firma = (tuple(sorted(c for *_, c in combo_l)), tuple(sorted(c for *_, c in combo_x)))
                        candidatos.append(((preferencia(ids_x), abs(dif), ids_l, ids_x), ids_l, ids_x, firma))
                        if len(candidatos) >= MAX_CANDIDATOS:
                            break
                    j += 1
                if len(candidatos) >= MAX_CANDIDATOS:
                    break
        if candidatos:
            _, ids_l, ids_x, firma = min(candidatos)
            piezas = set(ids_l) | {("x", i) for i in ids_x}
            ambigua = any((set(l) | {("x", i) for i in x}) & piezas and f != firma
                          for _, l, x, f in candidatos)
            return list(ids_l), list(ids_x), ambigua, truncada
    return None


def conciliar(lineas_txt, catalogo, facturas, anio, mes, sap=None, historial=None, estados_sat=None):
    por_id = catalogo.set_index("IDProveedor").to_dict("index")
    uuids_vistos = {}
    for f in facturas:
        if f["UUID"]:
            uuids_vistos[f["UUID"]] = uuids_vistos.get(f["UUID"], 0) + 1
    ctx = {"uuids": uuids_vistos, "anio": anio, "mes": mes, "periodo": f"{anio}-{mes:02d}",
           "historial": historial, "estados_sat": estados_sat}
    validacion = [validar_xml(f, ctx) for f in facturas]      # se calcula una vez por factura

    xml_por_rfc = {}
    for i, f in enumerate(facturas):
        xml_por_rfc.setdefault(f["RFCEmisor"], []).append(i)

    filas, obs, rech, deten, pasadas = [], [], [], [], []
    pendientes = {}                 # RFC -> índices de líneas aún sin XML
    repetidas = {}                  # (proveedor, folio, importe) -> líneas
    for n, linea in enumerate(lineas_txt):
        fila = {k: linea[k] for k in ("Linea", "Fecha", "IDProveedor", "Folio", "Importe",
                                      "CodigoImpuesto", "CentroCosto", "Cuenta")}
        fila["Estatus"] = "Sin XML"
        o, d = [], []
        if linea["CodigoImpuesto"] not in CODIGOS_IMPUESTO_TXT:
            d.append(f"Código de impuesto '{linea['CodigoImpuesto']}' desconocido en el TXT (se espera E0 o F2)")
        if FECHA_TXT_DEBE_SER_DEL_PERIODO and linea["_fecha"].strftime("%Y-%m") != ctx["periodo"]:
            o.append(f"La fecha del TXT ({linea['Fecha']}) no es del periodo {mes:02d}/{anio}")
        repetidas.setdefault((linea["IDProveedor"], linea["Folio"], linea["Importe"]), []).append(n)

        prov = por_id.get(linea["IDProveedor"])
        if not linea["IDProveedor"]:
            fila.update(Estatus="Línea sin IDProveedor", RFC="", Nombre="")
            d.append("La línea del TXT no trae IDProveedor: no se puede saber a quién se le paga")
        elif prov is None:
            fila.update(Estatus="ID no está en catálogo", RFC="", Nombre="")
        else:
            fila["RFC"], fila["Nombre"] = prov["RFC"], prov["Nombre"]
            tipo = prov.get("TipoRFC") or clasificar_rfc(prov["RFC"])
            if tipo == "Vacío":
                fila["Estatus"] = "Proveedor sin RFC en catálogo"
            elif tipo == "Extranjero":
                fila["Estatus"] = "Proveedor extranjero"
                o.append(f"ID fiscal extranjero '{prov['RFC']}': no emite CFDI, se revisa con su factura "
                         "del extranjero (y retención RE en SAP)")
            elif tipo == "Inválido":
                fila["Estatus"] = "RFC inválido en catálogo"
                d.append(f"RFC '{prov['RFC']}' mal escrito en el catálogo: corregirlo para poder cruzar")
            elif linea["Importe"] <= 0:
                d.append(f"Importe {linea['Importe']:,.2f} en el TXT: debe ser mayor a cero")
            else:
                pendientes.setdefault(prov["RFC"], []).append(n)
        filas.append(fila)
        obs.append(o)
        rech.append([])
        deten.append(d)
        pasadas.append(set())

    for ns in repetidas.values():
        if len(ns) > 1:
            nums = ", ".join(str(lineas_txt[n]["Linea"]) for n in ns)
            for n in ns:
                deten[n].append(f"Líneas {nums} del TXT repetidas (mismo proveedor, folio e importe): "
                                "posible pago doble")

    usadas = set()
    libres = lambda rfc: [i for i in xml_por_rfc.get(rfc, []) if i not in usadas]

    def poner_xml(n, idxs, estatus, notas=(), con_importes=True):
        fs = [facturas[i] for i in idxs]
        fila = filas[n]
        fila["Estatus"] = estatus
        fila["_grupo"] = tuple(sorted(idxs))
        fila["ArchivoXML"] = ", ".join(f["ArchivoXML"] for f in fs)
        fila["UUID"] = ", ".join(f["UUID"] for f in fs)
        if estados_sat is not None:
            fila["EstadoSAT"] = ", ".join(sorted({describir_estado_sat(estados_sat.get(f["UUID"])) or "Sin UUID"
                                                  for f in fs}))
        fila["FechaXML"] = ", ".join(sorted({f["FechaXML"] for f in fs}))
        fila["RegimenEmisor"] = ", ".join(sorted({f["RegimenEmisor"] for f in fs}))
        if con_importes:
            for c in CAMPOS_XML:
                fila[c] = sum((f[c] for f in fs), Decimal("0"))
        fila["_iva_xml"] = any(f["IVA"] > 0 for f in fs)
        fila["_tasas_iva"] = set().union(*(f["TasasIVA"] for f in fs))
        fila["NombreRegimen"] = ", ".join(sorted({REGIMENES.get(f["RegimenEmisor"], "Desconocido")
                                                  for f in fs}))
        fila["MetodoPago"] = ", ".join(sorted({f["MetodoPago"] for f in fs}))
        fila["FormaPago"] = ", ".join(sorted({f"{f['FormaPago']} {FORMAS_PAGO.get(f['FormaPago'], '')}".strip()
                                              for f in fs}))
        aprobadas = None
        for i in idxs:
            r, o, d, ok = validacion[i]
            rech[n].extend(r)
            obs[n].extend(o)
            deten[n].extend(d)
            aprobadas = set(ok) if aprobadas is None else aprobadas & ok   # debe aprobar en TODAS las facturas
        pasadas[n] |= aprobadas
        obs[n].extend(notas)

    def cancelada(i):
        return estados_sat is not None and (estados_sat.get(facturas[i]["UUID"]) or {}).get("Estado") == "Cancelado"

    def sustitutas(i, candidatas):
        """Facturas vigentes del mismo RFC y por el mismo importe que la factura i (que está cancelada)."""
        return [j for j in candidatas if j != i and not cancelada(j)
                and facturas[j]["RFCEmisor"] == facturas[i]["RFCEmisor"]
                and abs(facturas[j]["ImporteConIVA"] - facturas[i]["ImporteConIVA"]) <= TOLERANCIA]

    def calidad(i):
        """0 = factura válida, 1 = con motivo para detener, 2 = a rechazar. Ante importes iguales se
        prefiere la válida: p. ej. el proveedor canceló y reemitió por el mismo importe."""
        r, _, d, _ = validacion[i]
        return 2 if r else (1 if d else 0)

    def emparejar(lineas, candidatas_de):
        """Emparejamiento máximo línea <-> XML (algoritmo de Kuhn). A diferencia de asignar en orden,
        ninguna línea se queda sin factura si existe un acomodo en el que todas cruzan.
        candidatas_de(n) devuelve los XML posibles ya ordenados por preferencia."""
        asignada = {}                                  # xml -> línea

        def intentar(n, vistos):
            for i in candidatas_de(n):
                if i not in vistos:
                    vistos.add(i)
                    if i not in asignada or intentar(asignada[i], vistos):
                        asignada[i] = n
                        return True
            return False

        for n in lineas:
            intentar(n, set())
        return {n: i for i, n in asignada.items()}

    def uno_a_uno(limite):
        """Una línea = una factura, con emparejamiento máximo por proveedor."""
        for rfc, lineas in pendientes.items():
            pool = libres(rfc)
            cerca = lambda n, lim: [i for i in pool
                                    if abs(facturas[i]["ImporteConIVA"] - lineas_txt[n]["Importe"]) <= lim]
            orden = lambda n: sorted(cerca(n, limite), key=lambda i: (
                calidad(i), abs(facturas[i]["ImporteConIVA"] - lineas_txt[n]["Importe"])))
            asignacion = emparejar(list(lineas), orden)
            tomadas = set(asignacion.values())
            for n in sorted(asignacion):
                i = asignacion[n]
                otras = [j for j in cerca(n, TOLERANCIA) if j != i and j not in tomadas
                         and facturas[j]["UUID"] != facturas[i]["UUID"]]     # copia del mismo UUID: otro aviso
                # Una factura cancelada por el mismo importe no es facturación doble: es la que esta sustituye
                sustituidas = [j for j in otras if cancelada(j)] if not cancelada(i) else []
                otras = [j for j in otras if j not in sustituidas]
                notas = ([f"{len(otras) + 1} facturas distintas (UUID diferente) del mismo RFC por el mismo "
                          "importe: posible facturación doble del proveedor"]
                         if otras else [])
                if sustituidas:
                    notas.append("Sustituye a la factura cancelada en el SAT "
                                 + ", ".join(facturas[j]["ArchivoXML"] for j in sustituidas))
                usadas.add(i)
                poner_xml(n, [i], "Cruzado", notas)
                lineas.remove(n)

    def poner_grupo(ns, idxs, notas):
        ns = sorted(ns, key=lambda n: lineas_txt[n]["Linea"])
        usadas.update(idxs)
        for k, n in enumerate(ns):
            poner_xml(n, idxs, "Cruzado (pago quebrado)", notas, con_importes=(k == 0))
            pendientes[filas[n]["RFC"]].remove(n)

    def describir(ns, idxs):
        nums = ", ".join(str(lineas_txt[n]["Linea"]) for n in sorted(ns, key=lambda n: lineas_txt[n]["Linea"]))
        archivos = ", ".join(facturas[i]["ArchivoXML"] for i in idxs)
        total_txt = sum((lineas_txt[n]["Importe"] for n in ns), Decimal("0"))
        total_xml = sum((facturas[i]["ImporteConIVA"] for i in idxs), Decimal("0"))
        return (f"Pago quebrado: {len(ns)} línea(s) del TXT ({nums}) por {total_txt:,.2f} contra "
                f"{len(idxs)} factura(s) ({archivos}) por {total_xml:,.2f}")

    def lineas_por_id(lineas):
        """El ID del proveedor delimita el grupo: nunca se mezclan líneas de IDs distintos,
        aunque en el catálogo compartan RFC."""
        por_id = {}
        for n in lineas:
            por_id.setdefault(lineas_txt[n]["IDProveedor"], []).append(n)
        return por_id.values()

    def agrupados(tolerancia):
        """Pago quebrado N:M. Se repite por proveedor hasta que ya no aparece ningún grupo."""
        for rfc in pendientes:
            for grupo_id in lineas_por_id(list(pendientes[rfc])):
                grupo_id = list(grupo_id)
                while grupo_id:
                    pool = [i for i in libres(rfc) if facturas[i]["ImporteConIVA"] > 0]
                    if not pool or len(grupo_id) + len(pool) < 3:
                        break
                    hallado = buscar_grupo([(n, lineas_txt[n]["Importe"]) for n in grupo_id],
                                           [(i, facturas[i]["ImporteConIVA"]) for i in pool],
                                           tolerancia, lambda idxs: sum(calidad(i) for i in idxs))
                    if hallado is None:
                        break
                    ns, idxs, ambigua, truncada = hallado
                    notas = [describir(ns, idxs)]
                    if ambigua:
                        notas.append("Hay más de una combinación de líneas y facturas que suma lo mismo: "
                                     "confirmar qué pago corresponde a qué factura")
                    if truncada:
                        notas.append(f"El proveedor tiene muchas piezas pendientes; la búsqueda se limitó a "
                                     f"grupos de hasta {MAX_PIEZAS_GRUPO} y pudo no encontrar el más chico")
                    poner_grupo(ns, idxs, notas)
                    grupo_id = [n for n in grupo_id if n not in ns]

    def por_total_del_proveedor():
        """Red de seguridad: si lo pendiente del TXT de un proveedor suma lo mismo que sus facturas
        pendientes, cuadra en total aunque no se sepa qué pago corresponde a qué factura."""
        for rfc, lineas in pendientes.items():
            ids = {lineas_txt[n]["IDProveedor"] for n in lineas}
            pool = [i for i in libres(rfc) if facturas[i]["ImporteConIVA"] > 0]
            if len(ids) != 1 or not pool or len(lineas) + len(pool) < 3:
                continue
            ns = list(lineas)
            dif = (sum(facturas[i]["ImporteConIVA"] for i in pool) - sum(lineas_txt[n]["Importe"] for n in ns))
            if abs(dif) <= TOLERANCIA:
                poner_grupo(ns, pool, [describir(ns, pool),
                                       "Cuadre por total del proveedor: no se pudo determinar qué pago "
                                       "corresponde a qué factura, revisar el reparto"])

    # Orden de los cruces: primero la evidencia más fuerte (exacto) y los grupos más chicos.
    #   1. Uno a uno exacto (hasta 0.05)
    #   2. Pago quebrado exacto (N líneas contra M facturas, hasta 0.05)
    #   3. Uno a uno con tolerancia (hasta 1.00)
    #   4. Pago quebrado con tolerancia (hasta 1.00): queda con observación por la diferencia
    #   5. Total del proveedor
    uno_a_uno(UMBRAL_OK)
    agrupados(UMBRAL_OK)
    uno_a_uno(TOLERANCIA)
    agrupados(TOLERANCIA)
    por_total_del_proveedor()

    # Lo que quedó sin cruzar
    for rfc, lineas in pendientes.items():
        disponibles = libres(rfc)
        pendiente_txt = sum((lineas_txt[n]["Importe"] for n in lineas), Decimal("0"))
        pendiente_xml = sum((facturas[i]["ImporteConIVA"] for i in disponibles), Decimal("0"))
        for n in lineas:
            imp = lineas_txt[n]["Importe"]
            if disponibles:
                i = min(disponibles, key=lambda i: abs(facturas[i]["ImporteConIVA"] - imp))
                filas[n]["Estatus"] = "Diferencia de importe"
                obs[n].append(f"XML más cercano: {facturas[i]['ArchivoXML']} por "
                              f"{facturas[i]['ImporteConIVA']:,.2f} "
                              f"(diferencia {facturas[i]['ImporteConIVA'] - imp:,.2f})")
                obs[n].append(f"Pendiente del proveedor: TXT {pendiente_txt:,.2f} en {len(lineas)} línea(s) "
                              f"contra XML {pendiente_xml:,.2f} en {len(disponibles)} factura(s) "
                              f"(diferencia {pendiente_xml - pendiente_txt:,.2f})")
            else:
                filas[n]["Estatus"] = "Sin XML"

    # Por cada factura (o grupo de facturas) cruzada: diferencia de importe e impuestos.
    # En pago quebrado todas las líneas del grupo comparten el resultado.
    grupos = {}
    for n, fila in enumerate(filas):
        if "_grupo" in fila:
            grupos.setdefault(fila["_grupo"], []).append(n)
    for grupo, ns in grupos.items():
        dif = (sum(facturas[i]["ImporteConIVA"] for i in grupo)
               - sum(lineas_txt[n]["Importe"] for n in ns))
        if DIFERENCIA_IMPORTE_ES_OBSERVACION and abs(dif) > UMBRAL_OK:
            for n in ns:
                obs[n].append(f"Diferencia de importe de {dif:,.2f} entre factura y TXT "
                              f"(dentro de la tolerancia de {TOLERANCIA})")
        con_importes = next(n for n in ns if filas[n].get("SubTotal") is not None)
        motivo = revisar_impuestos(filas[con_importes])
        codigos = sorted({lineas_txt[n]["CodigoImpuesto"] for n in ns})
        if len(codigos) > 1:
            nums = ", ".join(str(lineas_txt[n]["Linea"]) for n in ns)
            motivo = "; ".join(filter(None, [motivo, f"Las líneas {nums} pagan la misma factura con códigos de "
                                                     f"impuesto distintos ({', '.join(codigos)})"]))
        for n in ns:
            pasadas[n].add("impuestos")
            if motivo:
                deten[n].append(motivo)

    # Pago doble real: el mismo UUID quedó usado en más de un XML cruzado (dos líneas pagan la misma factura)
    usos_uuid = {}
    for n, fila in enumerate(filas):
        for i in fila.get("_grupo", ()):
            if facturas[i]["UUID"]:
                usos_uuid.setdefault(facturas[i]["UUID"], set()).add(i)
    for n, fila in enumerate(filas):
        dobles = [u for u, idxs in usos_uuid.items()
                  if len(idxs) > 1 and any(facturas[i]["UUID"] == u for i in fila.get("_grupo", ()))]
        if dobles:
            lineas_doble = sorted({lineas_txt[m]["Linea"] for m, x in enumerate(filas)
                                   if any(facturas[i]["UUID"] in dobles for i in x.get("_grupo", ()))})
            deten[n].append(f"Pago doble: la misma factura (UUID {', '.join(dobles)}) la pagan las líneas "
                            f"{', '.join(map(str, lineas_doble))} del TXT. {MSG_UUID_REPETIDO}")

    # Validación contra SAP
    for n, fila in enumerate(filas):
        if "_grupo" not in fila:
            continue
        if sap is None:
            continue                                   # no se pidió validar SAP en esta corrida
        consulta = sap.get(fila["IDProveedor"])
        if consulta is None:
            continue                                   # no se consultó: queda como validación no realizada
        codigos, error = consulta
        fila["CodigosSAP"] = ", ".join(codigos)
        if error:
            obs[n].append(f"No se pudo validar en SAP: {error}")
            continue
        fila["RegimenSAP"], o, d = validar_sap(lineas_txt[n]["CodigoImpuesto"], fila.get("RegimenEmisor", ""),
                                               codigos, fila.get("_iva_xml", False))
        obs[n].extend(o)
        deten[n].extend(d)
        pasadas[n].add("sap")

    # Estatus final. "Cruzado" solo si no hay rechazos, ni motivos para detener, ni observaciones
    # que requieran revisión, y además se aprobaron TODAS las validaciones de la lista.
    no_pedidas = {clave for clave, datos in (("sap", sap), ("estado_sat", estados_sat)) if datos is None}
    for n, (fila, o, r, d) in enumerate(zip(filas, obs, rech, deten)):
        cruzada = fila["Estatus"].startswith("Cruzado")
        if cruzada:
            faltan = [nombre for clave, nombre in VALIDACIONES.items()
                      if clave not in pasadas[n] and clave not in no_pedidas]
            if faltan and not r:                     # si ya se rechaza, el faltante no aporta
                o.append("Validaciones no realizadas: " + ", ".join(faltan))
        fila["DiscrepanciaSAP"] = any("SAP" in x and not x.startswith(NO_ES_DISCREPANCIA_SAP) for x in o + d)
        if cruzada and r:
            fila["Estatus"] = "Rechazar"
        elif cruzada and d:
            fila["Estatus"] = "Detenida"
        elif cruzada and any(not x.startswith(INFORMATIVAS) for x in o):
            fila["Estatus"] = "Cruzado con observaciones"
        elif not cruzada and d and lineas_txt[n]["Importe"] <= 0:
            fila["Estatus"] = "Detenida"             # el propio TXT está mal; en los demás casos se conserva
                                                     # el estatus (Sin XML, etc.) y el motivo va en MotivoRechazo
        fila["MotivoRechazo"] = "; ".join(dict.fromkeys(r + d))
        fila["Observaciones"] = "; ".join(dict.fromkeys(o))

    # XML que ninguna línea del TXT usó
    rfc_a_id = {}
    for i, r in zip(catalogo["IDProveedor"], catalogo["RFC"]):
        if r:
            rfc_a_id.setdefault(r, []).append(i)
    sobrantes = []
    for i, f in enumerate(facturas):
        if i in usadas:
            continue
        s = {k: f[k] for k in ("ArchivoXML", "UUID", "FechaXML", "RFCEmisor", "NombreEmisor",
                               "RegimenEmisor", "RFCReceptor", "MetodoPago", "FormaPago",
                               "SubTotal", "Descuento", "IVA", "RetISR", "RetIVA", "Total",
                               "ImporteConIVA")}
        s["NombreRegimen"] = REGIMENES.get(f["RegimenEmisor"], "Desconocido")
        if estados_sat is not None:
            s["EstadoSAT"] = describir_estado_sat(estados_sat.get(f["UUID"]))
        r, o_xml, d, _ = validacion[i]
        motivos = r + d
        copia = sorted(facturas[j]["ArchivoXML"] for j in usadas
                       if j != i and f["UUID"] and facturas[j]["UUID"] == f["UUID"])
        if copia:                                     # la otra copia ya se cruzó: esta sobra
            motivos.append(f"Copia repetida de {', '.join(copia)} (ya cruzado): {MSG_UUID_REPETIDO}")
        else:
            motivos += [x for x in o_xml if x.startswith(MSG_UUID_REPETIDO)]
        if cancelada(i):
            nueva = [facturas[j]["ArchivoXML"] for j in sustitutas(i, sorted(usadas))]
            motivos.append(f"Ya se cruzó la factura que la sustituye (mismo RFC e importe): {', '.join(nueva)}"
                           if nueva else MSG_FACTURA_CANCELADA)
        s["MotivoRechazo"] = "; ".join(motivos)
        s["IDProveedorCatalogo"] = ", ".join(rfc_a_id.get(f["RFCEmisor"], ["No está en catálogo"]))
        sobrantes.append(s)

    verificar_integridad(lineas_txt, facturas, filas, sobrantes)
    return filas, sobrantes


def verificar_integridad(lineas_txt, facturas, filas, sobrantes):
    """Controles de cuadre del propio programa. Si alguno falla es un error de lógica, no de los datos:
    se detiene ANTES de generar un reporte o tocar el historial."""
    fallas = []
    if sorted(f["Linea"] for f in filas) != sorted(l["Linea"] for l in lineas_txt):
        fallas.append("el resultado no tiene exactamente una fila por cada línea del TXT")
    grupos = {}
    for n, fila in enumerate(filas):
        if "_grupo" in fila:
            grupos.setdefault(fila["_grupo"], []).append(n)
    usados = [i for g in grupos for i in g]
    if len(usados) != len(set(usados)):
        fallas.append("un mismo XML quedó asignado a dos cruces distintos")
    if len(usados) + len(sobrantes) != len(facturas):
        fallas.append(f"XML usados ({len(usados)}) + sin TXT ({len(sobrantes)}) no suman los leídos "
                      f"({len(facturas)})")
    for g, ns in grupos.items():
        dif = (sum(facturas[i]["ImporteConIVA"] for i in g) - sum(lineas_txt[n]["Importe"] for n in ns))
        if abs(dif) > TOLERANCIA:
            fallas.append(f"el cruce de las líneas {[lineas_txt[n]['Linea'] for n in ns]} difiere {dif}")
        if len({lineas_txt[n]["IDProveedor"] for n in ns}) != 1 \
                or {facturas[i]["RFCEmisor"] for i in g} != {filas[ns[0]]["RFC"]}:
            fallas.append(f"el cruce de las líneas {[lineas_txt[n]['Linea'] for n in ns]} mezcla proveedores")
    for fila in filas:
        if fila["Estatus"] in ("Cruzado", "Cruzado (pago quebrado)") \
                and fila["MotivoRechazo"]:
            fallas.append(f"la línea {fila['Linea']} quedó Cruzado con motivo de rechazo")
    if fallas:
        raise RuntimeError("Falló el control de integridad de la conciliación:\n  " + "\n  ".join(fallas))
