"""TXT de pagos para SAP: leerlo tal cual, revisar sus problemas y volver a escribirlo con exactamente el mismo
formato (11 columnas separadas por tabulador, sin encabezado, misma codificación y mismos saltos de línea).

Cada valor se guarda como texto, sin recortar ni reformatear: lo que el usuario no editó sale idéntico,
carácter por carácter. Así el archivo descargado sigue el flujo de carga a SAP sin sorpresas."""

import re
from collections import defaultdict
from dataclasses import dataclass, field

from .config import COLUMNAS_TXT, CODIGOS_IMPUESTO_TXT, TOLERANCIA
from .entradas import es_encabezado_txt, leer_linea_txt
from .utilidades import ErrorUsuario, normalizar_id

# Nombres que ve el usuario (mismo orden que COLUMNAS_TXT y que el archivo)
COLUMNAS = ["Fecha", "IDProveedor", "Folio", "Monto", "X", "Tax", "1", "Centro de Costo", "Cuenta Contable",
            "PM", "Referencia"]
# Lo que Contabilidad puede editar. Fecha, X, 1 y PM son campos fijos que pide SAP; el Folio y la Referencia solo
# cambian con "Formatear folios" (folios únicos, ver folios.py; la Referencia empieza con el folio).
EDITABLES = ("IDProveedor", "Monto", "Tax", "Centro de Costo", "Cuenta Contable")
# Separa el folio del resto de la Referencia al formatear: "A1893      Q18-2026" -> "7KQ2M9XHRT_Q18-2026".
SEPARADOR_REFERENCIA = "_"
assert len(COLUMNAS) == len(COLUMNAS_TXT)


@dataclass
class TxtSap:
    lineas: dict                    # número de línea del archivo -> lista de 11 valores (texto tal cual)
    codificacion: str = "utf-8"
    bom: bool = False
    salto: str = "\r\n"
    salto_final: bool = True
    estructura: dict = field(default_factory=dict)   # número de línea -> error de columnas (no se puede descargar)


def leer(datos):
    """bytes del TXT -> TxtSap. Los números de línea son los del archivo (los mismos que usa la conciliación)."""
    bom = datos.startswith(b"\xef\xbb\xbf")
    for codificacion in ("utf-8", "latin-1"):
        try:
            texto = datos[3 if bom else 0:].decode(codificacion)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ErrorUsuario("No se pudo leer el TXT.")
    txt = TxtSap({}, codificacion, bom, "\r\n" if "\r\n" in texto else "\n", texto.endswith(("\n", "\r")))
    primera = True
    for n, linea in enumerate(texto.splitlines(), start=1):
        if not linea.strip():
            continue                                   # renglones vacíos: no van al TXT de SAP
        campos = linea.split("\t")
        if primera:
            primera = False
            if es_encabezado_txt(campos):
                continue                               # el encabezado no va al TXT de SAP
        if len(campos) != len(COLUMNAS):
            txt.estructura[n] = (f"Trae {len(campos)} columnas y se esperaban {len(COLUMNAS)}: "
                                 "corrígela en el TXT original")
            campos = (campos + [""] * len(COLUMNAS))[:len(COLUMNAS)]
        txt.lineas[n] = campos
    if not txt.lineas:
        raise ErrorUsuario("El TXT no tiene líneas.")
    return txt


def escribir(txt, lineas=None):
    """TxtSap (y, si se dan, los valores editados) -> bytes listos para SAP, con el formato original."""
    lineas = txt.lineas if lineas is None else lineas
    texto = txt.salto.join("\t".join(lineas[n]) for n in sorted(lineas))
    if txt.salto_final:
        texto += txt.salto
    return (b"\xef\xbb\xbf" if txt.bom else b"") + texto.encode(txt.codificacion)


def lineas_para_el_motor(lineas):
    """Valores editados -> (líneas como las lee la conciliación, {línea: error de formato})."""
    filas, errores = [], {}
    for n in sorted(lineas):
        fila, error = leer_linea_txt(lineas[n], n)
        if error:
            errores[n] = error
        else:
            filas.append(fila)
    return filas, errores


def problemas(txt, lineas, ids_catalogo, periodo=None, folios_registrados=None):
    """Problemas del propio TXT, por línea y columna: [(línea, columna, texto, bloquea_descarga)].
    bloquea_descarga=True: SAP lo rechazaría (estructura, fecha, importe, caracteres)."""
    folios_registrados = folios_registrados or {}
    salida = [(n, "", texto, True) for n, texto in txt.estructura.items()]
    por_folio = defaultdict(list)
    for n in sorted(lineas):
        if n in txt.estructura:
            continue
        v = dict(zip(COLUMNAS, (c.strip() for c in lineas[n])))
        fila, error = leer_linea_txt(lineas[n], n)
        if error and "fecha" in error:
            salida.append((n, "Fecha", error[0].upper() + error[1:], True))
        elif error:
            salida.append((n, "Monto", error[0].upper() + error[1:], True))
        if fila and periodo and fila["_fecha"].strftime("%Y-%m") != periodo:
            salida.append((n, "Fecha", f"La fecha {v['Fecha']} no es del mes que se concilia", False))
        id_prov = normalizar_id(v["IDProveedor"])
        if not id_prov:
            salida.append((n, "IDProveedor", "Falta el IDProveedor", False))
        elif ids_catalogo is not None and id_prov not in ids_catalogo:
            salida.append((n, "IDProveedor", f"El ID {id_prov} no está en el catálogo de proveedores", False))
        if v["Tax"].upper() not in CODIGOS_IMPUESTO_TXT:
            salida.append((n, "Tax", f"Código '{v['Tax']}' desconocido: debe ser E0 (sin IVA) o F2 (con IVA)", False))
        if not v["Folio"]:
            salida.append((n, "Folio", "Falta el folio: usa \"Formatear folios\"", False))
        else:
            por_folio[(id_prov, v["Folio"])].append(n)
            previo = folios_registrados.get(v["Folio"])
            if previo and previo.get("Periodo") and previo["Periodo"] != periodo:
                salida.append((n, "Folio", f"El folio {v['Folio']} ya se usó en {previo['Periodo']}: "
                                           "SAP lo rechazaría. Usa \"Formatear folios\"", False))
        for columna, valor in zip(COLUMNAS, lineas[n]):
            try:
                valor.encode(txt.codificacion)
            except UnicodeEncodeError:
                salida.append((n, columna, "Tiene un carácter que el TXT no admite (¿acento raro o símbolo?)", True))
    for (id_prov, folio), ns in por_folio.items():
        if len(ns) > 1:
            for n in ns:
                otras = ", ".join(str(m) for m in ns if m != n)
                salida.append((n, "Folio", f"Folio {folio} repetido para el mismo proveedor (líneas {otras}): "
                                           "SAP lo rechazaría. Usa \"Formatear folios\"", False))
    return salida


def sugerencias_de_id(resultado, sobrantes, catalogo_nombres):
    """Línea sin factura + factura sin línea por el mismo importe -> POSIBLE ID equivocado en el TXT. Es solo apoyo:
    el portal nunca lo cambia solo (SAP le deposita a la cuenta del ID que va en el TXT; quien edita lo verifica).
    Devuelve {línea: (ID de la factura, texto)}."""
    huerfanas = [s for s in sobrantes if re.fullmatch(r"[^,]+", s.get("IDProveedorCatalogo", ""))
                 and s["IDProveedorCatalogo"] != "No está en catálogo"]
    salida = {}
    for fila in resultado:
        if fila["Estatus"] not in ("Sin XML", "Diferencia de importe", "ID no está en catálogo",
                                   "Proveedor sin RFC en catálogo", "RFC inválido en catálogo"):
            continue
        candidatas = [s for s in huerfanas if abs(s["ImporteConIVA"] - fila["Importe"]) <= TOLERANCIA
                      and s["IDProveedorCatalogo"] != fila["IDProveedor"]]
        ids = {s["IDProveedorCatalogo"] for s in candidatas}
        if len(ids) == 1:                               # una sola explicación: si hay varias, no se adivina
            s = candidatas[0]
            nuevo = s["IDProveedorCatalogo"]
            actual = fila["IDProveedor"] or "(vacío)"
            salida[fila["Linea"]] = (nuevo, f"Posible ID equivocado: la factura {s['ArchivoXML']} de "
                                            f"{catalogo_nombres.get(nuevo, s['NombreEmisor'])} (ID {nuevo} en el "
                                            f"catálogo) es por ${s['ImporteConIVA']:,.2f} y no tiene línea en el TXT. "
                                            f"Verifica en SAP quién es el ID {actual} antes de cambiar nada.")
    return salida


def referencia_con_folio(referencia, folio_actual, folio_nuevo):
    """La Referencia empieza con el folio ("A1893      Q18-2026"). Al formatear folios queda el folio nuevo y lo demás
    que escribió quien manda el TXT, unidos con guion bajo y sin espacios: "7KQ2M9XHRT_Q18-2026". Si la referencia
    no empezaba con el folio, el folio nuevo se le antepone igual; si estaba vacía, queda solo el folio.
    Formatear otra vez da lo mismo (no se acumulan guiones)."""
    referencia = referencia.strip()
    resto = referencia[len(folio_actual):] if folio_actual and referencia.startswith(folio_actual) else referencia
    partes = resto.replace(SEPARADOR_REFERENCIA, " ").split()
    return SEPARADOR_REFERENCIA.join([folio_nuevo] + partes)
