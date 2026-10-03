"""Folios únicos para el TXT de SAP.

SAP rechaza el TXT si un proveedor repite un folio que ya usó, aunque haya sido en otro mes. Antes los folios se
capturaban a mano (consecutivos reciclados) y chocaban. Aquí el folio NO es aleatorio: es una huella (SHA-256) del
UUID de la factura que cruzó con la línea:

- misma factura -> mismo folio siempre (reprocesar el mes no inventa folios nuevos, y si alguien intentara pagar
  otra vez la misma factura, SAP también lo rechazaría);
- facturas distintas -> folios distintos sin depender de qué computadora los generó;
- pago quebrado (varias líneas, una factura): cada línea agrega su número de parte, así no chocan entre sí;
- línea sin factura (sin UUID): código aleatorio criptográfico (módulo secrets).

Con LONGITUD_FOLIO = 10 y 32 símbolos hay 32^10 ≈ 1.1 × 10^15 folios (50 bits). Además, antes de aceptar un folio
se comprueba contra HistorialFolios.csv y contra los demás folios del archivo; si chocara, se recalcula.
"""

import hashlib
import secrets
from datetime import datetime

import pandas as pd

from .config import ARCHIVO_FOLIOS, CONFIG, LONGITUD_FOLIO
from .utilidades import guardar_seguro

# 32 símbolos sin los que se confunden al leer o dictar: sin 0/O ni 1/I.
ALFABETO = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"
COLUMNAS = ["Folio", "Clave", "IDProveedor", "Periodo", "Linea", "Registrado"]


def folio_de(clave, longitud=LONGITUD_FOLIO):
    """Huella de la clave en el alfabeto de 32 símbolos (determinista: misma clave -> mismo folio)."""
    numero = int.from_bytes(hashlib.sha256(clave.encode("utf-8")).digest(), "big")
    simbolos = []
    for _ in range(longitud):
        numero, resto = divmod(numero, len(ALFABETO))
        simbolos.append(ALFABETO[resto])
    return "".join(simbolos)


def folio_aleatorio(longitud=LONGITUD_FOLIO):
    return "".join(secrets.choice(ALFABETO) for _ in range(longitud))


def es_folio_generado(folio, longitud=LONGITUD_FOLIO):
    return len(folio) == longitud and all(c in ALFABETO for c in folio)


def cargar_folios():
    """{folio: registro} de los folios ya descargados en esta computadora."""
    ruta = CONFIG / ARCHIVO_FOLIOS
    if not ruta.exists():
        return {}
    datos = pd.read_csv(ruta, dtype=str, encoding="utf-8-sig").fillna("")
    return {r["Folio"]: r for r in datos.to_dict("records")}


def asignar_folios(lineas, registrados, actuales=None, longitud=LONGITUD_FOLIO):
    """lineas: [(Linea, UUIDs de la factura con que cruzó (tupla, vacía si no cruzó))].
    registrados: cargar_folios(). actuales: {Linea: folio que tiene ahora} (para conservar los aleatorios ya dados).
    Devuelve {Linea: (folio, clave)}. Ningún folio se repite entre sí ni con uno registrado de otra clave."""
    actuales = actuales or {}
    partes = {}
    for linea, uuids in sorted(lineas):
        if uuids:
            partes.setdefault(tuple(sorted(uuids)), []).append(linea)
    tomados = {f: r.get("Clave", "") for f, r in registrados.items()}
    asignados = {}

    def libre(folio, clave):
        return tomados.get(folio, clave) == clave

    for uuids, de_la_factura in partes.items():
        for parte, linea in enumerate(de_la_factura):
            clave = f"{'|'.join(uuids)}#{parte}"
            sal, folio = "", folio_de(clave, longitud)
            while not libre(folio, clave):                 # prácticamente imposible: se recalcula con sal
                sal += "+"
                folio = folio_de(clave + sal, longitud)
            tomados[folio] = clave
            asignados[linea] = (folio, clave)
    for linea, uuids in sorted(lineas):
        if uuids:
            continue
        actual = actuales.get(linea, "")
        clave = f"aleatorio:{actual}"
        if es_folio_generado(actual, longitud) and libre(actual, clave):
            folio = actual                                 # ya se le había dado uno: se conserva
        else:
            folio = folio_aleatorio(longitud)
            clave = f"aleatorio:{folio}"
            while not libre(folio, clave):
                folio = folio_aleatorio(longitud)
                clave = f"aleatorio:{folio}"
        tomados[folio] = clave
        asignados[linea] = (folio, clave)
    return asignados


def registrar_folios(registros):
    """Guarda los folios descargados (Folio, Clave, IDProveedor, Periodo, Linea). Un folio ya registrado con la
    misma clave no se duplica."""
    if not registros:
        return
    ruta = CONFIG / ARCHIVO_FOLIOS
    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    nuevos = pd.DataFrame([{**r, "Registrado": ahora} for r in registros], columns=COLUMNAS).astype(str)
    previo = (pd.read_csv(ruta, dtype=str, encoding="utf-8-sig").fillna("") if ruta.exists()
              else pd.DataFrame(columns=COLUMNAS))
    total = pd.concat([previo, nuevos], ignore_index=True).drop_duplicates(["Folio", "Clave"])
    ruta.parent.mkdir(parents=True, exist_ok=True)
    guardar_seguro(ruta, lambda tmp: total.to_csv(tmp, index=False, encoding="utf-8-sig"))
