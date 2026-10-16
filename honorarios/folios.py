"""Folios únicos y rastreables para el TXT de SAP (botón "Formatear folios" de Generar TXT).

SAP rechaza una línea si el proveedor repite un folio que ya usó, aunque haya sido en otro mes. El folio que pone el
portal mide 16 caracteres (lo más que acepta SAP) y se lee así:

    26 18 5113 K7Q3MXP4
    │  │   │      └─ 8 caracteres calculados con la factura (huella SHA-256 de IDProveedor + folio original +
    │  │   │         monto + fecha): misma factura -> mismo folio, siempre y en cualquier computadora
    │  │   └─ últimos 4 caracteres del IDProveedor
    │  └─ quincena de la fecha (Q18: del 16 al fin de septiembre; Q01 es del 1 al 15 de enero)
    └─ año (2026)

- Rastreable: con solo verlo se sabe la quincena y el proveedor; HistorialFolios.csv guarda además el folio original,
  el monto, la fecha, quién lo generó y cuándo.
- Misma factura -> mismo folio: si una línea ya cargada se vuelve a mandar, SAP la rechaza por folio repetido (no se
  paga dos veces). Un folio que ya es de este formato para ese proveedor se conserva tal cual.
- Pago quebrado con líneas idénticas: cada repetición agrega su número de parte, así no chocan entre sí.
- 32^8 ≈ 1.1 × 10^12 combinaciones por proveedor y quincena (los 8 caracteres usan 32 símbolos, sin 0/O ni 1/I).
  Además, antes de aceptar un folio se comprueba contra los demás del archivo y contra HistorialFolios.csv; si
  chocara con el de otra factura, se recalcula.
"""

import getpass
import hashlib
import re
from collections import Counter
from datetime import datetime
from decimal import Decimal

import pandas as pd

from .config import ARCHIVO_FOLIOS, CONFIG, LONGITUD_FOLIO
from .utilidades import guardar_seguro, normalizar_id

# 32 símbolos sin los que se confunden al leer o dictar: sin 0/O ni 1/I.
ALFABETO = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"
LARGO_HUELLA = LONGITUD_FOLIO - 8                     # año (2) + quincena (2) + IDProveedor (4) + huella
COLUMNAS = ["Folio", "Clave", "IDProveedor", "Periodo", "Linea", "Registrado", "FolioOriginal", "Monto", "Fecha",
            "Usuario", "Archivo"]


def huella(clave, longitud=LARGO_HUELLA):
    """Huella de la clave en el alfabeto de 32 símbolos (determinista: misma clave -> misma huella)."""
    numero = int.from_bytes(hashlib.sha256(clave.encode("utf-8")).digest(), "big")
    simbolos = []
    for _ in range(longitud):
        numero, resto = divmod(numero, len(ALFABETO))
        simbolos.append(ALFABETO[resto])
    return "".join(simbolos)


def quincena(fecha):
    """(año, quincena 1-24) de una fecha DD.MM.AAAA: del 1 al 15 es la primera del mes, del 16 al fin la segunda."""
    dia = datetime.strptime(fecha, "%d.%m.%Y")
    return dia.year, (dia.month - 1) * 2 + (1 if dia.day <= 15 else 2)


def texto_quincena(fecha):
    """"Q18-2026": como la escriben en la Referencia del TXT."""
    anio, q = quincena(fecha)
    return f"Q{q:02d}-{anio}"


def cola_del_id(id_proveedor):
    limpio = re.sub(r"[^0-9A-Z]", "", normalizar_id(id_proveedor).upper())
    return limpio[-4:].rjust(4, "0")


def prefijo(id_proveedor, fecha):
    anio, q = quincena(fecha)
    return f"{anio % 100:02d}{q:02d}{cola_del_id(id_proveedor)}"


def es_folio_rastreable(folio, id_proveedor):
    """True si el folio ya tiene este formato para ese proveedor (se conserva al volver a formatear)."""
    return (len(folio) == LONGITUD_FOLIO and folio[:4].isdigit() and 1 <= int(folio[2:4]) <= 24
            and folio[4:8] == cola_del_id(id_proveedor) and all(c in ALFABETO for c in folio[8:]))


def clave_de(linea):
    """Lo que identifica a la factura en la línea (no cambia al volver a armar el Excel)."""
    return "|".join([normalizar_id(linea["IDProveedor"]), linea["Folio"].strip().upper(), linea["Monto"],
                     linea["Fecha"]])


def formatear_folios(lineas, registrados):
    """lineas: [{"IDProveedor", "Folio", "Monto", "Fecha"}] (valores ya como van en el TXT).
    registrados: cargar_folios(). Devuelve, en el mismo orden, [(folio, clave, registro previo o None)].
    Ningún folio se repite entre sí ni con uno registrado de otra factura."""
    tomados = {f: r.get("Clave", "") for f, r in registrados.items()}
    partes, salida = {}, []
    for linea in lineas:
        if es_folio_rastreable(linea["Folio"], linea["IDProveedor"]):
            folio = linea["Folio"]                         # ya formateado: se conserva (misma factura, mismo folio)
            clave = tomados.get(folio) or f"conservado:{folio}"
        else:
            base = clave_de(linea)
            partes[base] = partes.get(base, 0) + 1
            clave = f"{base}#{partes[base]}"
            inicio = prefijo(linea["IDProveedor"], linea["Fecha"])
            sal, folio = "", inicio + huella(clave)
            while tomados.get(folio, clave) != clave:      # prácticamente imposible: se recalcula con sal
                sal += "+"
                folio = inicio + huella(clave + sal)
        previo = registrados.get(folio) if registrados.get(folio, {}).get("Clave") == clave else None
        tomados[folio] = clave
        salida.append((folio, clave, previo))
    return salida


def ya_generados(lineas, registrados=None):
    """Líneas del TXT original que ya salieron en un TXT descargado de Generar TXT.
    lineas: filas de la conciliación ({"Linea", "IDProveedor", "Folio", "Importe", "Fecha"}).
    -> {Linea: "16/10/2026 11:05 · Aprobadas.xlsx"}. Se reconocen por la misma clave que da el folio (ID, folio
    original, monto y fecha); si el TXT trae líneas idénticas, la primera se compara con la primera, etc."""
    registrados = cargar_folios() if registrados is None else registrados
    por_clave = {r.get("Clave", ""): r for r in registrados.values()}
    vistas, salida = Counter(), {}
    for linea in sorted(lineas, key=lambda l: l["Linea"]):
        try:
            monto = f"{Decimal(str(linea['Importe'])):.2f}"
        except Exception:                                  # noqa: BLE001 - importe inválido: no pudo ir en un TXT
            continue
        base = clave_de({"IDProveedor": str(linea["IDProveedor"]), "Folio": str(linea["Folio"]), "Monto": monto,
                         "Fecha": str(linea["Fecha"])})
        vistas[base] += 1
        registro = por_clave.get(f"{base}#{vistas[base]}")
        if registro:
            salida[linea["Linea"]] = " · ".join(x for x in (registro.get("Registrado", ""),
                                                            registro.get("Archivo", "")) if x)
    return salida


def cargar_folios():
    """{folio: registro} de los folios ya generados en esta computadora."""
    ruta = CONFIG / ARCHIVO_FOLIOS
    if not ruta.exists():
        return {}
    datos = pd.read_csv(ruta, dtype=str, encoding="utf-8-sig").fillna("")
    return {r["Folio"]: r for r in datos.to_dict("records")}


def registrar_folios(registros):
    """Guarda los folios del TXT descargado (Folio, Clave, IDProveedor, Periodo, Linea, FolioOriginal, Monto, Fecha,
    Archivo). Un folio ya registrado con la misma clave no se duplica (conserva la fecha en que se generó)."""
    if not registros:
        return
    ruta = CONFIG / ARCHIVO_FOLIOS
    ahora, usuario = datetime.now().strftime("%d/%m/%Y %H:%M"), getpass.getuser()
    nuevos = pd.DataFrame([{**r, "Registrado": ahora, "Usuario": usuario} for r in registros],
                          columns=COLUMNAS).astype(str)
    previo = (pd.read_csv(ruta, dtype=str, encoding="utf-8-sig").fillna("") if ruta.exists()
              else pd.DataFrame(columns=COLUMNAS))
    total = (pd.concat([previo, nuevos], ignore_index=True).fillna("")
             .drop_duplicates(["Folio", "Clave"]).reindex(columns=COLUMNAS, fill_value=""))
    ruta.parent.mkdir(parents=True, exist_ok=True)
    guardar_seguro(ruta, lambda tmp: total.to_csv(tmp, index=False, encoding="utf-8-sig"))
