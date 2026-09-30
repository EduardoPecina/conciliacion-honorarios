"""
Genera una carpeta "Honorarios" completa con datos INVENTADOS para probar el programa.
Ningún RFC, nombre ni importe corresponde a personas o facturas reales.

Cada caso está pensado para ejercitar una regla de negocio y tiene su resultado esperado
en CASOS_ESPERADOS (IDProveedor -> Estatus).

Uso manual:
    python tests/datos_prueba.py C:/ruta/destino
"""

import sys
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.table import Table

RFC_RECEPTOR = "SIN070530EI5"
NOMBRE_RECEPTOR = "SERVICIOS INTEGRALES NOVA DE MONTERREY"
ANIO, MES = 2026, 9

D = Decimal
C = D("0.01")


def r2(v):
    return v.quantize(C, rounding=ROUND_HALF_UP)


# (ID, nombre, RFC, activo). La columna Activo del catálogo es histórica y el programa la ignora.
PROVEEDORES = [
    ("100001", "MEDICO RESICO UNO", "MERU800101AB1", 1),
    ("100002", "NUTRIOLOGA IVA DOS", "NUID800202AB2", 1),
    ("100003", "MEDICO QUEBRADO TRES", "MEQT800303AB3", 1),
    ("100004", "MEDICO DOS FACTURAS", "MEDF800404AB4", 1),
    ("100005", "MEDICO SIN FACTURA", "MESF800505AB5", 1),
    ("100006", "MEDICO DIFERENCIA", "MEDI800606AB6", 1),
    ("100007", "MEDICO BLOQUEADO", "MEBL800707AB7", "Baja por revisar"),   # valor no reconocido: se ignora
    ("100008", "MEDICO AGOSTO", "MEAG800808AB8", 1),
    ("100009", "MEDICO EFECTIVO", "MEEF800909AB9", 1),
    ("100010", "MEDICO PPD", "MEPP801010AC1", 1),
    ("100011", "MEDICO OTRO RECEPTOR", "MEOR801111AC2", 1),
    ("100012", "MEDICO PUE 99", "MEPN801212AC3", 1),
    ("100013", "MEDICO NO RESICO", "MENR800113AC4", 1),
    ("PV0001", "PROVEEDOR SIN RFC", "", 1),
]

# ID -> Estatus esperado (la base para las pruebas automáticas)
CASOS_ESPERADOS = {
    "100001": "Cruzado",                        # E0 RESICO 1:1
    "100002": "Cruzado",                        # F2 con IVA, régimen 612
    "100003": "Cruzado (pago quebrado)",        # 3 líneas contra 1 factura
    "100004": "Cruzado (pago quebrado)",        # 1 línea contra 2 facturas
    "100005": "Sin XML",
    "100006": "Diferencia de importe",
    "100007": "Cruzado",                        # Activo=0 en catálogo: se ignora, no detiene
    "100008": "Rechazar",                       # factura fuera del periodo
    "100009": "Rechazar",                       # forma de pago 01 efectivo
    "100010": "Cruzado",                        # PPD + 99 (con nota informativa)
    "100011": "Rechazar",                       # facturada a otro RFC
    "100012": "Cruzado",                        # PUE + 99
    "100013": "Cruzado",                        # E0 no RESICO (612)
    "PV0001": "Proveedor sin RFC en catálogo",
    "999999": "ID no está en catálogo",
}


def xml_cfdi(rfc, subtotal, regimen="626", con_iva=False, fecha=f"{ANIO}-{MES:02d}-07T09:00:00",
             metodo="PUE", forma="03", receptor=RFC_RECEPTOR, uuid="00000000-0000-0000-0000-000000000000",
             cp="66450"):
    """CFDI 4.0 de honorarios con impuestos calculados como los calcularía el emisor."""
    sub = D(subtotal)
    iva = r2(sub * D("0.16")) if con_iva else D("0")
    ret_isr = r2(sub * (D("0.0125") if regimen == "626" else D("0.10")))
    ret_iva = r2(iva * 2 / 3) if con_iva else D("0")
    total = sub + iva - ret_isr - ret_iva
    tasa_isr = "0.012500" if regimen == "626" else "0.100000"
    if con_iva:
        tras_c = f'<cfdi:Traslado Base="{sub}" Impuesto="002" TipoFactor="Tasa" TasaOCuota="0.160000" Importe="{iva}" />'
        tras_g = f'<cfdi:Traslado Base="{sub}" Impuesto="002" TipoFactor="Tasa" TasaOCuota="0.160000" Importe="{iva}" />'
        ret_iva_c = (f'<cfdi:Retencion Base="{sub}" Impuesto="002" TipoFactor="Tasa" '
                     f'TasaOCuota="0.106667" Importe="{ret_iva}" />')
        ret_iva_g = f'<cfdi:Retencion Impuesto="002" Importe="{ret_iva}" />'
        tot_tras = f' TotalImpuestosTrasladados="{iva}"'
    else:
        tras_c = tras_g = f'<cfdi:Traslado Base="{sub}" Impuesto="002" TipoFactor="Exento" />'
        ret_iva_c = ret_iva_g = tot_tras = ""
    return f'''<?xml version="1.0" encoding="utf-8"?>
<cfdi:Comprobante Version="4.0" Fecha="{fecha}" FormaPago="{forma}" MetodoPago="{metodo}" SubTotal="{sub}" Moneda="MXN" Total="{total}" TipoDeComprobante="I" Exportacion="01" LugarExpedicion="64000" xmlns:cfdi="http://www.sat.gob.mx/cfd/4">
  <cfdi:Emisor Rfc="{rfc}" Nombre="EMISOR DE PRUEBA" RegimenFiscal="{regimen}" />
  <cfdi:Receptor Rfc="{receptor}" Nombre="{NOMBRE_RECEPTOR}" DomicilioFiscalReceptor="{cp}" RegimenFiscalReceptor="601" UsoCFDI="G03" />
  <cfdi:Conceptos>
    <cfdi:Concepto ClaveProdServ="85121600" Cantidad="1" ClaveUnidad="E48" Descripcion="Honorarios" ValorUnitario="{sub}" Importe="{sub}" ObjetoImp="02">
      <cfdi:Impuestos>
        <cfdi:Traslados>{tras_c}</cfdi:Traslados>
        <cfdi:Retenciones><cfdi:Retencion Base="{sub}" Impuesto="001" TipoFactor="Tasa" TasaOCuota="{tasa_isr}" Importe="{ret_isr}" />{ret_iva_c}</cfdi:Retenciones>
      </cfdi:Impuestos>
    </cfdi:Concepto>
  </cfdi:Conceptos>
  <cfdi:Impuestos TotalImpuestosRetenidos="{ret_isr + ret_iva}"{tot_tras}>
    <cfdi:Retenciones><cfdi:Retencion Impuesto="001" Importe="{ret_isr}" />{ret_iva_g}</cfdi:Retenciones>
    <cfdi:Traslados>{tras_g}</cfdi:Traslados>
  </cfdi:Impuestos>
  <cfdi:Complemento>
    <tfd:TimbreFiscalDigital xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital" Version="1.1" UUID="{uuid}" />
  </cfdi:Complemento>
</cfdi:Comprobante>'''


def linea_txt(id_prov, folio, importe, codigo="E0", dia=29):
    return (f"{dia:02d}.{MES:02d}.{ANIO}\t{id_prov}\t{folio}\t{importe}\tX\t{codigo}\t1\t"
            f"32635079\t50E03158\tT\t{folio}      Q18-{ANIO}")


def crear(destino):
    """Crea la estructura Honorarios/ con catálogo, TXT y XML de prueba. Devuelve la ruta raíz."""
    raiz = Path(destino)
    cfg = raiz / "_Configuracion"
    mes = raiz / str(ANIO) / f"{MES:02d} Sep"
    xml = mes / "Entrada" / "XML"
    for carpeta in (cfg, xml / "subcarpeta", mes / "Salida"):
        carpeta.mkdir(parents=True, exist_ok=True)

    # Catálogo con la tabla "Catalogo"
    wb = Workbook()
    ws = wb.active
    ws.append(["IDProveedor", "Nombre", "RFC", "RFCTipo", "Activo", "Observaciones"])
    for id_, nombre, rfc, activo in PROVEEDORES:
        tipo = "VACIO" if not rfc else ("FISICA" if len(rfc) == 13 else "MORAL")
        ws.append([id_, nombre, rfc, tipo, activo, None])
    ws.add_table(Table(displayName="Catalogo", ref=f"A1:F{len(PROVEEDORES) + 1}"))
    wb.save(cfg / "CatalogoProveedoresTM26.xlsx")

    u = iter(f"{i:08X}-AAAA-4BBB-8CCC-{i:012X}" for i in range(1, 100))
    facturas = {
        "01_resico.xml": xml_cfdi("MERU800101AB1", "32188.00", uuid=next(u)),
        "02_f2_iva.xml": xml_cfdi("NUID800202AB2", "36658.00", regimen="612", con_iva=True, uuid=next(u)),
        "03_quebrado.xml": xml_cfdi("MEQT800303AB3", "1000.00", uuid=next(u)),
        "04a_dos.xml": xml_cfdi("MEDF800404AB4", "1000.00", uuid=next(u)),
        "subcarpeta/04b_dos.XML": xml_cfdi("MEDF800404AB4", "500.00", uuid=next(u)),
        "06_diferencia.xml": xml_cfdi("MEDI800606AB6", "30000.00", uuid=next(u)),
        "07_bloqueado.xml": xml_cfdi("MEBL800707AB7", "4000.00", uuid=next(u)),
        "08_agosto.xml": xml_cfdi("MEAG800808AB8", "1000.00", fecha=f"{ANIO}-08-30T10:00:00", uuid=next(u)),
        "09_efectivo.xml": xml_cfdi("MEEF800909AB9", "2000.00", forma="01", uuid=next(u)),
        "10_ppd.xml": xml_cfdi("MEPP801010AC1", "3000.00", metodo="PPD", forma="99", uuid=next(u)),
        "11_otro_receptor.xml": xml_cfdi("MEOR801111AC2", "4500.00", receptor="XAXX010101000", uuid=next(u)),
        "12_pue99.xml": xml_cfdi("MEPN801212AC3", "5500.00", forma="99", uuid=next(u)),
        "13_no_resico.xml": xml_cfdi("MENR800113AC4", "8000.00", regimen="612", uuid=next(u)),
        "20_sin_linea_txt.xml": xml_cfdi("ZZZZ800101ZZ1", "777.00", uuid=next(u)),
        "30_danado.xml": "<esto no es un xml",
        "notas.pdf": "archivo que no es XML: se ignora",
    }
    for nombre, contenido in facturas.items():
        (xml / nombre).write_text(contenido, encoding="utf-8")

    lineas = [
        linea_txt("100001", "A1", "32188"),
        linea_txt("100002", "A2", "42523.28", "F2"),          # 36,658.00 + 16% IVA
        linea_txt("100003", "Q1", "500"),                     # pago quebrado: 500 + 200 + 300
        linea_txt("100003", "Q1", "200"),
        linea_txt("100003", "Q1", "300"),
        linea_txt("100004", "M1", "1500"),                    # una línea, dos facturas
        linea_txt("100005", "S1", "40180"),                   # sin XML
        linea_txt("100006", "D1", "30156"),                   # XML por 30,000: diferencia
        linea_txt("100007", "B1", "4000"),                    # Activo=0 en catálogo (se ignora)
        linea_txt("100008", "F1", "1000"),                    # factura de agosto
        linea_txt("100009", "E1", "2000"),                    # efectivo
        linea_txt("100010", "P1", "3000"),                    # PPD + 99
        linea_txt("100011", "R1", "4500"),                    # otro receptor
        linea_txt("100012", "N1", "5500"),                    # PUE + 99
        linea_txt("100013", "H1", "8000"),                    # no RESICO (612)
        linea_txt("999999", "X1", "100"),                     # ID que no existe
        linea_txt("PV0001", "V1", "500"),                     # proveedor sin RFC
    ]
    (mes / "Entrada" / "Honorarios.txt").write_text("\n".join(lineas) + "\n", encoding="latin-1")
    return raiz


if __name__ == "__main__":
    destino = crear(sys.argv[1] if len(sys.argv) > 1 else "Honorarios_prueba")
    print(f"Datos de prueba creados en: {destino.resolve()}")
