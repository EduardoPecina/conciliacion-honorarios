"""
Pruebas automáticas con datos inventados (tests/datos_prueba.py). No tocan datos reales ni SAP.

Correr desde la carpeta del repositorio:
    python -m unittest -v
"""

import os
import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

# La carpeta de datos de prueba se define ANTES de importar el programa (config lee la ruta al cargar).
_TMP = tempfile.TemporaryDirectory()
os.environ["HONORARIOS_RAIZ"] = str(Path(_TMP.name) / "Honorarios")

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import datos_prueba  # noqa: E402
from honorarios import cruce, entradas, estado_sat, historial, reglas, reporte, utilidades  # noqa: E402
from honorarios.config import RAIZ  # noqa: E402

ANIO, MES = datos_prueba.ANIO, datos_prueba.MES


def correr_conciliacion():
    """Ejecuta el flujo completo sin preguntas en pantalla y sin SAP."""
    rutas = entradas.obtener_rutas(ANIO, MES)
    catalogo = entradas.leer_catalogo(rutas["catalogo"])
    lineas = entradas.leer_txt(rutas["txt"])
    facturas, errores = entradas.leer_xmls(rutas["xmls"], rutas["carpeta_xml"])
    hist, _ = historial.cargar_historial()
    resultado, sobrantes = cruce.conciliar(lineas, catalogo, facturas, ANIO, MES, None, hist)
    return rutas, resultado, sobrantes, errores


class ConDatosDePrueba(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        datos_prueba.crear(RAIZ)
        cls.rutas, cls.resultado, cls.sobrantes, cls.errores = correr_conciliacion()
        cls.estatus = {}
        for fila in cls.resultado:
            cls.estatus.setdefault(fila["IDProveedor"], set()).add(fila["Estatus"])


class TestCruce(ConDatosDePrueba):
    def test_estatus_esperado_por_caso(self):
        for id_prov, esperado in datos_prueba.CASOS_ESPERADOS.items():
            with self.subTest(proveedor=id_prov):
                self.assertEqual(self.estatus.get(id_prov), {esperado})

    def test_todas_las_lineas_tienen_estatus(self):
        self.assertEqual(len(self.resultado), 17)
        self.assertTrue(all(f.get("Estatus") for f in self.resultado))

    def test_pago_quebrado_no_duplica_importes_del_xml(self):
        quebrado = [f for f in self.resultado if f["IDProveedor"] == "100003"]
        totales = [f.get("Total") for f in quebrado if f.get("Total") not in (None, "")]
        self.assertEqual(len(totales), 1, "los importes del XML van solo en la primera línea del grupo")

    def test_xml_sobrantes_y_errores(self):
        self.assertIn("20_sin_linea_txt.xml", {s["ArchivoXML"] for s in self.sobrantes})
        self.assertEqual([e["Archivo"] for e in self.errores], ["30_danado.xml"])

    def test_xml_en_subcarpeta_se_lee(self):
        usados = " ".join(str(f.get("ArchivoXML", "")) for f in self.resultado)
        self.assertIn("04b_dos.XML", usados)

    def test_ppd_lleva_nota_informativa(self):
        fila = next(f for f in self.resultado if f["IDProveedor"] == "100010")
        self.assertIn("comprobante de pago", fila.get("Observaciones", ""))


class TestReporte(ConDatosDePrueba):
    def test_genera_excel_con_hojas(self):
        from openpyxl import load_workbook
        estados = {"Validación SAP": "No se consultó"}
        destino, resumen = reporte.generar_reporte(self.rutas, ANIO, MES, self.resultado, self.sobrantes,
                                                   self.errores, len(self.rutas["xmls"]), estados)
        self.assertTrue(destino.exists())
        hojas = load_workbook(destino).sheetnames
        for hoja in ("Resumen", "Detalle", "XML sin TXT", "Errores XML", "Info"):
            self.assertIn(hoja, hojas)
        total = resumen.loc[resumen["Estatus"] == "TOTAL TXT", "Lineas"].iloc[0]
        self.assertEqual(int(total), 17)

    def test_texto_con_igual_no_se_vuelve_formula(self):
        """Un nombre de emisor o folio que empieza con '=' debe quedar como texto, no como fórmula."""
        import copy
        from openpyxl import load_workbook
        malicioso = '=HYPERLINK("http://ejemplo.invalid","clic")'
        resultado = copy.deepcopy(self.resultado)
        for fila in resultado:
            fila["Nombre"], fila["Folio"] = malicioso, malicioso
        sobrantes = copy.deepcopy(self.sobrantes)
        for s in sobrantes:
            s["NombreEmisor"] = malicioso
        destino, _ = reporte.generar_reporte(self.rutas, ANIO, MES, resultado, sobrantes, self.errores,
                                             len(self.rutas["xmls"]), {})
        libro = load_workbook(destino)
        encontrados = 0
        for hoja in ("Detalle", "XML sin TXT", "Por proveedor"):
            for fila in libro[hoja].iter_rows():
                for celda in fila:
                    if celda.value == malicioso:
                        encontrados += 1
                        self.assertEqual(celda.data_type, "s", f"{hoja}!{celda.coordinate}")
        self.assertGreater(encontrados, 0)
        # Las fórmulas propias del programa siguen siendo fórmulas
        self.assertEqual(libro["Resumen"]["B2"].data_type, "f")


def respuesta_sat(estado, codigo="S - Comprobante obtenido satisfactoriamente.", cancelacion=""):
    """Respuesta SOAP con la forma de la del servicio ConsultaCFDIService del SAT."""
    return ('<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body>'
            '<ConsultaResponse xmlns="http://tempuri.org/"><ConsultaResult '
            'xmlns:a="http://schemas.datacontract.org/2004/07/Sat.Cfdi.Negocio.ConsultaCfdi.Servicio" '
            'xmlns:i="http://www.w3.org/2001/XMLSchema-instance">'
            f'<a:CodigoEstatus>{codigo}</a:CodigoEstatus><a:EsCancelable>Cancelable con aceptación</a:EsCancelable>'
            f'<a:Estado>{estado}</a:Estado><a:EstatusCancelacion>{cancelacion}</a:EstatusCancelacion>'
            '<a:ValidacionEFOS>200</a:ValidacionEFOS></ConsultaResult></ConsultaResponse></s:Body></s:Envelope>')


class TestEstadoSat(unittest.TestCase):
    """Consulta del estatus del CFDI en el SAT, con respuestas simuladas (sin internet)."""

    FACTURA = {"UUID": "CEE4BE01-ADFA-4DEB-8421-ADD60F0BEDAC", "RFCEmisor": "MERU800101AB1",
               "RFCReceptor": "SIN070530EI5", "TotalTexto": "1000.00", "SelloFinal": "/OAgdg=="}

    def test_expresion_impresa(self):
        self.assertEqual(estado_sat.total_expresion("1000.00"), "1000.0")
        self.assertEqual(estado_sat.total_expresion("2010.010000"), "2010.01")
        self.assertEqual(estado_sat.expresion(self.FACTURA),
                         "https://verificacfdi.facturaelectronica.sat.gob.mx/default.aspx?"
                         "id=CEE4BE01-ADFA-4DEB-8421-ADD60F0BEDAC&re=MERU800101AB1&rr=SIN070530EI5"
                         "&tt=1000.0&fe=/OAgdg==")

    def test_lee_respuesta(self):
        datos = estado_sat.leer_respuesta(respuesta_sat("Vigente", cancelacion="En proceso"))
        self.assertEqual(datos["Estado"], "Vigente")
        self.assertEqual(datos["EstatusCancelacion"], "En proceso")

    def test_consulta_cada_uuid_una_vez(self):
        respuestas = {"U-1": respuesta_sat("Vigente"), "U-2": respuesta_sat("Cancelado"),
                      "U-3": respuesta_sat("No Encontrado", "N - 602: Comprobante no encontrado."),
                      "U-4": respuesta_sat("", "N - 601: La expresión impresa proporcionada no es válida.")}
        consultados = []

        def enviar(expresion):
            uuid = expresion.split("id=")[1].split("&")[0]
            consultados.append(uuid)
            return respuestas[uuid]

        facturas = [dict(self.FACTURA, UUID=u) for u in ("U-1", "U-2", "U-3", "U-4", "U-1")]
        resultado, _ = estado_sat.consultar_estados(facturas, enviar)
        self.assertEqual(sorted(consultados), ["U-1", "U-2", "U-3", "U-4"])
        self.assertEqual({u: c["Estado"] for u, c in resultado.items()},
                         {"U-1": "Vigente", "U-2": "Cancelado", "U-3": "No encontrado", "U-4": ""})
        self.assertIn("601", resultado["U-4"]["Error"])

    def test_sin_internet_no_espera_por_cada_factura(self):
        import urllib.error
        llamadas = []

        def enviar(expresion):
            llamadas.append(expresion)
            raise urllib.error.URLError("sin conexión")

        facturas = [dict(self.FACTURA, UUID=f"U-{i}") for i in range(50)]
        pausa, estado_sat.PAUSA_REINTENTO = estado_sat.PAUSA_REINTENTO, 0
        try:
            resultado, texto = estado_sat.consultar_estados(facturas, enviar)
        finally:
            estado_sat.PAUSA_REINTENTO = pausa
        self.assertEqual(len(llamadas), estado_sat.REINTENTOS)
        self.assertEqual(len(resultado), 50)
        self.assertTrue(all(not c["Estado"] for c in resultado.values()))
        self.assertIn("No se pudo conectar", texto)


class TestCruceConEstadoSat(ConDatosDePrueba):
    """Efecto del estatus del SAT en el cruce, con los datos de prueba."""

    VIGENTE = {"Estado": "Vigente", "EstatusCancelacion": "", "EsCancelable": "", "Error": "",
               "SinConexion": False}

    def conciliar(self, estado_de=lambda f: {}, facturas_extra=()):
        """estado_de(factura) -> campos de la consulta al SAT que cambian respecto de 'Vigente'."""
        lineas = entradas.leer_txt(self.rutas["txt"])
        catalogo = entradas.leer_catalogo(self.rutas["catalogo"])
        facturas, _ = entradas.leer_xmls(self.rutas["xmls"], self.rutas["carpeta_xml"])
        facturas += list(facturas_extra)
        estados = {f["UUID"]: dict(self.VIGENTE, **estado_de(f)) for f in facturas}
        resultado, sobrantes = cruce.conciliar(lineas, catalogo, facturas, ANIO, MES, None, {}, estados)
        por_id = {}
        for fila in resultado:
            por_id.setdefault(fila["IDProveedor"], []).append(fila)
        return por_id, sobrantes

    def test_estatus_segun_el_sat(self):
        por_rfc = {
            "MERU800101AB1": {"Estado": "Cancelado"},                                  # 100001
            "NUID800202AB2": {"Estado": "No encontrado"},                              # 100002
            "MENR800113AC4": {"EstatusCancelacion": "En proceso"},                     # 100013
            "MEPN801212AC3": {"Estado": "", "Error": "sin respuesta del SAT"},         # 100012
        }
        por_id, _ = self.conciliar(lambda f: por_rfc.get(f["RFCEmisor"], {}))
        estatus = {i: {f["Estatus"] for f in filas} for i, filas in por_id.items()}
        self.assertEqual(estatus["100001"], {"Rechazar"})
        self.assertIn("CANCELADA", por_id["100001"][0]["MotivoRechazo"])
        self.assertEqual(estatus["100002"], {"Detenida"})
        self.assertEqual(estatus["100013"], {"Detenida"})
        self.assertIn("pidió cancelar", por_id["100013"][0]["MotivoRechazo"])
        self.assertEqual(estatus["100012"], {"Cruzado con observaciones"})
        self.assertIn("estatus del CFDI en el SAT", por_id["100012"][0]["Observaciones"])
        self.assertEqual(estatus["100007"], {"Cruzado"})
        self.assertEqual(por_id["100007"][0]["EstadoSAT"], "Vigente")
        self.assertEqual(por_id["100013"][0]["EstadoSAT"], "Vigente (cancelación en proceso)")

    def test_prefiere_la_factura_que_sustituye_a_la_cancelada(self):
        """El médico canceló y reemitió por el mismo importe: se cruza la nueva y la cancelada sobra."""
        facturas, _ = entradas.leer_xmls(self.rutas["xmls"], self.rutas["carpeta_xml"])
        cancelada = next(f for f in facturas if f["RFCEmisor"] == "MERU800101AB1")
        nueva = dict(cancelada, UUID="99999999-AAAA-4BBB-8CCC-999999999999", ArchivoXML="01_resico_nueva.xml")
        por_id, sobrantes = self.conciliar(
            lambda f: {"Estado": "Cancelado"} if f["UUID"] == cancelada["UUID"] else {}, [nueva])
        fila = por_id["100001"][0]
        self.assertEqual(fila["UUID"], nueva["UUID"])
        self.assertTrue(fila["Estatus"].startswith("Cruzado"))
        sobra = next(s for s in sobrantes if s["UUID"] == cancelada["UUID"])
        self.assertIn("CANCELADA", sobra["MotivoRechazo"])
        self.assertEqual(sobra["EstadoSAT"], "Cancelado")

class TestReglas(unittest.TestCase):
    """Reglas de negocio sobre un XML individual."""

    CTX = {"uuids": {}, "anio": ANIO, "mes": MES, "periodo": f"{ANIO}-{MES:02d}", "historial": {}}

    def validar(self, **kw):
        with tempfile.TemporaryDirectory() as d:
            ruta = Path(d) / "f.xml"
            ruta.write_text(datos_prueba.xml_cfdi("MERU800101AB1", "1000.00", uuid="U-1", **kw), encoding="utf-8")
            rechazos, obs, detener, _ = reglas.validar_xml(entradas.leer_cfdi(ruta, Path(d)), self.CTX)
            return rechazos, obs, detener

    def test_factura_correcta_sin_rechazos(self):
        self.assertEqual(self.validar()[0], [])

    def test_formas_y_metodos_de_pago(self):
        casos = {("PUE", "03"): True, ("PUE", "99"): True, ("PPD", "99"): True,
                 ("PUE", "01"): False, ("PUE", "04"): False}
        for (metodo, forma), valido in casos.items():
            with self.subTest(metodo=metodo, forma=forma):
                self.assertEqual(self.validar(metodo=metodo, forma=forma)[0] == [], valido)

    def test_receptor_distinto_se_rechaza(self):
        self.assertTrue(any("Facturado a" in r for r in self.validar(receptor="XAXX010101000")[0]))

    def test_factura_de_otro_mes_se_rechaza(self):
        rechazos = self.validar(fecha=f"{ANIO}-{MES - 1:02d}-28T10:00:00")[0]
        self.assertTrue(any("refacturación" in r for r in rechazos))


class TestLecturaTxt(unittest.TestCase):
    def test_encabezado_despues_de_lineas_vacias(self):
        with tempfile.TemporaryDirectory() as d:
            ruta = Path(d) / "pagos.txt"
            datos = ["01.09.2026", "100001", "F1", "1000.00", "X", "E0", "1", "CC", "CTA", "PM", "REF"]
            ruta.write_text("\n\n" + "\t".join(entradas.COLUMNAS_TXT) + "\n" + "\t".join(datos) + "\n",
                            encoding="utf-8")
            lineas = entradas.leer_txt(ruta)
        self.assertEqual(len(lineas), 1)
        self.assertEqual(lineas[0]["IDProveedor"], "100001")

    def test_primera_linea_sin_encabezado_no_se_descarta(self):
        """Sin encabezado, una primera línea con fecha distinta de DD.MM.AAAA antes se tomaba por
        encabezado y su importe desaparecía del reporte sin avisar."""
        with tempfile.TemporaryDirectory() as d:
            ruta = Path(d) / "pagos.txt"
            datos = ["1.09.2026", "100001", "F1", "1000.00", "X", "E0", "1", "CC", "CTA", "PM", "REF"]
            ruta.write_text("\t".join(datos) + "\n", encoding="utf-8")
            lineas = entradas.leer_txt(ruta)
            self.assertEqual([l["Importe"] for l in lineas], [Decimal("1000.00")])

            datos[0] = "01-09-2026"
            ruta.write_text("\t".join(datos) + "\n", encoding="utf-8")
            with self.assertRaises(utilidades.ErrorUsuario) as error:
                entradas.leer_txt(ruta)
        self.assertIn("01-09-2026", str(error.exception))


class TestUtilidades(unittest.TestCase):
    def test_importe_txt(self):
        self.assertEqual(utilidades.importe_txt("1,160.00"), Decimal("1160.00"))
        self.assertEqual(utilidades.importe_txt("1160.00-"), Decimal("-1160.00"))
        self.assertIsNone(utilidades.importe_txt("1.160,00"))      # formato europeo: ambiguo
        self.assertIsNone(utilidades.importe_txt("10.123"))        # tres decimales

    def test_normalizar_id(self):
        self.assertEqual(utilidades.normalizar_id("100002.0"), "100002")
        self.assertEqual(utilidades.normalizar_id("0000100002"), "100002")
        self.assertEqual(utilidades.normalizar_id("PV0001"), "PV0001")

    def test_clasificar_rfc(self):
        self.assertEqual(utilidades.clasificar_rfc("LEHV730331NP6"), "MX")
        self.assertEqual(utilidades.clasificar_rfc("HYL850731H2A"), "MX")
        self.assertEqual(utilidades.clasificar_rfc("880443249"), "Extranjero")
        self.assertEqual(utilidades.clasificar_rfc(""), "Vacío")


if __name__ == "__main__":
    unittest.main(verbosity=2)
