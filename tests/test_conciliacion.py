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
from honorarios import cruce, entradas, historial, reglas, reporte, utilidades  # noqa: E402
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
