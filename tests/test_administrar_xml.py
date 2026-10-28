"""Pruebas de Administrar carpeta de XML: quitar de un ZIP los XML que no deben estar y descargar uno nuevo.
Datos inventados.

    python -m unittest discover -s tests -v
"""

import sys
import tempfile
import unittest
import zipfile
from io import BytesIO
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import test_conciliacion  # noqa: E402,F401  (define HONORARIOS_RAIZ antes de importar el programa)
import datos_prueba  # noqa: E402
from honorarios.entradas import leer_xmls  # noqa: E402
from honorarios.utilidades import ErrorUsuario  # noqa: E402
from portal.archivos import buscar_en_zip, leer_zip, preparar_xmls, zip_sin  # noqa: E402

U = {n: f"{n:08X}-0000-4000-8000-000000000000" for n in range(1, 5)}


def cfdi(uuid, rfc="MERU800101AB1", subtotal="1000"):
    return datos_prueba.xml_cfdi(rfc, subtotal, uuid=uuid).encode("utf-8")


def zip_de_prueba():
    contenido = {
        "Septiembre/A1872.xml": cfdi(U[1]),
        "Septiembre/A1872.pdf": b"%PDF factura A1872",
        "Septiembre/B2.xml": cfdi(U[2]),
        "Octubre/B2.xml": cfdi(U[3]),            # mismo nombre que otro, en otra carpeta
        "C3.xml": cfdi(U[4], subtotal="2500"),
        "notas.txt": b"no se toca",
    }
    salida = BytesIO()
    with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, datos in contenido.items():
            z.writestr(zipfile.ZipInfo(nombre, date_time=(2026, 10, 1, 9, 30, 0)), datos)
    return salida.getvalue(), contenido


class TestAdministrarXml(unittest.TestCase):
    def setUp(self):
        self.datos, self.contenido = zip_de_prueba()
        self.archivo, self.xmls = leer_zip("Lote Q18.zip", self.datos)

    def test_lee_cada_xml_con_su_ruta_como_en_el_excel(self):
        rutas = {x["ArchivoXML"]: x for x in self.xmls}
        self.assertEqual(set(rutas), {"Lote Q18/Septiembre/A1872.xml", "Lote Q18/Septiembre/B2.xml",
                                      "Lote Q18/Octubre/B2.xml", "Lote Q18/C3.xml"})
        self.assertEqual(rutas["Lote Q18/C3.xml"]["UUID"], U[4])
        self.assertEqual(rutas["Lote Q18/C3.xml"]["RFCEmisor"], "MERU800101AB1")

    def test_la_ruta_es_la_misma_que_escribe_el_conciliador(self):
        with tempfile.TemporaryDirectory() as tmp:
            carpeta = Path(tmp) / "XML"
            preparar_xmls(carpeta, [("Lote Q18.zip", self.datos)])
            facturas, _ = leer_xmls(sorted(carpeta.rglob("*.xml")), carpeta)
        del_conciliador = {f["ArchivoXML"].replace("\\", "/") for f in facturas}
        self.assertEqual(del_conciliador, {x["ArchivoXML"] for x in self.xmls})

    def test_busca_por_ruta_del_excel_nombre_o_uuid(self):
        pegado = ["ArchivoXML",                                    # el encabezado de la columna, si se copia
                  "Lote Q18\\Septiembre\\A1872.xml",               # como sale en el Excel en Windows
                  "C3.xml",                                        # solo el nombre
                  U[3].lower(),                                    # el UUID, en minúsculas
                  "  ",
                  "B2.xml",                                        # hay dos: no se adivina
                  "Z9.xml"]                                        # no existe
        encontrados, problemas = buscar_en_zip(self.xmls, pegado)
        self.assertEqual([x["Ruta"] for x in encontrados], ["Septiembre/A1872.xml", "C3.xml", "Octubre/B2.xml"])
        self.assertEqual([p[0] for p in problemas], ["B2.xml", "Z9.xml"])
        self.assertIn("ruta completa", problemas[0][1])

    def test_el_zip_nuevo_es_igual_sin_lo_que_se_quita(self):
        encontrados, _ = buscar_en_zip(self.xmls, ["Lote Q18/Septiembre/A1872.xml", U[3]])
        nuevo, quitados = zip_sin(self.archivo, encontrados)
        self.assertEqual(quitados, ["Octubre/B2.xml", "Septiembre/A1872.pdf", "Septiembre/A1872.xml"])
        with zipfile.ZipFile(BytesIO(nuevo)) as z:
            self.assertEqual(sorted(z.namelist()), ["C3.xml", "Septiembre/B2.xml", "notas.txt"])
            for nombre in z.namelist():
                self.assertEqual(z.read(nombre), self.contenido[nombre], "el contenido no cambia")
                self.assertEqual(z.getinfo(nombre).date_time, (2026, 10, 1, 9, 30, 0), "ni la fecha")

    def test_se_puede_dejar_el_pdf(self):
        encontrados, _ = buscar_en_zip(self.xmls, ["Lote Q18/Septiembre/A1872.xml"])
        nuevo, quitados = zip_sin(self.archivo, encontrados, con_mismo_nombre=False)
        self.assertEqual(quitados, ["Septiembre/A1872.xml"])
        self.assertIn("Septiembre/A1872.pdf", zipfile.ZipFile(BytesIO(nuevo)).namelist())

    def test_zip_que_no_sirve(self):
        with self.assertRaises(ErrorUsuario):
            leer_zip("malo.zip", b"no es un zip")


if __name__ == "__main__":
    unittest.main(verbosity=2)
