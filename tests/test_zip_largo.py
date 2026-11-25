import io
import tempfile
import unittest
import zipfile
from pathlib import Path

from honorarios.entradas import leer_xmls
from portal.archivos import preparar_xmls


class TestZipLargo(unittest.TestCase):
    def test_facturas_homonimas_con_ruta_larga(self):
        contenido = io.BytesIO()
        carpeta_interna = ("Revision preliminar/" * 12)
        nombres = [carpeta_interna + grupo + "/factura.xml" for grupo in ("A", "B")]
        with zipfile.ZipFile(contenido, "w") as archivo:
            for nombre in nombres:
                archivo.writestr(nombre, '<Comprobante xmlns="http://www.sat.gob.mx/cfd/4" SubTotal="1" Total="1"/>')
        with tempfile.TemporaryDirectory() as tmp:
            carpeta = Path(tmp)
            preparar_xmls(carpeta, [("Noviembre.zip", contenido.getvalue())])
            rutas = list(carpeta.rglob("*.xml"))
            self.assertEqual(len(rutas), 2)
            # Incluso en Windows con rutas largas habilitadas, comprobar la compatibilidad
            # con equipos que conservan el límite clásico.
            self.assertTrue(all(len(str(ruta)) < 260 for ruta in rutas), rutas)
            facturas, errores = leer_xmls(rutas, carpeta)
            self.assertEqual(errores, [])
            self.assertCountEqual([f["ArchivoXML"] for f in facturas],
                                  [str(Path("Noviembre") / n) for n in nombres])
