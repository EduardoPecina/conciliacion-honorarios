"""Pruebas del servicio (conciliación sin preguntas) y del portal web. Datos inventados, sin internet.

    python -m unittest discover -s tests -v
"""

import io
import os
import sys
import tempfile
import unittest
import zipfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import test_conciliacion  # noqa: E402,F401  (define HONORARIOS_RAIZ antes de importar el programa)
import datos_prueba  # noqa: E402
from honorarios import historial, servicio, utilidades  # noqa: E402
from honorarios.config import ARCHIVO_HISTORIAL, CONFIG, RAIZ  # noqa: E402
from honorarios.entradas import obtener_rutas, rutas_en  # noqa: E402
from portal.archivos import preparar_carpeta  # noqa: E402

ANIO, MES = datos_prueba.ANIO, datos_prueba.MES
HOY = datetime(ANIO, MES, 30)


def sap_simulado(ids, carpeta_salida, anio, mes, reusar=True, progreso=None, avisar=None):
    return datos_prueba.sap_simulado(), "consultado (simulado)"


def sat_todo_vigente(facturas, progreso=None, avisar=None):
    estados = {f["UUID"]: dict(test_conciliacion.VIGENTE) for f in facturas if f["UUID"]}
    return estados, f"{len(estados)} UUID consultados (simulado)"


class TestServicio(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        datos_prueba.crear(RAIZ)

    def test_mismo_resultado_que_el_programa_de_consola_y_no_registra_solo(self):
        ruta_historial = CONFIG / ARCHIVO_HISTORIAL
        ruta_historial.unlink(missing_ok=True)
        avisos = []
        conciliacion = servicio.conciliar_periodo(obtener_rutas(ANIO, MES), ANIO, MES, hoy=HOY, avisar=avisos.append,
                                                  progreso_sat=lambda h, t: None, consultar_sat=sat_todo_vigente,
                                                  consultar_sap=sap_simulado)
        estatus = {}
        for fila in conciliacion.resultado:
            estatus.setdefault(fila["IDProveedor"], set()).add(fila["Estatus"])
        for id_prov, esperado in datos_prueba.CASOS_ESPERADOS.items():
            with self.subTest(proveedor=id_prov):
                self.assertEqual(estatus.get(id_prov), {esperado})
        self.assertTrue(Path(conciliacion.destino).exists())
        self.assertTrue(any("Cruzando" in a for a in avisos))
        self.assertFalse(ruta_historial.exists(), "conciliar no debe registrar el historial por sí solo")

        self.assertTrue(conciliacion.se_puede_registrar)
        self.assertIn("UUID registrados", servicio.registrar_historial(conciliacion))
        registrados, _ = historial.cargar_historial()
        self.assertTrue(registrados)
        ruta_historial.unlink()


    def test_sin_sap_no_se_registra_el_historial(self):
        ruta_historial = CONFIG / ARCHIVO_HISTORIAL
        ruta_historial.unlink(missing_ok=True)
        sin_sap = lambda *a, **k: (None, "No se pudo conectar a SAP")
        conciliacion = servicio.conciliar_periodo(obtener_rutas(ANIO, MES), ANIO, MES, hoy=HOY, avisar=lambda t: None,
                                                  progreso_sat=lambda h, t: None, consultar_sat=sat_todo_vigente,
                                                  consultar_sap=sin_sap)
        self.assertEqual(conciliacion.sin_consultar, ["SAP"])
        self.assertFalse(conciliacion.se_puede_registrar)
        self.assertIn("no se pudo consultar SAP", servicio.registrar_historial(conciliacion))
        self.assertFalse(ruta_historial.exists())


class TestArchivosSubidos(unittest.TestCase):
    def zip_con(self, archivos):
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as z:
            for nombre, contenido in archivos.items():
                z.writestr(nombre, contenido)
        return buffer.getvalue()

    def test_xml_sueltos_zip_y_nombres_repetidos(self):
        with tempfile.TemporaryDirectory() as d:
            zip_ = self.zip_con({"sep/a.xml": "<a/>", "sep/leeme.pdf": "no es xml", "b.XML": "<b/>"})
            n = preparar_carpeta(d, ("pagos.txt", b"linea"),
                                 [("a.xml", b"<a/>"), ("a.xml", b"<a2/>"), ("facturas.zip", zip_)])
            self.assertEqual(n, 4)
            xml = Path(d) / "Entrada" / "XML"
            self.assertTrue((xml / "a.xml").exists() and (xml / "a (2).xml").exists())
            self.assertTrue((xml / "facturas" / "sep" / "a.xml").exists())
            self.assertFalse(list(xml.rglob("*.pdf")))
            rutas = rutas_en(Path(d))                       # el motor acepta la carpeta tal cual
            self.assertEqual(len(rutas["xmls"]), 4)
            self.assertEqual(rutas["txt"].name, "pagos.txt")

    def test_zip_no_escribe_fuera_de_la_carpeta(self):
        with tempfile.TemporaryDirectory() as d:
            zip_ = self.zip_con({"../../fuera.xml": "<x/>", "/absoluto.xml": "<y/>"})
            preparar_carpeta(d, ("pagos.txt", b"linea"), [("malo.zip", zip_)])
            creados = [p for p in Path(d).parent.rglob("fuera.xml")]
            self.assertTrue(all(Path(d) in p.parents for p in creados))
            self.assertTrue((Path(d) / "Entrada" / "XML" / "malo" / "fuera.xml").exists())

    def test_sin_xml_avisa(self):
        with tempfile.TemporaryDirectory() as d, self.assertRaises(utilidades.ErrorUsuario):
            preparar_carpeta(d, ("pagos.txt", b"linea"), [("notas.pdf", b"x")])


class TestLanzador(unittest.TestCase):
    """portal/iniciar.py (EJECUTAR_PORTAL.bat): abre el navegador sin ventana negra."""

    def test_portal_ya_abierto_solo_abre_otra_pestana(self):
        from unittest import mock
        from portal import iniciar
        with mock.patch.object(iniciar, "portal_abierto", return_value=True), \
                mock.patch("webbrowser.open") as navegador, mock.patch("subprocess.Popen") as arrancar:
            self.assertEqual(iniciar.iniciar(), 0)
        navegador.assert_called_once_with(iniciar.URL)
        arrancar.assert_not_called()                          # no se arranca un segundo portal

    def test_si_el_portal_falla_avisa_con_ventana(self):
        from unittest import mock
        from portal import iniciar
        fallido = mock.Mock(**{"poll.return_value": 1})       # el proceso terminó antes de estar listo
        with tempfile.TemporaryDirectory() as d, \
                mock.patch.object(iniciar, "BITACORA", Path(d) / "portal.log"), \
                mock.patch.object(iniciar, "portal_abierto", return_value=False), \
                mock.patch("subprocess.Popen", return_value=fallido), \
                mock.patch("webbrowser.open") as navegador, mock.patch.object(iniciar, "avisar") as avisar:
            self.assertEqual(iniciar.iniciar(), 1)
        navegador.assert_not_called()
        self.assertIn("No se pudo iniciar", avisar.call_args.args[0])


class TestCerrarPortal(unittest.TestCase):
    """'Cerrar el portal' detiene solo los procesos de ESTE portal, nunca otro Python (p. ej. el de VS Code)."""

    def test_reconoce_solo_los_procesos_de_este_portal(self):
        from portal import apagado
        repo = apagado.PORTAL.parent
        casos = {
            ("python", "-m", "streamlit", "run", "portal/app.py"): (repo, True),         # el portal
            ("pythonw.exe", str(apagado.PORTAL / "iniciar.py")): ("/", True),            # el lanzador
            ("python", "-c", "import time", "portal/app.py"): (repo, True),              # portal viejo colgado
            ("python", "otro_programa.py"): (repo, False),                               # otro Python
            ("python", "portal/app.py"): ("/otra/instalacion", False),                   # otro portal ajeno
            ("node", "portal/app.py"): (repo, False),                                    # no es Python
            (): (repo, False),
        }
        for linea, (carpeta, esperado) in casos.items():
            with self.subTest(linea=linea):
                self.assertEqual(apagado.es_del_portal(list(linea), carpeta), esperado)


try:
    from streamlit.testing.v1 import AppTest
except ImportError:                                       # el portal es opcional: sin Streamlit no se prueba
    AppTest = None


@unittest.skipIf(AppTest is None, "Streamlit no está instalado (pip install -r requirements-portal.txt)")
class TestPaginasDelPortal(unittest.TestCase):
    PAGINAS = Path(__file__).resolve().parent.parent / "portal" / "paginas"

    @classmethod
    def setUpClass(cls):
        datos_prueba.crear(RAIZ)
        os.chdir(Path(__file__).resolve().parent.parent)

    def abrir(self, pagina):
        app = AppTest.from_file(str(self.PAGINAS / pagina), default_timeout=30).run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        return app

    def test_conciliar_sin_archivos_no_deja_procesar(self):
        app = self.abrir("conciliar.py")
        self.assertTrue(app.button[0].disabled)
        self.assertEqual(app.title[0].value, "Conciliación de Honorarios")

    def test_no_se_ofrecen_meses_del_futuro(self):
        from honorarios.utilidades import fecha_actual
        hoy, _ = fecha_actual()
        app = self.abrir("conciliar.py")
        anio, mes = app.selectbox[0], app.selectbox[1]
        self.assertEqual(anio.options, [str(hoy.year - 1), str(hoy.year)])
        self.assertEqual(len(mes.options), hoy.month)              # año actual: de enero al mes actual
        anio.set_value(hoy.year - 1).run()
        self.assertEqual(len(app.selectbox[1].options), 12)       # año anterior: los 12 meses

    def test_catalogo(self):
        app = self.abrir("catalogo.py")
        self.assertEqual(app.metric[0].value, str(len(datos_prueba.PROVEEDORES)))

    def test_historial_vacio(self):
        (CONFIG / ARCHIVO_HISTORIAL).unlink(missing_ok=True)
        app = self.abrir("historial.py")
        self.assertIn("Todavía no hay historial", app.info[0].value)


if __name__ == "__main__":
    unittest.main(verbosity=2)
