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


class TestCandadosDelPeriodo(unittest.TestCase):
    """Archivos de un mes procesados con otro mes elegido: se detiene antes del SAT y SAP, y confirmar nunca
    borra del historial un mes ya registrado."""

    @classmethod
    def setUpClass(cls):
        datos_prueba.crear(RAIZ)

    def test_archivos_de_otro_mes_se_detienen_antes_del_sat(self):
        consultado = []
        sat = lambda facturas, **k: consultado.append(1) or sat_todo_vigente(facturas)
        siguiente = (ANIO, MES + 1) if MES < 12 else (ANIO + 1, 1)
        with self.assertRaises(utilidades.ErrorUsuario) as error:
            servicio.conciliar_periodo(obtener_rutas(ANIO, MES), *siguiente, hoy=datetime(*siguiente, 28),
                                       avisar=lambda t: None, consultar_sat=sat, consultar_sap=sap_simulado)
        self.assertIn("parecen ser de", str(error.exception))
        self.assertIn("No se consultó al SAT ni a SAP", str(error.exception))
        self.assertFalse(consultado)

    def test_confirmar_no_borra_un_mes_ya_registrado(self):
        ruta = CONFIG / ARCHIVO_HISTORIAL
        ruta.unlink(missing_ok=True)
        cruzadas = [{"Estatus": "Cruzado", "UUID": f"U{i}", "IDProveedor": "1", "Linea": i, "Importe": 1}
                    for i in range(10)]
        self.assertIn("10 UUID registrados", historial.guardar_historial(cruzadas, ANIO, MES))
        rechazadas = [dict(f, Estatus="Rechazar") for f in cruzadas]       # archivos de otro mes: nada cruza
        self.assertIn("No se actualizó el historial", historial.guardar_historial(rechazadas, ANIO, MES))
        registrados, _ = historial.cargar_historial()
        self.assertEqual(len(registrados), 10, "los UUID del mes siguen protegidos")
        self.assertIn("9 UUID registrados", historial.guardar_historial(cruzadas[:9], ANIO, MES))  # corrección normal
        ruta.unlink()


class TestConsultasSapGuardadas(unittest.TestCase):
    def test_lo_reusado_conserva_la_fecha_de_su_consulta(self):
        import pandas as pd
        from unittest import mock
        from honorarios import sap
        with tempfile.TemporaryDirectory() as d:
            cache = Path(d) / "CodigosSAP_2026-09.xlsx"
            pd.DataFrame([{"IDProveedor": "1", "Codigos": "RS", "Error": "", "Consulta": "15/09/2026 10:42"},
                          {"IDProveedor": "2", "Codigos": "", "Error": "falló", "Consulta": "15/09/2026 10:43"}]) \
                .to_excel(cache, index=False)
            self.assertEqual(sap.fecha_consultas_guardadas(cache), "15/09/2026 10:42")
            with mock.patch.object(sap, "conectar_sap"), \
                    mock.patch.object(sap, "codigos_retencion", return_value=(["RH"], "")) as consultar:
                resultado, _ = sap.consultar_sap(["1", "2"], Path(d), 2026, 9, reusar=True,
                                                 progreso=lambda *a: None, avisar=lambda t: None)
            consultar.assert_called_once()                    # solo el que había fallado
            self.assertEqual(resultado, {"1": (["RS"], ""), "2": (["RH"], "")})
            guardado = pd.read_excel(cache, dtype=str).set_index("IDProveedor")["Consulta"]
            self.assertEqual(guardado["1"], "15/09/2026 10:42")    # no se finge que se consultó hoy
            self.assertNotEqual(guardado["2"], "15/09/2026 10:43")
            self.assertIsNone(sap.fecha_consultas_guardadas(Path(d) / "no_existe.xlsx"))


class TestArchivosSubidos(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        datos_prueba.crear(RAIZ)                          # rutas_en() necesita el catálogo

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
    """portal/iniciar.py (EJECUTAR_PORTAL.bat): abre el navegador sin ventana negra, un solo portal aunque se den
    muchos clics, y si el portal quedó trabado lo reinicia."""

    def setUp(self):
        from unittest import mock
        from portal import iniciar
        self.iniciar, self.mock = iniciar, mock
        carpeta = tempfile.TemporaryDirectory()
        self.addCleanup(carpeta.cleanup)
        for nombre, valor in {"CANDADO": Path(carpeta.name) / "portal.lock", "BITACORA": Path(carpeta.name) / "log",
                              "PAUSA_ENTRE_CLICS": 0}.items():
            parche = mock.patch.object(iniciar, nombre, valor)
            parche.start()
            self.addCleanup(parche.stop)

    def lanzar(self, abierto, responde=True, proceso=None, al_avisar=None):
        """Corre iniciar() simulando el puerto, el navegador y el arranque. Devuelve (salida, navegador, Popen, avisar)."""
        mock, iniciar = self.mock, self.iniciar
        abierto = abierto if isinstance(abierto, list) else [abierto]
        with mock.patch.object(iniciar, "portal_abierto", side_effect=abierto + [abierto[-1]] * 50), \
                mock.patch.object(iniciar, "portal_responde", return_value=responde), \
                mock.patch("subprocess.Popen", return_value=proceso or mock.Mock(**{"poll.return_value": None})) \
                as arrancar, mock.patch("webbrowser.open") as navegador, \
                mock.patch.object(iniciar, "avisar", side_effect=al_avisar) as avisar:
            salida = iniciar.iniciar()
        return salida, navegador, arrancar, avisar

    def test_portal_ya_abierto_solo_abre_otra_pestana(self):
        salida, navegador, arrancar, _ = self.lanzar(abierto=True)
        self.assertEqual(salida, 0)
        navegador.assert_called_once_with(self.iniciar.URL)
        arrancar.assert_not_called()                          # no se arranca un segundo portal

    def test_portal_cerrado_lo_arranca_y_abre_una_pestana(self):
        salida, navegador, arrancar, avisar = self.lanzar(abierto=[False, False, True])
        self.assertEqual(salida, 0)
        arrancar.assert_called_once()
        navegador.assert_called_once()
        avisar.assert_not_called()

    def test_clics_de_mas_mientras_otro_arranca_no_hacen_nada(self):
        turno = self.iniciar.tomar_turno()                    # otro clic ya está arrancando el portal
        self.assertIsNotNone(turno)
        try:
            salida, navegador, arrancar, avisar = self.lanzar(abierto=False)
        finally:
            self.iniciar.soltar_turno(turno)
        self.assertEqual(salida, 0)
        arrancar.assert_not_called()                          # ni otro portal...
        navegador.assert_not_called()                         # ...ni otra pestaña...
        avisar.assert_not_called()                            # ...ni ventana de error
        self.assertIsNotNone(self.iniciar.tomar_turno(), "al terminar, el turno queda libre")

    def test_si_otro_portal_gano_el_puerto_no_es_error(self):
        termino = self.mock.Mock(**{"poll.return_value": 1})  # "Port 8501 is not available"...
        salida, navegador, _, avisar = self.lanzar(abierto=[False, False, True], proceso=termino)
        self.assertEqual(salida, 0)                           # ...pero el puerto ya lo tiene un portal: sirve
        navegador.assert_called_once()
        avisar.assert_not_called()

    def test_si_el_portal_falla_avisa_con_ventana_y_suelta_el_turno(self):
        fallido = self.mock.Mock(**{"poll.return_value": 1})  # el proceso terminó antes de estar listo
        turno_con_la_ventana_abierta = []

        def ventana(mensaje, error=False):                    # mientras se ve la ventana de error...
            turno = self.iniciar.tomar_turno()
            turno_con_la_ventana_abierta.append(turno is not None)
            self.iniciar.soltar_turno(turno)

        salida, navegador, _, avisar = self.lanzar(abierto=False, proceso=fallido, al_avisar=ventana)
        self.assertEqual(salida, 1)
        navegador.assert_not_called()
        self.assertIn("No se pudo iniciar", avisar.call_args.args[0])
        self.assertEqual(turno_con_la_ventana_abierta, [True], "...otro clic ya puede reintentar")

    def test_portal_trabado_se_detiene_y_se_arranca_de_nuevo(self):
        with self.mock.patch.object(self.iniciar, "detener_portal_trabado", return_value=True) as detener:
            salida, navegador, arrancar, avisar = self.lanzar(abierto=[True, True], responde=False)
        self.assertEqual(salida, 0)
        detener.assert_called_once()
        arrancar.assert_called_once()
        navegador.assert_called_once()
        avisar.assert_not_called()

    def test_portal_trabado_que_no_se_deja_cerrar_avisa(self):
        with self.mock.patch.object(self.iniciar, "detener_portal_trabado", return_value=False):
            salida, navegador, arrancar, avisar = self.lanzar(abierto=True, responde=False)
        self.assertEqual(salida, 1)
        arrancar.assert_not_called()
        navegador.assert_not_called()
        self.assertIn("no responde", avisar.call_args.args[0])

    def test_al_destrabar_no_se_detiene_a_si_mismo_ni_a_quien_lo_lanzo(self):
        import psutil
        from portal import apagado
        padre = psutil.Process().parent()                     # p. ej. pyw.exe, que también parece "del portal"
        trabado = self.mock.Mock(pid=999999)
        with self.mock.patch.object(apagado, "procesos_del_portal",
                                    return_value=[self.mock.Mock(pid=padre.pid), trabado]), \
                self.mock.patch.object(apagado, "detener") as detener, \
                self.mock.patch.object(self.iniciar, "portal_abierto", return_value=False):
            self.assertTrue(self.iniciar.detener_portal_trabado())
        detener.assert_called_once_with([trabado])

    def test_portal_que_contesta_no_esta_trabado(self):
        import http.server
        import threading
        class Contesta(http.server.BaseHTTPRequestHandler):
            def do_GET(self):
                self.send_response(200)
                self.end_headers()
                self.wfile.write(b"ok")

            def log_message(self, *args):
                pass

        servidor = http.server.HTTPServer(("127.0.0.1", 0), Contesta)
        threading.Thread(target=servidor.serve_forever, daemon=True).start()
        try:
            with self.mock.patch.object(self.iniciar, "PUERTO", servidor.server_address[1]):
                self.assertTrue(self.iniciar.portal_responde(intentos=1))
        finally:
            servidor.shutdown()
            servidor.server_close()
        mudo = socket_que_no_contesta()                       # acepta la conexión pero nunca responde
        try:
            with self.mock.patch.object(self.iniciar, "PUERTO", mudo.getsockname()[1]):
                self.assertFalse(self.iniciar.portal_responde(intentos=1, tiempo_espera=0.5))
        finally:
            mudo.close()

    def test_un_solo_turno_entre_procesos_al_mismo_tiempo(self):
        """Cinco "clics" reales (procesos) a la vez: exactamente uno obtiene el turno."""
        import subprocess
        codigo = ("import sys, time; sys.path.insert(0, sys.argv[1]); from portal import iniciar; "
                  "from pathlib import Path; iniciar.CANDADO = Path(sys.argv[2]); "
                  "print(iniciar.tomar_turno() is not None, flush=True); time.sleep(3)")
        procesos = [subprocess.Popen([sys.executable, "-c", codigo, str(Path(__file__).resolve().parents[1]),
                                      str(self.iniciar.CANDADO)], stdout=subprocess.PIPE, text=True)
                    for _ in range(5)]
        resultados = [p.communicate(timeout=60)[0].strip() for p in procesos]
        self.assertEqual(sorted(resultados), ["False"] * 4 + ["True"])


def socket_que_no_contesta():
    import socket
    mudo = socket.socket()
    mudo.bind(("127.0.0.1", 0))
    mudo.listen()
    return mudo


class TestJavaScriptDelPortal(unittest.TestCase):
    """st.html limpia el HTML con DOMPurify y BORRA en silencio un <script> cuyo texto contenga etiquetas ("<div",
    "</p"...). Si alguien edita estos scripts y mete una etiqueta literal, el aviso dejaría de funcionar sin error."""

    def test_los_scripts_no_traen_etiquetas_literales(self):
        import re
        from portal.estilo import aviso_al_salir, pagina_portal_cerrado
        for nombre, html in (("aviso activo", aviso_al_salir(True)), ("aviso apagado", aviso_al_salir(False)),
                             ("portal cerrado", pagina_portal_cerrado())):
            with self.subTest(nombre):
                self.assertTrue(html.startswith("<script>") and html.endswith("</script>"))
                self.assertEqual(re.findall(r"<[/\w]", html[len("<script>"):-len("</script>")]), [])


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

    def test_mes_que_no_es_el_vigente_pide_confirmar(self):
        app = self.abrir("conciliar.py")
        self.assertFalse(app.checkbox, "mes vigente: no se pregunta nada")
        self.assertIn("deben coincidir", app.info[0].value)
        app.selectbox[0].set_value(int(app.selectbox[0].options[0])).run()     # el año pasado: nunca es vigente
        self.assertIn("no es el mes vigente", app.warning[0].value)
        self.assertTrue(app.checkbox[0].label.startswith("Sí, quiero conciliar"))

    def test_fechas_dinamicas_al_cambiar_de_anio(self):
        from unittest import mock
        import streamlit as st
        for hoy, anios, meses_del_actual in ((datetime(2027, 1, 5), ["2026", "2027"], 1),
                                             (datetime(2031, 12, 31), ["2030", "2031"], 12)):
            with self.subTest(hoy=hoy):
                st.cache_data.clear()
                with mock.patch("honorarios.utilidades.fecha_actual", return_value=(hoy, "prueba")):
                    app = self.abrir("conciliar.py")
                    self.assertEqual(app.selectbox[0].options, anios)
                    self.assertEqual(len(app.selectbox[1].options), meses_del_actual)
                    app.selectbox[0].set_value(hoy.year - 1).run()
                    self.assertEqual(len(app.selectbox[1].options), 12)
        st.cache_data.clear()

    def test_reusar_sap_solo_si_ya_se_proceso_y_desmarcado_por_defecto(self):
        import pandas as pd
        from honorarios.config import MESES, REPORTES
        from honorarios.utilidades import fecha_actual
        hoy, _ = fecha_actual()
        carpeta = REPORTES / str(hoy.year) / MESES[hoy.month - 1]
        cache = carpeta / f"CodigosSAP_{hoy.year}-{hoy.month:02d}.xlsx"
        app = self.abrir("conciliar.py")
        self.assertFalse(app.checkbox, "primera corrida del mes: no se ofrece reusar")
        carpeta.mkdir(parents=True, exist_ok=True)
        pd.DataFrame([{"IDProveedor": "1", "Codigos": "RS", "Error": "", "Consulta": "15/09/2026 10:42"},
                      {"IDProveedor": "2", "Codigos": "RH", "Error": "", "Consulta": "16/09/2026 08:00"}]) \
            .to_excel(cache, index=False)
        try:
            app = self.abrir("conciliar.py")
            casilla = app.checkbox[0]
            self.assertFalse(casilla.value, "desmarcada por defecto: lo seguro es consultar todo de nuevo")
            self.assertEqual(casilla.label, "Reusar lo consultado a SAP el 15/09/2026 a las 10:42 (más rápido)")
            self.assertIn("RS↔RH", casilla.help)
        finally:
            cache.unlink()

    def test_catalogo(self):
        app = self.abrir("catalogo.py")
        self.assertEqual(app.metric[0].value, str(len(datos_prueba.PROVEEDORES)))

    def test_historial_vacio(self):
        (CONFIG / ARCHIVO_HISTORIAL).unlink(missing_ok=True)
        app = self.abrir("historial.py")
        self.assertIn("Todavía no hay historial", app.info[0].value)


if __name__ == "__main__":
    unittest.main(verbosity=2)
