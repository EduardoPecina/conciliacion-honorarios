"""Pruebas de Edición TXT: lectura/escritura fiel del TXT de SAP, problemas por celda, folios únicos y
revalidación del cruce al editar. Datos inventados, sin internet.

    python -m unittest discover -s tests -v
"""

import sys
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import test_conciliacion  # noqa: E402,F401  (define HONORARIOS_RAIZ antes de importar el programa)
import datos_prueba  # noqa: E402
from test_portal import sap_simulado, sat_todo_vigente  # noqa: E402
from honorarios import folios, servicio, txt_sap  # noqa: E402
from honorarios.config import ARCHIVO_FOLIOS, CONFIG, RAIZ  # noqa: E402
from honorarios.entradas import obtener_rutas  # noqa: E402

ANIO, MES = datos_prueba.ANIO, datos_prueba.MES
L1 = datos_prueba.linea_txt("100001", "A1", "32,188.00")
L2 = datos_prueba.linea_txt("100002", "A2", "42523.28", "F2")


class TestFormatoDelTxt(unittest.TestCase):
    def test_sin_cambios_sale_identico_byte_por_byte(self):
        for nombre, datos in {
            "CRLF latin-1 con acento": (L1 + "\r\n" + L2.replace("Q18", "Ñ18") + "\r\n").encode("latin-1"),
            "LF utf-8 sin salto final": (L1 + "\n" + L2).encode("utf-8"),
            "utf-8 con BOM": b"\xef\xbb\xbf" + (L1 + "\r\n" + L2 + "\r\n").encode("utf-8"),
            "espacios en los valores": (L1.replace("A1", " A1 ") + "\n").encode("utf-8"),
        }.items():
            with self.subTest(nombre):
                self.assertEqual(txt_sap.escribir(txt_sap.leer(datos)), datos)

    def test_encabezado_y_renglones_vacios_no_van_a_sap(self):
        encabezado = "\t".join(txt_sap.COLUMNAS)
        txt = txt_sap.leer(f"{encabezado}\r\n{L1}\r\n\r\n{L2}\r\n".encode("utf-8"))
        self.assertEqual(sorted(txt.lineas), [2, 4])            # números de línea del archivo (como el motor)
        self.assertEqual(txt_sap.escribir(txt), f"{L1}\r\n{L2}\r\n".encode("utf-8"))

    def test_una_edicion_solo_cambia_esa_celda(self):
        datos = f"{L1}\r\n{L2}\r\n".encode("latin-1")
        txt = txt_sap.leer(datos)
        lineas = {n: list(v) for n, v in txt.lineas.items()}
        lineas[2][3] = "42500.00"
        esperado = f"{L1}\r\n{L2.replace('42523.28', '42500.00')}\r\n".encode("latin-1")
        self.assertEqual(txt_sap.escribir(txt, lineas), esperado)
        self.assertEqual(len(txt_sap.escribir(txt, lineas).split(b"\r\n")[0].split(b"\t")), 11)

    def test_problemas_por_celda(self):
        malas = [L1, L1.replace("32,188.00", "32.188,00"), L2.replace("\tF2\t", "\tF3\t"),
                 datos_prueba.linea_txt("888888", "Z1", "10"), "10.09.2026\t100001\tsolo tres"]
        txt = txt_sap.leer("\n".join(malas).encode("utf-8"))
        lista = txt_sap.problemas(txt, txt.lineas, {"100001", "100002"}, f"{ANIO}-{MES:02d}")
        por = {(n, c): (texto, bloquea) for n, c, texto, bloquea in lista}
        self.assertIn("repetido para el mismo proveedor", por[(1, "Folio")][0])       # A1 en líneas 1 y 2
        self.assertTrue(por[(2, "Monto")][1], "un importe inválido bloquea la descarga")
        self.assertIn("E0", por[(3, "Tax")][0])
        self.assertIn("no está en el catálogo", por[(4, "IDProveedor")][0])
        self.assertTrue(por[(5, "")][1], "una línea con columnas de menos bloquea la descarga")
        self.assertNotIn((1, "Monto"), por, "1,234.56 es un importe válido")


class TestReferencia(unittest.TestCase):
    """La Referencia empieza con el folio: "Formatear folios" cambia solo esa parte."""

    def test_solo_cambia_la_parte_del_folio(self):
        casos = [
            ("A1893      Q18-2026", "A1893", "7KQ2M9XHRT_Q18-2026"),
            ("7      Q18-2026", "7", "7KQ2M9XHRT_Q18-2026"),
            ("A1805 Q18-2026 extra", "A1805", "7KQ2M9XHRT_Q18-2026_extra"),       # sin espacios: guion bajo
            ("Q18-2026", "A1", "7KQ2M9XHRT_Q18-2026"),                           # no traía folio: se antepone
            ("  A1   Q18-2026  ", "A1", "7KQ2M9XHRT_Q18-2026"),                    # espacios al inicio y al final
            ("", "A1", "7KQ2M9XHRT"),
        ]
        for referencia, folio, esperado in casos:
            with self.subTest(referencia=referencia):
                self.assertEqual(txt_sap.referencia_con_folio(referencia, folio, "7KQ2M9XHRT"), esperado)

    def test_formatear_otra_vez_no_cambia_nada(self):
        una = txt_sap.referencia_con_folio("A1893      Q18-2026", "A1893", "7KQ2M9XHRT")
        self.assertEqual(txt_sap.referencia_con_folio(una, "7KQ2M9XHRT", "7KQ2M9XHRT"), una)
        self.assertNotIn(" ", una)
        self.assertEqual(len(una), len("A1893      Q18-2026"), "con folio de 10 queda del mismo largo que antes")

    def test_la_referencia_ya_no_se_edita_a_mano_y_tax_si(self):
        self.assertNotIn("Referencia", txt_sap.EDITABLES)
        self.assertNotIn("Folio", txt_sap.EDITABLES)
        self.assertIn("Tax", txt_sap.EDITABLES)


class TestFolios(unittest.TestCase):
    def test_misma_factura_mismo_folio_y_formato(self):
        f = folios.folio_de("UUID-1#0")
        self.assertEqual(f, folios.folio_de("UUID-1#0"))
        self.assertNotEqual(f, folios.folio_de("UUID-1#1"))
        self.assertTrue(folios.es_folio_generado(f))
        self.assertFalse(set("01IO") & set(folios.ALFABETO))
        self.assertEqual(len(folios.ALFABETO), 32)

    def test_pago_quebrado_y_lineas_sin_factura(self):
        asignados = folios.asignar_folios([(1, ("U1",)), (2, ("U1",)), (3, ()), (4, ("U3", "U2"))], {})
        todos = [f for f, _ in asignados.values()]
        self.assertEqual(len(set(todos)), 4, "nunca se repiten dentro del archivo")
        self.assertEqual(asignados[4][1], "U2|U3#0", "varias facturas: el orden de los UUID no importa")
        otra_vez = folios.asignar_folios([(1, ("U1",)), (2, ("U1",)), (3, ()), (4, ("U2", "U3"))], {},
                                         actuales={3: asignados[3][0]})
        self.assertEqual(otra_vez, asignados, "formatear otra vez no cambia nada")

    def test_choque_con_un_folio_registrado_se_recalcula(self):
        ocupado = folios.folio_de("U9#0")
        asignados = folios.asignar_folios([(1, ("U9",))], {ocupado: {"Clave": "otra factura"}})
        self.assertNotEqual(asignados[1][0], ocupado)
        mismo = folios.asignar_folios([(1, ("U9",))], {ocupado: {"Clave": "U9#0"}})
        self.assertEqual(mismo[1][0], ocupado, "la misma factura ya registrada conserva su folio")

    def test_cien_mil_folios_sin_repetirse(self):
        generados = {folios.folio_de(f"{i:032X}#0") for i in range(100_000)}
        self.assertEqual(len(generados), 100_000)

    def test_registro(self):
        ruta = CONFIG / ARCHIVO_FOLIOS
        ruta.unlink(missing_ok=True)
        registro = {"Folio": "ABCDEFGHJK", "Clave": "U1#0", "IDProveedor": "1", "Periodo": "2026-09", "Linea": 1}
        folios.registrar_folios([registro])
        folios.registrar_folios([registro])                       # descargar dos veces no duplica
        self.assertEqual(list(folios.cargar_folios()), ["ABCDEFGHJK"])
        ruta.unlink()


class TestRevalidar(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        datos_prueba.crear(RAIZ)
        cls.conciliacion = servicio.conciliar_periodo(
            obtener_rutas(ANIO, MES), ANIO, MES, hoy=datetime(ANIO, MES, 30), avisar=lambda t: None,
            progreso_sat=lambda h, t: None, consultar_sat=sat_todo_vigente, consultar_sap=sap_simulado)
        cls.txt = txt_sap.leer(obtener_rutas(ANIO, MES)["txt"].read_bytes())

    def estatus(self, conciliacion, linea):
        return next(f["Estatus"] for f in conciliacion.resultado if f["Linea"] == linea)

    def test_sin_cambios_da_lo_mismo(self):
        lineas, errores = txt_sap.lineas_para_el_motor(self.txt.lineas)
        self.assertFalse(errores)
        otra = servicio.revalidar(self.conciliacion, lineas)
        self.assertEqual([f["Estatus"] for f in otra.resultado], [f["Estatus"] for f in self.conciliacion.resultado])
        self.assertTrue(otra.editada)

    def test_id_equivocado_se_sugiere_y_al_corregirlo_cruza(self):
        # La línea 7 (ID 100005, sin XML) en realidad pagaba la factura de 30,000.00 del proveedor 100006
        lineas = {n: list(v) for n, v in self.txt.lineas.items()}
        lineas[7][3] = "30000.00"
        editada = servicio.revalidar(self.conciliacion, txt_sap.lineas_para_el_motor(lineas)[0])
        sugerencias = txt_sap.sugerencias_de_id(editada.resultado, editada.sobrantes, {})
        self.assertEqual(sugerencias[7][0], "100006")
        self.assertIn("30,000.00", sugerencias[7][1])
        lineas[7][1] = "100006"
        corregida = servicio.revalidar(editada, txt_sap.lineas_para_el_motor(lineas)[0])
        self.assertTrue(self.estatus(corregida, 7).startswith("Cruzado"), self.estatus(corregida, 7))

    def test_id_nuevo_sin_consultar_en_sap(self):
        lineas = {n: list(v) for n, v in self.txt.lineas.items()}
        sap = dict(self.conciliacion.entradas["sap"])
        sap.pop("100001")
        sin_100001 = servicio.replace(self.conciliacion, entradas={**self.conciliacion.entradas, "sap": sap})
        editada = servicio.revalidar(sin_100001, txt_sap.lineas_para_el_motor(lineas)[0])
        self.assertEqual(servicio.ids_sin_consultar_en_sap(editada), ["100001"])
        self.assertNotEqual(self.estatus(editada, 1), "Cruzado", "sin SAP no queda Cruzado")
        consulta = mock.Mock(return_value=(datos_prueba.sap_simulado(), "consultado"))
        with mock.patch.object(servicio, "consultar_sap", consulta):
            completa, _ = servicio.consultar_sap_faltantes(editada, Path("."))
        self.assertIn("100001", consulta.call_args.args[0])
        self.assertEqual(servicio.ids_sin_consultar_en_sap(completa), [])
        self.assertEqual(self.estatus(completa, 1), "Cruzado")


class TestTrabajoDeVariosDias(unittest.TestCase):
    """Cerrar el portal y continuar otro día: se recupera todo sin volver a consultar SAT ni SAP."""

    @classmethod
    def setUpClass(cls):
        datos_prueba.crear(RAIZ)
        cls.rutas = obtener_rutas(ANIO, MES)
        cls.conciliacion = servicio.conciliar_periodo(
            cls.rutas, ANIO, MES, hoy=datetime(ANIO, MES, 30), avisar=lambda t: None,
            progreso_sat=lambda h, t: None, consultar_sat=sat_todo_vigente, consultar_sap=sap_simulado)

    def setUp(self):
        import tempfile
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.carpeta = Path(self.tmp.name) / str(ANIO) / "09 Sep"

    def test_guardar_y_continuar(self):
        txt = ("Honorarios.txt", self.rutas["txt"].read_bytes())
        editada = servicio.replace(self.conciliacion, cambios=[{"Linea": 1, "Columna": "Monto"}], version=7)
        self.assertTrue(servicio.guardar_trabajo(editada, self.carpeta, txt))
        (ruta, nombre, fecha), = servicio.trabajos_guardados(Path(self.tmp.name))
        self.assertEqual(nombre, "septiembre 2026")
        conciliacion, txt_recuperado, carpeta = servicio.cargar_trabajo(ruta)
        self.assertEqual(txt_recuperado, txt)
        self.assertEqual(carpeta, self.carpeta)
        self.assertEqual([f["Estatus"] for f in conciliacion.resultado], [f["Estatus"] for f in editada.resultado])
        self.assertEqual((conciliacion.version, conciliacion.cambios), (7, [{"Linea": 1, "Columna": "Monto"}]))
        lineas, _ = txt_sap.lineas_para_el_motor(txt_sap.leer(txt[1]).lineas)
        servicio.revalidar(conciliacion, lineas)          # se puede seguir editando sin volver a consultar

    def test_trabajo_danado_o_de_otra_version(self):
        import pickle
        self.carpeta.mkdir(parents=True)
        danado = self.carpeta / "Trabajo_2026-09.pkl"
        danado.write_bytes(b"no es un trabajo")
        with self.assertRaises(servicio.ErrorUsuario):
            servicio.cargar_trabajo(danado)
        campos = {f.name: getattr(self.conciliacion, f.name) for f in servicio.fields(self.conciliacion)}
        campos["campo_de_una_version_futura"] = 1
        del campos["consultado"]                          # campo que una versión vieja no tenía
        danado.write_bytes(pickle.dumps({"programa": "0.0", "guardado": "", "txt": ("a.txt", b""),
                                         "conciliacion": campos}))
        conciliacion, _, _ = servicio.cargar_trabajo(danado)
        self.assertEqual(conciliacion.consultado, "")

    def test_actualizar_sat_detecta_una_cancelacion_nueva(self):
        uuid = next(f["UUID"] for f in self.conciliacion.resultado if f["Linea"] == 1)
        def sat_con_cancelada(facturas, progreso=None, avisar=None):
            estados, texto = sat_todo_vigente(facturas)
            estados[uuid] = {**estados[uuid], "Estado": "Cancelado"}
            return estados, texto
        nueva, _ = servicio.actualizar_sat(self.conciliacion, consultar_sat=sat_con_cancelada)
        self.assertEqual(next(f["Estatus"] for f in nueva.resultado if f["Linea"] == 1), "Rechazar")
        self.assertGreater(nueva.version, self.conciliacion.version)
        sin_internet = lambda facturas, **k: ({f["UUID"]: {"SinConexion": True} for f in facturas if f["UUID"]}, "x")
        igual, _ = servicio.actualizar_sat(self.conciliacion, consultar_sat=sin_internet)
        self.assertIs(igual, self.conciliacion, "sin conexión no se pierde lo que ya se tenía")

    def test_actualizar_sap_consulta_todo_sin_reusar(self):
        consulta = mock.Mock(return_value=(datos_prueba.sap_simulado(), "consultado"))
        nueva, _ = servicio.actualizar_sap(self.conciliacion, self.carpeta, consultar=consulta)
        self.assertFalse(consulta.call_args.args[4], "no se reusan las consultas guardadas")
        self.assertEqual(nueva.estados["Validación SAP"], "consultado")

    def test_dias_desde_la_consulta(self):
        c = servicio.replace(self.conciliacion, consultado="01/10/2026 09:00")
        self.assertEqual(servicio.dias_desde_consulta(c, datetime(2026, 10, 3, 10)), 2)
        self.assertEqual(servicio.dias_desde_consulta(servicio.replace(c, consultado="")), 0)

    def test_excel_regenerado_con_los_cambios(self):
        from openpyxl import load_workbook
        cambios = [{"Linea": 7, "Columna": "IDProveedor", "Anterior": "100005", "Nuevo": "100006",
                    "Motivo": "Sugerencia del portal", "Usuario": "conta", "Fecha": "03/10/2026 10:00"}]
        editada = servicio.replace(self.conciliacion, cambios=cambios, editada=True)
        nueva = servicio.regenerar_reporte(editada, self.carpeta, b"TXT editado")
        libro = load_workbook(nueva.destino)
        self.assertIn("Cambios TXT", libro.sheetnames)
        self.assertEqual(libro["Cambios TXT"]["D2"].value, "100006")
        info = {fila[0].value: fila[1].value for fila in libro["Info"].iter_rows()}
        self.assertEqual(info["Archivo TXT"], "TXT_editado_2026-09.txt")
        self.assertIn("1 (ver hoja Cambios TXT", info["Cambios hechos en Edición TXT"])
        self.assertEqual((self.carpeta / "TXT_editado_2026-09.txt").read_bytes(), b"TXT editado")


if __name__ == "__main__":
    unittest.main(verbosity=2)
