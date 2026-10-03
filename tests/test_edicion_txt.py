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


if __name__ == "__main__":
    unittest.main(verbosity=2)
