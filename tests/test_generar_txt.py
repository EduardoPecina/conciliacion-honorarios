"""Pruebas de Generar TXT: lectura del Excel aprobado (sin editar nada), folios rastreables de 16 caracteres, TXT con el
formato del original y la hoja "TXT para SAP" del Excel de la conciliación. Datos inventados, sin internet.

    python -m unittest discover -s tests -v
"""

import sys
import unittest
from datetime import datetime
from io import BytesIO
from pathlib import Path
from unittest import mock

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import test_conciliacion  # noqa: E402,F401  (define HONORARIOS_RAIZ antes de importar el programa)
import datos_prueba  # noqa: E402
from test_portal import sap_simulado, sat_todo_vigente  # noqa: E402
from honorarios import folios, servicio, txt_sap  # noqa: E402
from honorarios.config import ARCHIVO_FOLIOS, CONFIG, LONGITUD_FOLIO, RAIZ  # noqa: E402
from honorarios.entradas import obtener_rutas  # noqa: E402
from honorarios.utilidades import ErrorUsuario  # noqa: E402

ANIO, MES = datos_prueba.ANIO, datos_prueba.MES
# Una línea como queda en Excel al pegar el TXT: Excel convierte la fecha, el monto, el 1 y los IDs en números
COMO_EN_EXCEL = [datetime(2026, 9, 29), 145362, "A1893", 9096, "X", "E0", 1, 32635079, "50E03158", "T",
                 "A1893      Q18-2026"]


def excel(*filas, hoja="Hoja1", antes=None):
    """Excel en memoria. antes: otra hoja que va primero (como el Resumen del Excel de la conciliación)."""
    libro = Workbook()
    if antes:
        libro.active.title = antes
        libro.active.append(["no se lee"])
        destino = libro.create_sheet(hoja)
    else:
        destino = libro.active
        destino.title = hoja
    for fila in filas:
        destino.append(list(fila))
    salida = BytesIO()
    libro.save(salida)
    return salida.getvalue()


def linea(**valores):
    base = dict(zip(txt_sap.COLUMNAS, ["29.09.2026", "145362", "A1893", "9096.00", "X", "E0", "001", "32635079",
                                       "50E03158", "T", "A1893      Q18-2026"]))
    return {"Fila": 2, **base, **valores}


class TestExcelAprobado(unittest.TestCase):
    def test_regresa_lo_que_excel_cambia_al_formato_del_txt(self):
        hoja, lineas, problemas = txt_sap.leer_excel(excel(txt_sap.COLUMNAS, COMO_EN_EXCEL))
        self.assertEqual(problemas, [])
        self.assertEqual([lineas[0][c] for c in txt_sap.COLUMNAS],
                         ["29.09.2026", "145362", "A1893", "9096.00", "X", "E0", "001", "32635079", "50E03158", "T",
                          "A1893      Q18-2026"])
        self.assertEqual(lineas[0]["Fila"], 2)

    def test_encabezado_opcional_filas_vacias_y_columnas_de_apoyo(self):
        con_apoyo = COMO_EN_EXCEL + ["Cruzado", "DR. PRUEBA", 7]       # L, M, N: no se leen
        _, lineas, problemas = txt_sap.leer_excel(excel(COMO_EN_EXCEL, [], [None] * 11, con_apoyo))
        self.assertEqual((len(lineas), problemas), (2, []))
        self.assertEqual([l["Fila"] for l in lineas], [1, 4])
        self.assertNotIn("Estatus", lineas[1])

    def test_lee_la_hoja_txt_para_sap_del_excel_de_la_conciliacion(self):
        hoja, lineas, _ = txt_sap.leer_excel(excel(txt_sap.COLUMNAS, COMO_EN_EXCEL, hoja="TXT para SAP",
                                                   antes="Resumen"))
        self.assertEqual((hoja, len(lineas)), ("TXT para SAP", 1))

    def test_texto_con_formato_del_txt_tambien_sirve(self):
        valores = ["29.09.2026", "145362", "A1893", "9,096.00", "x", "e0", "001", "32635079", "50E03158", "t",
                   "A1893      Q18-2026"]
        _, lineas, problemas = txt_sap.leer_excel(excel(valores))
        self.assertEqual(problemas, [])
        self.assertEqual([lineas[0][c] for c in ("Monto", "X", "Tax", "PM")], ["9096.00", "X", "E0", "T"])

    def test_cada_problema_dice_fila_y_columna_y_no_se_corrige_solo(self):
        malas = [
            ("Fecha", "31.02.2026"), ("IDProveedor", None), ("Folio", ""), ("Monto", -5), ("Monto", "mil"),
            ("Monto", 10.555), ("X", "Y"), ("Tax", "E1"), ("1", 2), ("Centro de Costo", None),
            ("Cuenta Contable", 5e31), ("PM", "C"), ("Referencia", "A1893\tQ18"),
        ]
        filas = []
        for columna, valor in malas:
            fila = list(COMO_EN_EXCEL)
            fila[txt_sap.COLUMNAS.index(columna)] = valor
            filas.append(fila)
        _, lineas, problemas = txt_sap.leer_excel(excel(*filas))
        self.assertEqual(len(lineas), len(malas))
        self.assertEqual([(f, c) for f, c, _ in problemas], [(i, c) for i, (c, _) in enumerate(malas, start=1)])
        textos = dict(((f, c), t) for f, c, t in problemas)
        self.assertIn("E0 (sin IVA) o F2", textos[(8, "Tax")])
        self.assertIn("Texto", textos[(11, "Cuenta Contable")])
        self.assertIn("más de 2 decimales", textos[(6, "Monto")])

    def test_archivo_que_no_es_excel_o_sin_lineas(self):
        with self.assertRaises(ErrorUsuario):
            txt_sap.leer_excel(b"no es un excel")
        with self.assertRaises(ErrorUsuario) as e:
            txt_sap.leer_excel(excel(txt_sap.COLUMNAS))
        self.assertIn("de la A a la K", str(e.exception))

    def test_misma_factura_repetida_se_avisa(self):
        lineas = [linea(Fila=2), linea(Fila=3), linea(Fila=4, Monto="100.00")]
        self.assertEqual(txt_sap.misma_factura_repetida(lineas), [[2, 3]])


class TestFoliosRastreables(unittest.TestCase):
    def test_formato_de_16_caracteres(self):
        (folio, clave, previo), = folios.formatear_folios([linea(IDProveedor="30165113")], {})
        self.assertEqual(len(folio), LONGITUD_FOLIO)
        self.assertEqual(LONGITUD_FOLIO, 16, "lo más que acepta SAP")
        self.assertEqual(folio[:8], "26185113")                     # 2026, Q18, últimos 4 del IDProveedor
        self.assertTrue(all(c in folios.ALFABETO for c in folio[8:]))
        self.assertIsNone(previo)

    def test_quincena_de_la_referencia_o_de_la_fecha(self):
        self.assertEqual(folios.quincena("29.09.2026", "A1      Q18-2026"), (2026, 18))
        self.assertEqual(folios.quincena("29.09.2026", "sin quincena"), (2026, 18))
        self.assertEqual(folios.quincena("15.01.2027"), (2027, 1))
        self.assertEqual(folios.quincena("16.12.2026"), (2026, 24))

    def test_id_corto_o_con_letras(self):
        self.assertEqual(folios.cola_del_id("7"), "0007")
        self.assertEqual(folios.cola_del_id("PV0001"), "0001")
        self.assertEqual(folios.cola_del_id("145362.0"), "5362")

    def test_misma_factura_mismo_folio_y_formatear_otra_vez_no_cambia_nada(self):
        primera, _ = txt_sap.formatear([linea()], {})
        otra, _ = txt_sap.formatear([linea()], {})
        self.assertEqual(primera[0]["Folio"], otra[0]["Folio"])
        de_nuevo, _ = txt_sap.formatear(primera, {})                 # el TXT ya formateado, vuelto a subir
        self.assertEqual((de_nuevo[0]["Folio"], de_nuevo[0]["Referencia"]),
                         (primera[0]["Folio"], primera[0]["Referencia"]))

    def test_otra_factura_otro_folio(self):
        base = folios.formatear_folios([linea()], {})[0][0]
        for cambio in ({"Folio": "A1894"}, {"Monto": "9096.01"}, {"Fecha": "30.09.2026"}):
            with self.subTest(cambio):
                self.assertNotEqual(folios.formatear_folios([linea(**cambio)], {})[0][0], base)

    def test_pago_quebrado_en_partes_iguales_no_choca(self):
        asignados = folios.formatear_folios([linea(), linea()], {})
        self.assertNotEqual(asignados[0][0], asignados[1][0])

    def test_si_choca_con_otra_factura_registrada_se_recalcula(self):
        ocupado = folios.formatear_folios([linea()], {})[0][0]
        (folio, _, previo), = folios.formatear_folios([linea()], {ocupado: {"Clave": "otra factura"}})
        self.assertNotEqual(folio, ocupado)
        self.assertEqual(folio[:8], ocupado[:8])
        self.assertIsNone(previo)

    def test_si_ya_se_genero_antes_se_avisa(self):
        folio, clave, _ = folios.formatear_folios([linea()], {})[0]
        registro = {"Clave": clave, "Registrado": "02/10/2026 10:00"}
        (mismo, _, previo), = folios.formatear_folios([linea()], {folio: registro})
        self.assertEqual((mismo, previo), (folio, registro))

    def test_la_misma_linea_ya_formateada_dos_veces_no_se_deja(self):
        formateada, _ = txt_sap.formatear([linea()], {})
        _, problemas = txt_sap.formatear([formateada[0], {**formateada[0], "Fila": 3}], {})
        self.assertEqual([(f, c) for f, c, _ in problemas], [(2, "Folio"), (3, "Folio")])

    def test_millones_de_combinaciones_sin_repetir(self):
        generados = {folios.formatear_folios([linea(Folio=f"F{i}")], {})[0][0] for i in range(100_000)}
        self.assertEqual(len(generados), 100_000, "100 mil facturas del mismo proveedor y quincena: ninguna se repite")

    def test_referencia_con_el_folio_nuevo(self):
        formateada, _ = txt_sap.formatear([linea()], {})
        folio = formateada[0]["Folio"]
        self.assertEqual(formateada[0]["Referencia"], f"{folio}_Q18-2026")
        self.assertEqual(len(formateada[0]["Referencia"]), 25)
        self.assertEqual(formateada[0]["FolioOriginal"], "A1893")


class TestTxtParaSap(unittest.TestCase):
    def test_mismo_formato_que_el_txt_original(self):
        formateadas, _ = txt_sap.formatear([linea(), linea(Fila=3, IDProveedor="152091", Folio="7",
                                                           Referencia="7      Q18-2026")], {})
        datos = txt_sap.escribir(formateadas)
        renglones = datos.split(b"\r\n")
        self.assertEqual(renglones[-1], b"", "cada línea termina en CRLF")
        self.assertEqual(len(renglones), 3, "sin encabezado")
        campos = renglones[0].decode("cp1252").split("\t")
        self.assertEqual(len(campos), 11)
        f = formateadas[0]["Folio"]
        self.assertEqual(campos, ["29.09.2026", "145362", f, "9096.00", "X", "E0", "001", "32635079", "50E03158",
                                  "T", f"{f}_Q18-2026"])
        self.assertNotIn(b"\n", datos.replace(b"\r\n", b""))

    def test_registro_de_folios_para_rastrear(self):
        ruta = CONFIG / ARCHIVO_FOLIOS
        ruta.parent.mkdir(parents=True, exist_ok=True)
        ruta.write_text("Folio,Clave,IDProveedor,Periodo,Linea,Registrado\nABCDEFGHJK,U1#0,1,2026-09,1,01/09/2026\n",
                        encoding="utf-8-sig")                        # registro de la versión anterior
        formateadas, _ = txt_sap.formatear([linea()], folios.cargar_folios())
        folios.registrar_folios(txt_sap.registros_de(formateadas, "Aprobadas.xlsx"))
        folios.registrar_folios(txt_sap.registros_de(formateadas, "Aprobadas.xlsx"))     # dos descargas
        registrados = folios.cargar_folios()
        self.assertEqual(set(registrados), {"ABCDEFGHJK", formateadas[0]["Folio"]})
        r = registrados[formateadas[0]["Folio"]]
        self.assertEqual((r["FolioOriginal"], r["IDProveedor"], r["Periodo"], r["Monto"], r["Archivo"]),
                         ("A1893", "145362", "Q18-2026", "9096.00", "Aprobadas.xlsx"))
        otra_vez, _ = txt_sap.formatear([linea()], registrados)
        self.assertEqual(otra_vez[0]["Previo"]["Registrado"], r["Registrado"])
        ruta.unlink()

    def test_periodo_y_control_de_folios(self):
        formateadas, _ = txt_sap.formatear([linea(), linea(Fila=3, Fecha="01.10.2026", Folio="B1")], {})
        self.assertEqual(txt_sap.periodo(formateadas)[2], "Q18-2026")
        libro = load_workbook(BytesIO(txt_sap.control_de_folios(formateadas, {"145362": "DR. PRUEBA"})))
        hoja = libro.active
        self.assertEqual([c.value for c in hoja[2]][:5], [2, "DR. PRUEBA", "145362", "A1893", formateadas[0]["Folio"]])

    def test_plantilla_vacia(self):
        hoja = load_workbook(BytesIO(txt_sap.plantilla()))["TXT para SAP"]
        self.assertEqual([c.value for c in hoja[1]], txt_sap.COLUMNAS)


class TestHojaTxtParaSap(unittest.TestCase):
    """El Excel de la conciliación trae la hoja "TXT para SAP": se deja solo lo aprobado y se sube tal cual."""

    @classmethod
    def setUpClass(cls):
        datos_prueba.crear(RAIZ)
        cls.rutas = obtener_rutas(ANIO, MES)
        cls.conciliacion = servicio.conciliar_periodo(
            cls.rutas, ANIO, MES, hoy=datetime(ANIO, MES, 30), avisar=lambda t: None,
            progreso_sat=lambda h, t: None, consultar_sat=sat_todo_vigente, consultar_sap=sap_simulado)

    def test_trae_todas_las_lineas_del_txt_tal_cual(self):
        hoja = load_workbook(self.conciliacion.destino)["TXT para SAP"]
        filas = list(hoja.iter_rows(values_only=True))
        self.assertEqual(list(filas[0]), txt_sap.COLUMNAS + ["Estatus", "Médico", "Línea"])
        originales = txt_sap.lineas_del_txt(self.rutas["txt"].read_bytes())
        self.assertEqual(len(filas) - 1, len(originales))
        for fila in filas[1:]:
            self.assertEqual(list(fila[:11]), originales[fila[13]], "A a K: el TXT tal cual, como texto")
        self.assertEqual(hoja["A2"].number_format, "@")

    def test_dejar_solo_lo_aprobado_y_subirlo(self):
        libro = load_workbook(self.conciliacion.destino)
        hoja = libro["TXT para SAP"]
        for r in range(hoja.max_row, 1, -1):                         # el analista borra lo que no se paga
            if not str(hoja.cell(r, 12).value).startswith("Cruzado"):
                hoja.delete_rows(r)
        salida = BytesIO()
        libro.save(salida)
        nombre, lineas, problemas = txt_sap.leer_excel(salida.getvalue())
        cruzadas = [f for f in self.conciliacion.resultado if f["Estatus"].startswith("Cruzado")]
        self.assertEqual((nombre, problemas, len(lineas)), ("TXT para SAP", [], len(cruzadas)))
        formateadas, problemas = txt_sap.formatear(lineas, {})
        self.assertEqual(problemas, [])
        self.assertEqual(len({(l["IDProveedor"], l["Folio"]) for l in formateadas}), len(formateadas))
        self.assertTrue(all(len(l["Folio"]) == 16 for l in formateadas))


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
        editada = servicio.replace(self.conciliacion, version=7)
        self.assertTrue(servicio.guardar_trabajo(editada, self.carpeta, txt))
        (ruta, nombre, fecha), = servicio.trabajos_guardados(Path(self.tmp.name))
        self.assertEqual(nombre, "septiembre 2026")
        conciliacion, txt_recuperado, carpeta = servicio.cargar_trabajo(ruta)
        self.assertEqual(txt_recuperado, txt)
        self.assertEqual(carpeta, self.carpeta)
        self.assertEqual([f["Estatus"] for f in conciliacion.resultado], [f["Estatus"] for f in editada.resultado])
        self.assertEqual(conciliacion.version, 7)
        servicio.revalidar(conciliacion, conciliacion.entradas["lineas"])   # se vuelve a cruzar sin consultar

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
        danado.write_bytes(servicio.firmar(pickle.dumps({"programa": "0.0", "guardado": "", "txt": ("a.txt", b""),
                                                         "conciliacion": campos})))
        conciliacion, _, _ = servicio.cargar_trabajo(danado)
        self.assertEqual(conciliacion.consultado, "")

    def test_solo_abre_trabajos_firmados_en_esta_computadora(self):
        import os
        import pickle
        txt = ("Honorarios.txt", b"x")
        self.assertTrue(servicio.guardar_trabajo(self.conciliacion, self.carpeta, txt))
        ruta = servicio.ruta_trabajo(self.carpeta, ANIO, MES)
        firmado = ruta.read_bytes()
        self.assertTrue(firmado.startswith(servicio.ENCABEZADO_TRABAJO))

        class Trampa:                                     # lo que un atacante metería en un pickle
            def __reduce__(self):
                return (os.system, ("echo hackeado",))
        casos = {
            "sin firma (pickle suelto)": pickle.dumps(Trampa()),
            "modificado después de firmar": firmado[:-5] + b"XXXXX",
            "firma de otra computadora": servicio.ENCABEZADO_TRABAJO + b"\0" * 32 + pickle.dumps(Trampa()),
        }
        for nombre, contenido in casos.items():
            with self.subTest(nombre):
                ruta.write_bytes(contenido)
                with mock.patch("os.system") as ejecutado, self.assertRaises(servicio.ErrorUsuario):
                    servicio.cargar_trabajo(ruta)
                ejecutado.assert_not_called()             # nunca se llega a leer el contenido
        ruta.write_bytes(firmado)
        self.assertEqual(servicio.cargar_trabajo(ruta)[1], txt)

    def test_la_clave_se_crea_una_vez_y_no_cambia(self):
        clave = servicio.clave_de_esta_computadora()
        self.assertEqual(len(clave), 32)
        self.assertEqual(servicio.clave_de_esta_computadora(), clave)
        self.assertNotIn("OneDrive", str(servicio.ruta_clave()))

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

    def test_excel_regenerado_conserva_el_txt_original(self):
        from openpyxl import load_workbook
        txt = ("Honorarios.txt", self.rutas["txt"].read_bytes())
        nueva = servicio.regenerar_reporte(self.conciliacion, self.carpeta, txt)
        libro = load_workbook(nueva.destino)
        self.assertNotIn("Cambios TXT", libro.sheetnames)
        self.assertIn("TXT para SAP", libro.sheetnames)
        info = {fila[0].value: fila[1].value for fila in libro["Info"].iter_rows()}
        self.assertEqual(info["Archivo TXT"], "Honorarios.txt")
        self.assertEqual([p.suffix for p in self.carpeta.iterdir()], [".xlsx"], "no deja copias del TXT")


if __name__ == "__main__":
    unittest.main(verbosity=2)
