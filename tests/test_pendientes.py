"""Pruebas de "Facturas pendientes" (segunda vuelta del mes): refacturas y XML que faltaban se buscan solo entre las
líneas que no cruzaron; lo cruzado no se toca. También: marcar en el Excel lo que ya salió en un TXT.
Datos inventados, sin internet.

    python -m unittest discover -s tests -v
"""

import sys
import tempfile
import unittest
import uuid
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import test_conciliacion  # noqa: E402,F401  (define HONORARIOS_RAIZ antes de importar el programa)
import datos_prueba  # noqa: E402
from test_portal import sap_simulado, sat_todo_vigente  # noqa: E402
from honorarios import cruce, folios, servicio, txt_sap  # noqa: E402
from honorarios.config import ARCHIVO_FOLIOS, CONFIG, RAIZ  # noqa: E402
from honorarios.entradas import obtener_rutas  # noqa: E402
from honorarios.utilidades import ErrorUsuario  # noqa: E402

ANIO, MES = datos_prueba.ANIO, datos_prueba.MES
AHORA = datetime(2026, 10, 16, 11, 40)
RFC = {id_: rfc for id_, _, rfc, _ in datos_prueba.PROVEEDORES}


def xml(id_proveedor, subtotal, fecha="2026-10-05", uuid_=None, **otros):
    return (f"{id_proveedor}_{subtotal}.xml",
            datos_prueba.xml_cfdi(RFC[id_proveedor], subtotal, fecha=f"{fecha}T09:00:00",
                                  uuid=uuid_ or str(uuid.uuid4()).upper(), **otros))


class TestFacturasPendientes(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        datos_prueba.crear(RAIZ)
        cls.conciliacion = servicio.conciliar_periodo(
            obtener_rutas(ANIO, MES), ANIO, MES, hoy=datetime(ANIO, MES, 30), avisar=lambda t: None,
            progreso_sat=lambda h, t: None, consultar_sat=sat_todo_vigente, consultar_sap=sap_simulado)

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def agregar(self, *archivos, conciliacion=None, consultar_sat=sat_todo_vigente):
        carpeta = Path(self.tmp.name) / "Pendientes 16-10-2026 1140"
        carpeta.mkdir(exist_ok=True)
        rutas = []
        for nombre, contenido in archivos:
            (carpeta / nombre).write_text(contenido, encoding="utf-8")
            rutas.append(carpeta / nombre)
        return servicio.agregar_pendientes(conciliacion or self.conciliacion, rutas, Path(self.tmp.name),
                                           avisar=lambda t: None, progreso_sat=lambda h, t: None,
                                           consultar_sat=consultar_sat, ahora=AHORA)

    @staticmethod
    def fila(conciliacion, id_proveedor):
        return next(f for f in conciliacion.resultado if f["IDProveedor"] == id_proveedor)

    def test_refactura_de_una_rechazada_y_xml_que_faltaba(self):
        self.assertEqual(self.fila(self.conciliacion, "100008")["Estatus"], "Rechazar")    # factura de agosto
        self.assertEqual(self.fila(self.conciliacion, "100005")["Estatus"], "Sin XML")
        nueva, resumen = self.agregar(xml("100008", "1000", fecha="2026-10-05"),           # refactura de octubre
                                      xml("100005", "40180", fecha="2026-09-20"))           # el XML que faltaba
        refactura = self.fila(nueva, "100008")
        self.assertEqual(refactura["Estatus"], "Cruzado con observaciones")
        self.assertIn("posterior al periodo", refactura["Observaciones"])
        self.assertIn("16/10/2026", refactura["Observaciones"])
        self.assertTrue(refactura["ArchivoXML"].startswith("Pendientes 16-10-2026 1140"), "se sabe en qué vuelta")
        self.assertEqual(self.fila(nueva, "100005")["Estatus"], "Cruzado")
        self.assertEqual({(c["Antes"], c["Ahora"]) for c in resumen["cambios"]},
                         {("Rechazar", "Cruzado con observaciones"), ("Sin XML", "Cruzado")})
        self.assertGreater(nueva.version, self.conciliacion.version)
        self.assertIn("Facturas pendientes (vuelta 2)", nueva.estados)
        # La factura de agosto que se rechazó queda en "XML sin TXT"
        self.assertTrue(any(s["FechaXML"].startswith("2026-08") for s in nueva.sobrantes))

    def test_lo_cruzado_no_se_toca(self):
        nueva, resumen = self.agregar(xml("100001", "32188", fecha="2026-09-10"))   # mismo médico e importe
        antes = [f for f in self.conciliacion.resultado if servicio.es_cruzada(f)]
        despues = [f for f in nueva.resultado if f["Linea"] in {a["Linea"] for a in antes}]
        self.assertEqual(despues, antes, "las líneas cruzadas quedan idénticas, con su misma factura")
        self.assertEqual(resumen["sin_linea"], 1, "el XML nuevo no tiene línea pendiente: va a XML sin TXT")

    def test_el_cruce_completo_cuadra(self):
        nueva, _ = self.agregar(xml("100008", "1000"), xml("100005", "40180", fecha="2026-09-20"),
                                xml("100001", "32188", fecha="2026-09-10"))
        e = nueva.entradas
        cruce.verificar_integridad(e["lineas"], e["facturas"], nueva.resultado, nueva.sobrantes)
        self.assertEqual(len(nueva.resultado), len(self.conciliacion.resultado))

    def test_una_copia_de_un_xml_ya_pagado_no_cruza_otra_linea(self):
        e = self.conciliacion.entradas
        pagada = self.fila(self.conciliacion, "100001")
        copia = {**e["facturas"][pagada["_grupo"][0]], "ArchivoXML": "copia.xml"}
        con_copia = servicio.replace(self.conciliacion, entradas={**e, "facturas": e["facturas"] + [copia]},
                                     sobrantes=self.conciliacion.sobrantes + [{"UUID": copia["UUID"]}])
        nueva, _ = self.agregar(xml("100005", "40180", fecha="2026-09-20"), conciliacion=con_copia)
        usadas = [i for f in nueva.resultado for i in f.get("_grupo", ())]
        self.assertNotIn(len(e["facturas"]), usadas, "la copia no cruza ninguna línea")

    def test_refactura_de_dos_meses_despues_se_sigue_rechazando(self):
        nueva, _ = self.agregar(xml("100008", "1000", fecha="2026-11-03"))
        self.assertEqual(self.fila(nueva, "100008")["Estatus"], "Rechazar")
        noviembre = next(x for x in nueva.sobrantes if x["FechaXML"].startswith("2026-11"))
        self.assertIn("se acepta hasta el mes siguiente", noviembre["MotivoRechazo"])

    def test_en_la_primera_vuelta_la_regla_del_periodo_no_cambia(self):
        self.assertIn("fuera del periodo", self.fila(self.conciliacion, "100008")["MotivoRechazo"])

    def test_xml_que_ya_estaba_se_ignora(self):
        e = self.conciliacion.entradas
        ya = e["facturas"][0]
        with self.assertRaises(ErrorUsuario) as error:
            self.agregar(xml("100001", "32188", uuid_=ya["UUID"]))
        self.assertIn("ya estaban", str(error.exception))
        nueva, resumen = self.agregar(xml("100001", "32188", uuid_=ya["UUID"]),
                                      xml("100005", "40180", fecha="2026-09-20"))
        self.assertEqual((len(resumen["ya_estaban"]), resumen["nuevas"]), (1, 1))

    def test_sin_conexion_con_el_sat_no_se_agrega_nada(self):
        sin_internet = lambda facturas, **k: ({f["UUID"]: {"SinConexion": True} for f in facturas}, "x")
        with self.assertRaises(ErrorUsuario):
            self.agregar(xml("100005", "40180"), consultar_sat=sin_internet)

    def test_actualizar_sat_despues_conserva_la_refactura(self):
        nueva, _ = self.agregar(xml("100008", "1000"))
        otra, _ = servicio.actualizar_sat(nueva, consultar_sat=sat_todo_vigente)
        self.assertEqual(self.fila(otra, "100008")["Estatus"], "Cruzado con observaciones")

    def test_se_guarda_y_se_puede_continuar(self):
        nueva, _ = self.agregar(xml("100008", "1000"))
        carpeta = Path(self.tmp.name) / "2026" / "09 Sep"
        self.assertTrue(servicio.guardar_trabajo(nueva, carpeta, ("Honorarios.txt", b"x")))
        recuperada, _, _ = servicio.cargar_trabajo(servicio.ruta_trabajo(carpeta, ANIO, MES))
        self.assertEqual(len(recuperada.agregadas), 1)
        self.assertEqual(self.fila(recuperada, "100008")["Estatus"], "Cruzado con observaciones")


class TestTxtGenerado(unittest.TestCase):
    """El Excel marca las líneas que ya salieron en un TXT de Generar TXT: la segunda vuelta no las repite."""

    @classmethod
    def setUpClass(cls):
        datos_prueba.crear(RAIZ)
        cls.rutas = obtener_rutas(ANIO, MES)
        cls.conciliacion = servicio.conciliar_periodo(
            cls.rutas, ANIO, MES, hoy=datetime(ANIO, MES, 30), avisar=lambda t: None,
            progreso_sat=lambda h, t: None, consultar_sat=sat_todo_vigente, consultar_sap=sap_simulado)

    def tearDown(self):
        (CONFIG / ARCHIVO_FOLIOS).unlink(missing_ok=True)

    def generar_txt_de(self, lineas_del_txt):
        """Como Generar TXT con un Excel de esas líneas del TXT original."""
        originales = txt_sap.lineas_del_txt(self.rutas["txt"].read_bytes())
        lineas = []
        for fila, n in enumerate(lineas_del_txt, start=2):
            valores = dict(zip(txt_sap.COLUMNAS_EXCEL, originales[n]))
            valores["Monto"] = f"{float(valores['Monto']):.2f}"
            lineas.append({"Fila": fila, **valores, "Referencia": ""})
        formateadas, _ = txt_sap.formatear(lineas, folios.cargar_folios())
        folios.registrar_folios(txt_sap.registros_de(formateadas, "Primera vuelta.xlsx"))

    def test_las_lineas_con_txt_quedan_marcadas(self):
        cruzadas = [f["Linea"] for f in self.conciliacion.resultado if servicio.es_cruzada(f)]
        self.generar_txt_de(cruzadas[:3])
        marcadas = folios.ya_generados(self.conciliacion.resultado)
        self.assertEqual(sorted(marcadas), sorted(cruzadas[:3]))
        self.assertIn("Primera vuelta.xlsx", next(iter(marcadas.values())))

    def test_lineas_identicas_se_marcan_una_por_una(self):
        quebradas = [f["Linea"] for f in self.conciliacion.resultado if f["IDProveedor"] == "100003"]
        self.generar_txt_de(quebradas[:1])                      # 3 líneas del mismo médico, importes distintos
        self.assertEqual(sorted(folios.ya_generados(self.conciliacion.resultado)), quebradas[:1])

    def test_la_hoja_txt_para_sap_lo_muestra(self):
        cruzadas = [f["Linea"] for f in self.conciliacion.resultado if servicio.es_cruzada(f)]
        self.generar_txt_de(cruzadas[:2])
        nueva = servicio.regenerar_reporte(self.conciliacion, Path(tempfile.mkdtemp()),
                                           ("Honorarios.txt", self.rutas["txt"].read_bytes()))
        hoja = load_workbook(nueva.destino)["TXT para SAP"]
        filas = list(hoja.iter_rows(min_row=2, values_only=True))
        con_txt = {f[13] for f in filas if f[11]}
        self.assertEqual(con_txt, set(cruzadas[:2]))
        cruzado = [f for f in filas if f[10] == "Cruzado"]
        self.assertFalse(any(f[11] for f in cruzado[:len(cruzado) - 2]), "primero lo que todavía no tiene TXT")
        detalle = load_workbook(nueva.destino)["Detalle"]
        encabezados = [c.value for c in detalle[1]]
        self.assertIn("TXT generado", encabezados)


if __name__ == "__main__":
    unittest.main(verbosity=2)
