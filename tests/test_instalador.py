"""Pruebas del instalador y del desinstalador (instalador/) con un Escritorio falso: sin pip ni accesos directos."""

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "instalador"))

import desinstalar  # noqa: E402
import instalar  # noqa: E402


class ConEscritorioFalso(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.escritorio = Path(self.tmp.name) / "Escritorio"
        self.raiz = self.escritorio / "Honorarios"
        entorno = mock.patch.dict(os.environ, {"HONORARIOS_ESCRITORIO": str(self.escritorio)})
        entorno.start()
        self.addCleanup(entorno.stop)

    def instalar(self):
        with mock.patch.object(instalar, "portal_abierto", return_value=False), mock.patch("builtins.print"):
            return instalar.instalar(sin_librerias=True, sin_accesos=True)


class TestInstalador(ConEscritorioFalso):
    def test_instala_estructura_simple(self):
        self.assertEqual(self.instalar(), 0)
        self.assertEqual(sorted(p.name for p in self.raiz.iterdir()),
                         ["LEEME.txt", "Reportes", "_Configuracion", "_Programa"])
        self.assertTrue((self.raiz / "_Programa" / "portal" / "app.py").exists())
        self.assertTrue((self.raiz / "_Programa" / "DESINSTALAR.bat").exists())
        self.assertFalse((self.raiz / "_Programa" / ".git").exists())
        leeme = (self.raiz / "LEEME.txt").read_text(encoding="utf-8-sig")
        self.assertIn("CatalogoProveedoresTM26.xlsx", leeme)
        self.assertIn("Desinstalar Portal de Honorarios", leeme)

    def test_actualizar_no_toca_catalogo_historial_ni_reportes(self):
        self.instalar()
        catalogo = self.raiz / "_Configuracion" / "CatalogoProveedoresTM26.xlsx"
        historial = self.raiz / "_Configuracion" / "HistorialUUID.csv"
        reporte = self.raiz / "Reportes" / "2026" / "09 Sep" / "Conciliacion.xlsx"
        reporte.parent.mkdir(parents=True)
        for archivo in (catalogo, historial, reporte):
            archivo.write_text("datos del usuario")
        viejo = self.raiz / "_Programa" / "modulo_de_una_version_vieja.py"
        viejo.write_text("x")

        self.assertEqual(self.instalar(), 0)
        for archivo in (catalogo, historial, reporte):
            self.assertEqual(archivo.read_text(), "datos del usuario")
        self.assertFalse(viejo.exists(), "al actualizar se reemplaza el programa completo")
        self.assertFalse((self.raiz / "_Programa_anterior").exists())

    def test_con_el_portal_abierto_no_instala(self):
        with mock.patch.object(instalar, "portal_abierto", return_value=True), \
                mock.patch.object(instalar, "avisar") as avisar:
            self.assertEqual(instalar.instalar(sin_librerias=True, sin_accesos=True), 1)
        self.assertIn("Ciérralo", avisar.call_args.args[0])
        self.assertFalse(self.raiz.exists())


class TestDesinstalador(ConEscritorioFalso):
    """Instala en el Escritorio falso y desinstala respondiendo a las preguntas."""

    def setUp(self):
        super().setUp()
        self.instalar()
        self.datos = [self.raiz / "_Configuracion" / "CatalogoProveedoresTM26.xlsx",
                      self.raiz / "_Configuracion" / "HistorialUUID.csv",
                      self.raiz / "Reportes" / "2026" / "09 Sep" / "Conciliacion.xlsx"]
        for archivo in self.datos:
            archivo.parent.mkdir(parents=True, exist_ok=True)
            archivo.write_text("datos del usuario")
        self.accesos = [self.escritorio / "Portal de Honorarios.lnk", self.raiz / "Portal de Honorarios.lnk",
                        self.raiz / "Desinstalar Portal de Honorarios.lnk"]
        for acceso in self.accesos:                           # en Windows los crea el instalador
            acceso.write_text("acceso")
        self.avisar = mock.patch.object(desinstalar, "avisar").start()
        mock.patch("time.sleep").start()
        mock.patch.object(desinstalar, "portal_abierto", return_value=False).start()
        self.addCleanup(mock.patch.stopall)

    def desinstalar(self, *respuestas):
        preguntas = iter(respuestas)
        return desinstalar.desinstalar(self.raiz, preguntar=lambda mensaje, con_cancelar=False: next(preguntas))

    def test_conserva_catalogo_historial_y_reportes(self):
        self.assertEqual(self.desinstalar(desinstalar.SI, desinstalar.SI), 0)
        self.assertFalse((self.raiz / "_Programa").exists())
        self.assertFalse((self.raiz / "LEEME.txt").exists())
        self.assertFalse(any(a.exists() for a in self.accesos))
        self.assertTrue(all(a.read_text() == "datos del usuario" for a in self.datos))
        self.assertIn("se conservaron", self.avisar.call_args.args[0])

    def test_borra_todo_con_doble_confirmacion(self):
        self.assertEqual(self.desinstalar(desinstalar.SI, desinstalar.NO, desinstalar.SI), 0)
        self.assertFalse(self.raiz.exists())
        self.assertFalse(any(a.exists() for a in self.accesos))

    def test_nunca_borra_archivos_ajenos(self):
        ajeno = self.raiz / "mis notas.docx"
        ajeno.write_text("del usuario")
        self.desinstalar(desinstalar.SI, desinstalar.NO, desinstalar.SI)
        self.assertEqual(ajeno.read_text(), "del usuario")
        self.assertEqual([p.name for p in self.raiz.iterdir()], ["mis notas.docx"])

    def test_cancelar_no_borra_nada(self):
        for respuestas in ((desinstalar.NO,), (desinstalar.SI, desinstalar.CANCELAR),
                           (desinstalar.SI, desinstalar.NO, desinstalar.NO)):     # la última: no confirma
            with self.subTest(respuestas=respuestas):
                self.assertEqual(self.desinstalar(*respuestas), 0)
                self.assertTrue((self.raiz / "_Programa" / "portal" / "app.py").exists())
                self.assertTrue(all(a.exists() for a in self.datos))
                self.assertTrue(all(a.exists() for a in self.accesos))

    def test_con_el_portal_abierto_no_desinstala(self):
        with mock.patch.object(desinstalar, "portal_abierto", return_value=True):
            self.assertEqual(self.desinstalar(), 1)
        self.assertTrue((self.raiz / "_Programa").exists())

    def test_reinstalar_despues_de_desinstalar_recupera_los_datos(self):
        self.desinstalar(desinstalar.SI, desinstalar.SI)
        self.assertEqual(self.instalar(), 0)
        self.assertTrue((self.raiz / "_Programa" / "portal" / "app.py").exists())
        self.assertTrue(all(a.read_text() == "datos del usuario" for a in self.datos))


if __name__ == "__main__":
    unittest.main(verbosity=2)
