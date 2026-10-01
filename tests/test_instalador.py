"""Pruebas del instalador (instalador/instalar.py) con un Escritorio falso: sin pip ni accesos directos."""

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "instalador"))

import instalar  # noqa: E402


class TestInstalador(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.escritorio = Path(self.tmp.name) / "Escritorio"
        self.raiz = self.escritorio / "Honorarios"
        entorno = mock.patch.dict(os.environ, {"HONORARIOS_ESCRITORIO": str(self.escritorio)})
        entorno.start()
        self.addCleanup(entorno.stop)
        self.addCleanup(self.tmp.cleanup)

    def instalar(self):
        with mock.patch.object(instalar, "portal_abierto", return_value=False), mock.patch("builtins.print"):
            return instalar.instalar(sin_librerias=True, sin_accesos=True)

    def test_instala_estructura_simple(self):
        self.assertEqual(self.instalar(), 0)
        self.assertEqual(sorted(p.name for p in self.raiz.iterdir()),
                         ["LEEME.txt", "Reportes", "_Configuracion", "_Programa"])
        self.assertTrue((self.raiz / "_Programa" / "portal" / "app.py").exists())
        self.assertFalse((self.raiz / "_Programa" / ".git").exists())
        self.assertIn("CatalogoProveedoresTM26.xlsx", (self.raiz / "LEEME.txt").read_text(encoding="utf-8-sig"))

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

    def test_actualizar_quita_el_acceso_del_desinstalador_retirado(self):
        self.instalar()
        acceso_viejo = self.raiz / "Desinstalar Portal de Honorarios.lnk"
        acceso_viejo.write_text("lo dejó la versión 1.10.0")
        self.instalar()
        self.assertFalse(acceso_viejo.exists())

    def test_con_el_portal_abierto_no_instala(self):
        with mock.patch.object(instalar, "portal_abierto", return_value=True), \
                mock.patch.object(instalar, "avisar") as avisar:
            self.assertEqual(instalar.instalar(sin_librerias=True, sin_accesos=True), 1)
        self.assertIn("Ciérralo", avisar.call_args.args[0])
        self.assertFalse(self.raiz.exists())


if __name__ == "__main__":
    unittest.main(verbosity=2)
