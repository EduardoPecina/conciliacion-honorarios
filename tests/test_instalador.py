"""Pruebas del instalador (instalador/instalar.py) con un Escritorio falso: sin pip ni accesos directos."""

import os
import re
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


class TestVersionesFijas(unittest.TestCase):
    """requirements-lock.txt: versiones exactas con las que el programa está probado. Si esta prueba falla, el entorno
    de pruebas tiene otras versiones: o se instalan las del lock, o se actualiza el lock a propósito (README)."""

    RAIZ = Path(__file__).resolve().parent.parent
    SOLO_WINDOWS = {"pywin32", "colorama", "tzdata"}

    def lock(self):
        lineas = (self.RAIZ / "requirements-lock.txt").read_text(encoding="utf-8").splitlines()
        return dict(l.split("==") for l in lineas if l.strip() and not l.startswith("#"))

    def test_todo_lo_que_se_pide_esta_fijado(self):
        lock = self.lock()
        for archivo in ("requirements.txt", "requirements-portal.txt"):
            for linea in (self.RAIZ / archivo).read_text(encoding="utf-8").splitlines():
                nombre = re.split(r"[<>=;# ]", linea.strip(), maxsplit=1)[0].lower()
                if nombre and not nombre.startswith("-"):
                    with self.subTest(archivo=archivo, libreria=nombre):
                        self.assertIn(nombre, lock)

    def test_las_pruebas_corren_con_las_versiones_fijadas(self):
        from importlib.metadata import PackageNotFoundError, version
        for nombre, fijada in self.lock().items():
            if nombre in self.SOLO_WINDOWS:
                continue
            with self.subTest(libreria=nombre):
                try:
                    self.assertEqual(version(nombre), fijada)
                except PackageNotFoundError:
                    self.skipTest(f"{nombre} no está instalado (pip install -r requirements-portal.txt "
                                  "-c requirements-lock.txt)")

    def test_el_instalador_usa_el_lock_y_python_313(self):
        self.assertEqual(instalar.PYTHON_PROBADO, (3, 13))
        texto = (self.RAIZ / "INSTALAR.bat").read_text(encoding="ascii")
        self.assertIn("py -3.13", texto)
        self.assertIn("Python313", texto)
        with mock.patch("subprocess.run", return_value=mock.Mock(returncode=0)) as pip:
            self.assertTrue(instalar.instalar_librerias())
        comando = pip.call_args.args[0]
        self.assertIn("-c", comando)
        self.assertTrue(comando[comando.index("-c") + 1].endswith("requirements-lock.txt"))
        self.assertIn("--only-binary=:all:", comando)


if __name__ == "__main__":
    unittest.main(verbosity=2)
