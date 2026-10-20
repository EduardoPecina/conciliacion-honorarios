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
                         ["LEEME.txt", "Reportes", "_Configuracion", "_Programa", "instalacion.log"])
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

    def test_dentro_del_zip_sin_poder_copiarlo_no_instala(self):
        with mock.patch.object(instalar, "corre_desde_el_zip", return_value=True), \
                mock.patch.object(instalar, "avisar") as aviso:
            self.assertEqual(self.instalar(), 1)
        self.assertIn("Extraer todo", aviso.call_args.args[0])
        self.assertFalse((self.raiz / "_Programa").exists())

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
        self.assertFalse((self.raiz / "_Programa").exists())
        self.assertIn("RESULTADO: NO se completó", (self.raiz / "instalacion.log").read_text(encoding="utf-8"))

    # ---- Bitácora (instalacion.log)
    def test_la_bitacora_registra_cada_instalacion(self):
        self.instalar()
        self.instalar()
        texto = (self.raiz / "instalacion.log").read_text(encoding="utf-8")
        self.assertEqual(texto.count("INSTALACIÓN "), 2, "cada instalación se agrega al final")
        self.assertEqual(texto.count("RESULTADO: terminó bien"), 2)
        self.assertIn(f"Versión del programa: {instalar.version_del_programa()}", texto)
        self.assertIn("Python: ", texto)
        self.assertNotEqual(instalar.version_del_programa(), "desconocida")

    def test_lo_que_se_muestra_en_la_ventana_queda_en_la_bitacora(self):
        import io
        consola = io.StringIO()
        with mock.patch.object(instalar, "portal_abierto", return_value=False), mock.patch("sys.stdout", consola):
            instalar.instalar(sin_librerias=True, sin_accesos=True)
            self.assertIs(sys.stdout, consola, "al terminar, la consola vuelve a ser la de antes")
        texto = (self.raiz / "instalacion.log").read_text(encoding="utf-8")
        self.assertIn("Instalando en:", consola.getvalue())
        self.assertIn("Instalando en:", texto)
        self.assertIn("Instalación terminada", texto)

    def test_incluye_los_pasos_de_instalar_bat(self):
        bat = Path(self.tmp.name) / "bat.log"
        bat.write_text("INSTALAR.bat 03/10/2026\nPython 3.13 encontrado: py -3.13\n", encoding="utf-8")
        with mock.patch.dict(os.environ, {"HONORARIOS_BITACORA_BAT": str(bat)}):
            self.instalar()
        texto = (self.raiz / "instalacion.log").read_text(encoding="utf-8")
        self.assertIn("-- INSTALAR.bat --", texto)
        self.assertIn("Python 3.13 encontrado: py -3.13", texto)

    def test_un_error_inesperado_queda_registrado(self):
        with mock.patch.object(instalar, "portal_abierto", return_value=False), mock.patch("builtins.print"), \
                mock.patch.object(instalar, "crear_estructura", side_effect=RuntimeError("disco lleno")), \
                mock.patch.object(instalar, "avisar") as avisar:
            self.assertEqual(instalar.instalar(sin_librerias=True, sin_accesos=True), 1)
        texto = (self.raiz / "instalacion.log").read_text(encoding="utf-8")
        self.assertIn("RuntimeError: disco lleno", texto)
        self.assertIn("RESULTADO: NO se completó", texto)
        self.assertTrue(avisar.call_args.kwargs.get("error"))

    def test_los_errores_piden_mandar_la_bitacora(self):
        with mock.patch("builtins.print") as impreso:
            instalar.avisar("Algo falló.", error=True)
            instalar.avisar("Todo bien.")
        self.assertIn("instalacion.log", impreso.call_args_list[0].args[0])
        self.assertNotIn("instalacion.log", impreso.call_args_list[1].args[0])

    def test_la_bitacora_no_crece_sin_limite(self):
        self.raiz.mkdir(parents=True)
        bitacora = self.raiz / "instalacion.log"
        bitacora.write_text(("\n" + "=" * 78 + "\nINSTALACIÓN vieja\n" + "x" * 1000) * 3000, encoding="utf-8")
        self.instalar()
        texto = bitacora.read_text(encoding="utf-8")
        self.assertLess(len(texto.encode("utf-8")), instalar.TAMANO_MAXIMO)
        self.assertTrue(texto.startswith("====="), "se corta en el inicio de una instalación")
        self.assertIn("RESULTADO: terminó bien", texto)


class TestRespaldoDeLaActualizacion(unittest.TestCase):
    """_Programa_anterior: el respaldo que deja una actualización no debe quedarse a la vista del usuario."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.raiz = Path(self.tmp.name)

    def programa(self, nombre, archivos=("honorarios/__init__.py", "portal/app.py", "README.md")):
        carpeta = self.raiz / nombre
        for archivo in archivos:
            (carpeta / archivo).parent.mkdir(parents=True, exist_ok=True)
            (carpeta / archivo).write_text("x")
        return carpeta

    def test_borra_aunque_haya_archivos_de_solo_lectura(self):
        import stat
        carpeta = self.programa("_Programa_anterior")
        os.chmod(carpeta / "README.md", stat.S_IREAD)
        self.assertTrue(instalar.borrar_carpeta(carpeta, intentos=2, espera=0))
        self.assertFalse(carpeta.exists())

    def test_si_onedrive_no_deja_borrar_el_respaldo_viejo_usa_otro_nombre(self):
        destino = self.programa("_Programa")
        self.programa("_Programa_anterior")                       # quedó bloqueado de una actualización pasada
        origen = self.programa("zip_nuevo")
        with mock.patch.object(instalar, "ORIGEN", origen), mock.patch("builtins.print"), \
                mock.patch.object(instalar, "borrar_carpeta", return_value=False):
            instalar.copiar_programa(destino)                     # antes: FileExistsError al renombrar
        self.assertTrue((self.raiz / "_Programa_anterior_2").exists())
        self.assertTrue((destino / "portal" / "app.py").exists())

    def test_una_copia_incompleta_deja_la_version_anterior(self):
        destino = self.programa("_Programa", archivos=("honorarios/__init__.py", "portal/app.py", "VIEJO.txt"))
        origen = self.programa("zip_nuevo")

        def copia_vacia(_origen, hacia, **_):
            hacia.mkdir()                                         # lo que se vio en OneDrive: _Programa con 0 elementos
        with mock.patch.object(instalar, "ORIGEN", origen), mock.patch("shutil.copytree", copia_vacia), \
                mock.patch("builtins.print"):
            with self.assertRaises(instalar.ProgramaIncompleto):
                instalar.copiar_programa(destino)
        self.assertTrue((destino / "VIEJO.txt").exists(), "se debe quedar la versión anterior completa")
        self.assertEqual(instalar.respaldos(destino), [])

    def test_la_copia_completa_queda_revisada(self):
        destino = self.programa("_Programa")
        origen = self.programa("zip_nuevo", archivos=("honorarios/__init__.py", "portal/app.py", "nuevo.py",
                                                      "portal/__pycache__/x.pyc"))
        with mock.patch.object(instalar, "ORIGEN", origen), mock.patch("builtins.print"):
            instalar.copiar_programa(destino)
            self.assertEqual(instalar.faltantes(destino), [])
        self.assertTrue((destino / "nuevo.py").exists())
        self.assertFalse((destino / "README.md").exists())
        self.assertFalse((destino / "portal" / "__pycache__").exists())

    def test_detecta_cuando_se_corre_dentro_del_zip(self):
        temporal = Path(tempfile.gettempdir())
        zip_abierto = temporal / "d08f8439_conciliacion-honorarios-main.zip.127" / "conciliacion-honorarios-main"
        self.assertTrue(instalar.corre_desde_el_zip(zip_abierto))
        self.assertFalse(instalar.corre_desde_el_zip(temporal / "conciliacion-honorarios-main"))
        self.assertFalse(instalar.corre_desde_el_zip(Path.home() / "Downloads" / "conciliacion-honorarios-main"))

    def zip_abierto(self, archivos=instalar.ESENCIALES + ("README.md", "portal/__pycache__/x.pyc")):
        """Lo que deja Windows al abrir INSTALAR.bat dentro del ZIP: la carpeta extraída en Temp."""
        origen = self.programa("x_conciliacion-honorarios-main.zip.686/conciliacion-honorarios-main", archivos)
        estable = self.raiz / "AppData" / "Honorarios" / "Instalador"
        bitacora_bat = self.raiz / "InstalacionHonorarios_bat.log"
        bitacora_bat.write_text("INSTALAR.bat ...\n", encoding="utf-8")
        entorno = mock.patch.dict(os.environ, {"HONORARIOS_INSTALADOR": str(estable),
                                               "HONORARIOS_BITACORA_BAT": str(bitacora_bat)})
        entorno.start()
        self.addCleanup(entorno.stop)
        return origen, estable / origen.name, bitacora_bat

    def test_abierto_dentro_del_zip_se_copia_a_un_lugar_fijo_y_sigue_desde_ahi(self):
        origen, copia, bitacora_bat = self.zip_abierto()
        (copia / "viejo.txt").parent.mkdir(parents=True)
        (copia / "viejo.txt").write_text("de una instalación anterior")
        with mock.patch.object(instalar, "ORIGEN", origen), mock.patch("builtins.print"), \
                mock.patch("subprocess.call", return_value=0) as corre:
            self.assertEqual(instalar.instalar_desde_el_zip(["--sin-accesos"]), 0)
        self.assertEqual(corre.call_args.args[0],
                         [sys.executable, str(copia / "instalador" / "instalar.py"), "--sin-accesos"])
        self.assertEqual(corre.call_args.kwargs["cwd"], copia)
        self.assertTrue((copia / "portal" / "app.py").exists())
        self.assertFalse((copia / "viejo.txt").exists(), "la copia anterior se reemplaza")
        self.assertFalse((copia / "portal" / "__pycache__").exists())
        self.assertIn("se copió a", bitacora_bat.read_text(encoding="utf-8"))

    def test_si_el_zip_no_se_extrajo_completo_pide_extraerlo(self):
        origen, copia, bitacora_bat = self.zip_abierto(archivos=("INSTALAR.bat", "instalador/instalar.py"))
        with mock.patch.object(instalar, "ORIGEN", origen), mock.patch("builtins.print"), \
                mock.patch("subprocess.call") as corre:
            self.assertIsNone(instalar.instalar_desde_el_zip([]))
        corre.assert_not_called()
        self.assertIn("no se extrajo completo", bitacora_bat.read_text(encoding="utf-8"))

    def test_instalacion_completa_abriendo_el_zip(self):
        """De punta a punta, como en Windows: INSTALAR.bat desde la carpeta temporal del ZIP."""
        import shutil
        import subprocess
        repo = Path(__file__).resolve().parent.parent
        with tempfile.TemporaryDirectory(dir=tempfile.gettempdir()) as temporal:
            origen = Path(temporal) / "b56f_conciliacion-honorarios-main.zip.686" / "conciliacion-honorarios-main"
            shutil.copytree(repo, origen, ignore=instalar.NO_COPIAR)
            escritorio = self.raiz / "Escritorio"
            estable = self.raiz / "AppData" / "Honorarios" / "Instalador"
            entorno = {**os.environ, "HONORARIOS_ESCRITORIO": str(escritorio), "HONORARIOS_INSTALADOR": str(estable),
                       "PYTHONIOENCODING": "utf-8"}
            entorno.pop("HONORARIOS_BITACORA_BAT", None)
            hecho = subprocess.run([sys.executable, str(origen / "instalador" / "instalar.py"), "--sin-librerias",
                                    "--sin-accesos"], cwd=origen, env=entorno, capture_output=True, text=True,
                                   encoding="utf-8", stdin=subprocess.DEVNULL, timeout=120)
            self.assertEqual(hecho.returncode, 0, hecho.stdout + hecho.stderr)
            programa = escritorio / "Honorarios" / "_Programa"
            self.assertEqual(instalar.archivos_del_programa(programa), instalar.archivos_del_programa(origen))
        bitacora = (escritorio / "Honorarios" / instalar.BITACORA).read_text(encoding="utf-8")
        self.assertIn(f"desde {estable / origen.name}", bitacora)
        self.assertIn("RESULTADO: terminó bien", bitacora)

    def test_el_portal_borra_los_respaldos_que_quedaron(self):
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
        from portal import limpieza
        programa = self.programa("_Programa")
        viejo = self.programa("_Programa_anterior")
        a_medias = self.programa("_Programa_anterior_2", archivos=("portal/app.py",))
        ajeno = self.programa("_Programa_anterior_3", archivos=("mi_excel.xlsx",))    # no es del programa: no se toca
        limpieza.limpiar_respaldos(programa, en_segundo_plano=False)
        self.assertFalse(viejo.exists())
        self.assertFalse(a_medias.exists())
        self.assertTrue(ajeno.exists())
        self.assertTrue(programa.exists())

    def test_en_una_carpeta_de_desarrollo_no_borra_nada(self):
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
        from portal import limpieza
        repo = self.programa("conciliacion-honorarios")
        self.programa("conciliacion-honorarios_anterior")
        self.assertEqual(limpieza.respaldos_viejos(repo), [])


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
        with mock.patch.object(instalar, "ejecutar", return_value=0) as pip, mock.patch("builtins.print"):
            self.assertTrue(instalar.instalar_librerias())
        comando = pip.call_args.args[0]
        self.assertTrue(comando[comando.index("-r") + 1].endswith("requirements-instalar.txt"))
        self.assertIn("--require-hashes", comando)
        self.assertIn("--only-binary=:all:", comando)
        self.assertIn("HONORARIOS_BITACORA_BAT", texto, "INSTALAR.bat anota sus pasos para la bitácora")

    def test_el_python_de_respaldo_se_verifica_antes_de_ejecutarlo(self):
        """Sin winget, INSTALAR.bat baja Python de python.org: solo HTTPS, sin guardar páginas de error y solo se
        ejecuta si trae la firma digital válida de la Python Software Foundation."""
        texto = (self.RAIZ / "INSTALAR.bat").read_text(encoding="ascii")
        descarga = next(l for l in texto.splitlines() if l.startswith("curl.exe"))
        self.assertIn("--fail", descarga)
        self.assertIn("--proto =https", descarga)
        self.assertIn("https://www.python.org/ftp/python/%PYVERSION%/", descarga)
        version = re.search(r'set "PYVERSION=3\.13\.(\d+)"', texto)
        self.assertTrue(version and int(version.group(1)) >= 7, "una 3.13 con los parches de seguridad recientes")
        self.assertIn("Get-AuthenticodeSignature", texto)
        self.assertIn("O=Python Software Foundation", texto)
        self.assertLess(texto.index("Get-AuthenticodeSignature"), texto.index('"%INSTALADOR_PY%" /quiet'),
                        "la firma se revisa ANTES de ejecutar el instalador")
        self.assertIn("if errorlevel 1 goto :firma_invalida", texto)
        self.assertNotIn("Bypass", texto, "no se salta la política de ejecución de PowerShell")

    def test_las_huellas_corresponden_a_las_versiones_fijadas(self):
        """requirements-instalar.txt (con huellas) debe tener exactamente las versiones de requirements-lock.txt.
        Si falla: python instalador/generar_lock.py"""
        texto = (self.RAIZ / "requirements-instalar.txt").read_text(encoding="utf-8")
        bloques = re.findall(r"^([A-Za-z0-9_.-]+)==([^ ;\\\n]+)[^\n]*\n((?:\s+--hash=sha256:[0-9a-f]{64}.*\n?)+)",
                             texto, flags=re.M)
        con_huellas = {nombre.lower(): (version, len(re.findall("--hash", huellas)))
                       for nombre, version, huellas in bloques}
        self.assertEqual({n: v for n, (v, _) in con_huellas.items()}, self.lock())
        for nombre, (_, cuantas) in con_huellas.items():
            with self.subTest(libreria=nombre):
                self.assertGreater(cuantas, 0)

    def test_ejecutar_muestra_y_registra_la_salida_de_pip(self):
        import io
        consola = io.StringIO()
        with mock.patch("sys.stdout", consola):
            codigo = instalar.ejecutar([sys.executable, "-c", "print('Successfully installed algo-1.0')"])
        self.assertEqual(codigo, 0)
        self.assertIn("Successfully installed algo-1.0", consola.getvalue())


if __name__ == "__main__":
    unittest.main(verbosity=2)
