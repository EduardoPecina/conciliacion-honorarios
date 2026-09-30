"""Detalles visuales del portal: logo de la empresa y aviso al salir de la página mientras se procesa."""

import json
import time
from pathlib import Path

RECURSOS = Path(__file__).parent / "recursos"


def ruta_logo():
    """Logo de la empresa si existe portal/recursos/logo.png (o .jpg/.jpeg/.svg); si no, None."""
    for extension in (".png", ".jpg", ".jpeg", ".svg"):
        ruta = RECURSOS / f"logo{extension}"
        if ruta.exists():
            return ruta
    return None


MENSAJE_SALIR = ("La conciliación sigue en proceso.\n\n"
                 "Si sales de esta página ahora, el proceso se detiene y se pierde el avance.\n\n"
                 "¿Quieres salir de todas formas?")


def aviso_al_salir(activo):
    """HTML+JavaScript para st.components.v1.html: mientras activo=True, al dar clic en otra página del menú
    (o en "Cerrar el portal") el navegador pregunta si de verdad se quiere salir; al cerrar o recargar la
    pestaña, el navegador muestra su aviso de "¿Salir del sitio?". NO bloquea nada: se puede hacer scroll y la
    decisión final es del usuario. Las funciones se crean en la página principal (window.parent), no en el
    iframe del componente, para que sigan funcionando cuando el iframe se quite."""
    return f"""<script>
(function () {{
  const w = window.parent;
  w.__conciliando = {"true" if activo else "false"};
  w.__versionAviso = {time.time_ns()};                   // cada llamada es distinta: el iframe siempre se ejecuta
  if (w.__avisoInstalado) return;
  w.__avisoInstalado = true;
  const alClic = new w.Function("e", `
    if (!window.__conciliando) return;
    const el = e.target.closest('a[data-testid="stSidebarNavLink"], button');
    if (!el) return;
    if (el.tagName === "BUTTON" && !el.innerText.includes("Cerrar el portal")) return;
    if (window.confirm(${{JSON.stringify({json.dumps(MENSAJE_SALIR)})}})) {{ window.__conciliando = false; return; }}
    e.preventDefault(); e.stopImmediatePropagation();
  `);
  w.document.addEventListener("click", alClic, true);
  w.addEventListener("beforeunload", new w.Function("e",
    'if (window.__conciliando) {{ e.preventDefault(); e.returnValue = ""; }}'));
}})();
</script>"""
