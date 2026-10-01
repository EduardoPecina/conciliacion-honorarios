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


def pagina_portal_cerrado():
    """JavaScript para st.components.v1.html: reemplaza la página por un aviso claro de "El portal se cerró".
    Sin esto, al apagarse el portal Streamlit muestra "Connection error... restart it in your terminal"."""
    return """<script>
(function () {
  const w = window.parent, d = w.document;
  w.__conciliando = false;                               // que el navegador no pregunte "¿Salir del sitio?"
  d.title = "Portal de Honorarios (cerrado)";
  d.body.style.background = "#ffffff";
  // Streamlit, al perder la conexión, agrega a <body> su ventana de "Connection error": se oculta todo lo que no
  // sea este aviso, incluso lo que se agregue después.
  const estilo = d.createElement("style");
  estilo.textContent = "body > :not(#portal-cerrado) { display: none !important; }";
  d.head.appendChild(estilo);
  d.body.innerHTML = `
    <div id="portal-cerrado" style="font-family: 'Source Sans Pro', 'Segoe UI', sans-serif; max-width: 560px; margin: 15vh auto;
                padding: 28px 32px; border: 1px solid #d0d7e2; border-left: 6px solid #305496;
                border-radius: 12px; color: #262730; box-shadow: 0 6px 24px rgba(0, 0, 0, 0.08);">
      <h2 style="margin-top: 0;">El portal se cerró</h2>
      <p>Se detuvieron todos sus procesos. Ya puedes cerrar esta pestaña.</p>
      <p style="color: #6b7280;">Para volver a abrirlo: doble clic en <b>Portal de Honorarios</b>
         en tu Escritorio.</p>
    </div>`;
})();
</script>"""
