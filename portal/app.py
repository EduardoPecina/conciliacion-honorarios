"""Portal de Honorarios (Streamlit). Uso: doble clic en EJECUTAR_PORTAL.bat (sin ventana negra; abre el navegador).
Para ver los mensajes en una consola, desde la carpeta del repositorio:

    python -m streamlit run portal/app.py

REGLA: el portal NO tiene reglas de negocio. Solo recibe archivos, llama al motor (honorarios/) y muestra
el resultado. Cualquier regla nueva va en honorarios/, así el portal y el .bat siempre dan lo mismo."""

import sys
from pathlib import Path

import streamlit as st

RAIZ_REPO = Path(__file__).resolve().parents[1]
if str(RAIZ_REPO) not in sys.path:
    sys.path.insert(0, str(RAIZ_REPO))           # para importar honorarios/ y portal/

from honorarios import __version__  # noqa: E402
from portal.apagado import cerrar_portal, vigilar_inactividad  # noqa: E402
from portal.estilo import pagina_portal_cerrado, ruta_logo  # noqa: E402

st.set_page_config(page_title="Portal de Honorarios", page_icon="🩺", layout="wide")

if ruta_logo():
    st.logo(str(ruta_logo()), size="large")

paginas = Path(__file__).parent / "paginas"
# Catálogo e Historial de UUID no se muestran (Contabilidad no los usa); sus páginas siguen en paginas/ y el motor
# sigue usando el catálogo y el historial igual que antes.
navegacion = st.navigation([
    st.Page(paginas / "conciliar.py", title="Conciliador Automático", icon="✅", default=True),
    st.Page(paginas / "edicion_txt.py", title="Edición TXT", icon="✏️"),
])


@st.cache_resource
def _vigilante():
    """Una sola vez por portal: se apaga solo si nadie lo usa (ver portal/apagado.py)."""
    vigilar_inactividad()
    return True


_vigilante()
st.sidebar.caption(f"Conciliación de Honorarios · versión {__version__}")
if st.sidebar.button("Cerrar el portal", help="Detiene todo lo del portal (como Ctrl + C) y lo cierra. "
                                            "Para volver a abrirlo: acceso directo Portal de Honorarios"):
    st.html(pagina_portal_cerrado(), unsafe_allow_javascript=True)
    cerrar_portal()
    st.stop()
navegacion.run()
