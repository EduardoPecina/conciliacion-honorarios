"""Portal de Honorarios (Streamlit). Uso, desde la carpeta del repositorio:

    python -m streamlit run portal/app.py        (o doble clic en EJECUTAR_PORTAL.bat)

REGLA: el portal NO tiene reglas de negocio. Solo recibe archivos, llama al motor (honorarios/) y muestra
el resultado. Cualquier regla nueva va en honorarios/, así el portal y el .bat siempre dan lo mismo."""

import sys
from pathlib import Path

import streamlit as st

RAIZ_REPO = Path(__file__).resolve().parents[1]
if str(RAIZ_REPO) not in sys.path:
    sys.path.insert(0, str(RAIZ_REPO))           # para importar honorarios/ y portal/

from honorarios import __version__  # noqa: E402

st.set_page_config(page_title="Portal de Honorarios", page_icon="🩺", layout="wide")

paginas = Path(__file__).parent / "paginas"
navegacion = st.navigation([
    st.Page(paginas / "conciliar.py", title="Conciliar", icon="✅", default=True),
    st.Page(paginas / "catalogo.py", title="Catálogo de proveedores", icon="📇"),
    st.Page(paginas / "historial.py", title="Historial de UUID", icon="🗂️"),
])
st.sidebar.caption(f"Conciliación de Honorarios · versión {__version__}")
navegacion.run()
