"""Detalles visuales del portal: logo de la empresa y aviso que bloquea la pantalla mientras se procesa."""

from pathlib import Path

RECURSOS = Path(__file__).parent / "recursos"


def ruta_logo():
    """Logo de la empresa si existe portal/recursos/logo.png (o .jpg/.jpeg/.svg); si no, None."""
    for extension in (".png", ".jpg", ".jpeg", ".svg"):
        ruta = RECURSOS / f"logo{extension}"
        if ruta.exists():
            return ruta
    return None


# Capa semitransparente sobre TODA la página (menú incluido): mientras se procesa, un clic en otra página o en
# cualquier control reiniciaría la página y detendría la conciliación. Se quita sola al terminar.
BLOQUEO = """
<style>
.bloqueo-portal {
    position: fixed; inset: 0; z-index: 1000000;
    background: rgba(255, 255, 255, 0.45);
    display: flex; align-items: flex-end; justify-content: center; padding-bottom: 6vh;
}
.bloqueo-portal > div {
    background: #ffffff; border: 1px solid #c9d3e3; border-left: 6px solid #305496; border-radius: 10px;
    box-shadow: 0 6px 24px rgba(0, 0, 0, 0.15); padding: 16px 22px; max-width: 560px; font-size: 1rem;
}
</style>
<div class="bloqueo-portal"><div>
⏳ <b>Conciliación en proceso.</b><br>
Espera a que termine antes de cambiar de página o cerrar esta pestaña: si sales ahora, el proceso se detiene
y se pierde el avance. Puedes seguir usando otros programas (excepto SAP).
</div></div>
"""
