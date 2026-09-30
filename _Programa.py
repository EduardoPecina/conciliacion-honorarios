"""
Punto de entrada del programa. Uso:

    python _Programa.py

o doble clic en EJECUTAR.bat. Debe estar dentro de la carpeta _Programa de la estructura:

    Honorarios/
    ├── _Programa/          <- este repositorio
    ├── _Configuracion/     <- catálogo de proveedores e historial de UUID
    └── <Año>/<MM Mes>/Entrada y Salida
"""

from honorarios.principal import main
from honorarios.utilidades import ErrorUsuario

if __name__ == "__main__":
    try:
        main()
    except ErrorUsuario as e:
        print(f"\n*** NO SE PUDO COMPLETAR ***\n{e}")
    except Exception as e:
        print(f"\n*** ERROR INESPERADO ***\n{type(e).__name__}: {e}")
        print("Toma captura de este mensaje y compártela con quien mantiene el programa.")
    input("\nPresiona Enter para cerrar...")
