"""
Conciliación de Honorarios Médicos.

Cruza el TXT de pagos (control interno) contra el catálogo de proveedores y los XML (CFDI)
de las facturas, aplica las reglas de negocio y genera un Excel con el estatus de cada línea.

Regla de oro: el Estatus "Cruzado" solo se asigna si la línea pasó TODAS las validaciones
de config.VALIDACIONES. Si alguna no se pudo ejecutar (historial ilegible, SAP pedido pero
sin respuesta...), la línea sale "Cruzado con observaciones" y dice cuál faltó.
"""

__version__ = "1.10.1"
