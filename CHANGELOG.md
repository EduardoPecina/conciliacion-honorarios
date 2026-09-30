# Historial de cambios

## [1.0.0] - 2026-09-30

Primera versión en control de versiones.

### Estructura
- Código dividido en módulos dentro de `honorarios/` (antes, un solo archivo de ~1,900 líneas).
  El comportamiento es idéntico a la versión anterior: se verificó comparando el Excel de salida
  celda por celda sobre los mismos datos de prueba.
- `conciliar.py` queda como punto de entrada; se ejecuta igual que antes.
- `EJECUTAR.bat` para ejecutar con doble clic.
- Pruebas automáticas con datos inventados en `tests/`.

### Funcionalidad incluida
- Cruce TXT → catálogo → XML, uno a uno y pagos quebrados (N líneas contra M facturas).
- Reglas de receptor (RFC, nombre, régimen 601, uso G03), periodo, método (PUE/PPD) y forma de pago (03/99).
- Nota informativa en facturas PPD: el médico debe enviar el complemento de pago.
- Recálculo de impuestos por tipo (E0/F2) y régimen (612/626).
- Historial de UUID para evitar pagos dobles entre meses.
- Validación opcional contra SAP: todo F2 debe tener IH; régimen del XML contra RS/RH.
- Reporte en Excel con Resumen dinámico (fórmulas), Detalle, Cálculos, Por proveedor, XML sin TXT,
  Errores XML, Tasas e Info.
