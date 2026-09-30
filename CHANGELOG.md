# Historial de cambios

## [1.2.0] - 2026-09-30

### Agregado
- Verificación opcional del estatus de cada factura en el SAT (Vigente / Cancelado / No encontrado) con el
  servicio web oficial `ConsultaCFDIService`, que no tiene captcha. El programa pregunta al iniciar
  (por defecto "s", ver `ESTADO_SAT_POR_DEFECTO`).
  - Cancelada: "Rechazar", con la observación "Verificar si el médico envió la factura que sustituye a esta,
    que está cancelada en el SAT".
  - Si el médico reemitió por el mismo importe, se cruza la nueva ("Cruzado", sin marcarla como facturación
    doble) y la cancelada queda en XML sin TXT indicando cuál la sustituyó.
  - Nuevo indicador en el Resumen: "Líneas con factura cancelada en el SAT".
  - No encontrada, o con cancelación en proceso: "Detenida".
  - Sin respuesta del SAT: "Cruzado con observaciones" con "Validaciones no realizadas".
  - Nueva columna `EstadoSAT` en Detalle y en XML sin TXT; resumen de la consulta en la hoja Info.

## [1.1.1] - 2026-09-30

### Corregido
- Si el TXT no traía encabezado y la primera línea tenía la fecha en otro formato (p. ej. "1.09.2026"),
  esa línea se tomaba por encabezado y su pago desaparecía del reporte sin avisar. Ahora solo se
  considera encabezado un renglón cuya primera columna no trae dígitos.

## [1.1.0] - 2026-09-30

### Cambiado
- Se retira la regla "Proveedor bloqueado en el catálogo" (a petición de Contabilidad): la columna
  `Activo` del catálogo es histórica y trae registros que no reflejan el estado real del proveedor.
  El programa ya no la lee: no detiene líneas por ella ni rechaza el catálogo por valores raros en esa columna.

## [1.0.1] - 2026-09-30

### Corregido
- Un texto que empieza con "=" (nombre del emisor en un XML, folio del TXT, nombre de archivo...)
  se escribía en el Excel como fórmula y Excel lo ejecutaba. Ahora siempre se escribe como texto.
- El encabezado del TXT no se reconocía si el archivo empezaba con renglones vacíos.

### Documentación
- README: estatus "Proveedor extranjero" e instrucciones para clonar el repositorio como `_Programa`.
- Licencia propietaria (`LICENSE`).

## [1.0.0] - 2026-09-30

Primera versión en control de versiones.

### Estructura
- Código dividido en módulos dentro de `honorarios/` (antes, un solo archivo de ~1,900 líneas).
  El comportamiento es idéntico a la versión anterior: se verificó comparando el Excel de salida
  celda por celda sobre los mismos datos de prueba.
- `_Programa.py` es el punto de entrada (nombre pedido por Sistemas).
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
