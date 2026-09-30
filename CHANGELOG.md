# Historial de cambios

## [1.6.0] - 2026-09-30

### Cambiado
- **La validación contra SAP es obligatoria** (a petición de Contabilidad). Antes era opcional (el `.bat`
  preguntaba "¿Validar contra SAP?" con "no" por defecto). Si SAP no se pudo consultar, ninguna línea queda
  "Cruzado": sale "Cruzado con observaciones" con "Validaciones no realizadas: códigos de retención en SAP".
- El `.bat` ya no pregunta si validar SAP: pide abrir SAP e iniciar sesión y espera Enter.
- El portal también consulta SAP con la sesión de SAP GUI de la computadora donde corre (aviso antes de
  procesar, barra de avance, opción de reusar las consultas guardadas del mes, métricas de SAP).
- El historial de UUID **no se registra** si en la corrida no se pudo consultar SAP o el SAT (en el portal,
  el botón "Confirmar" queda bloqueado).

## [1.5.1] - 2026-09-30

### Corregido
- `EJECUTAR_PORTAL.bat` abría el navegador antes de que el portal estuviera listo ("localhost rechazó la
  conexión"). Ahora espera a que el portal escuche en el puerto 8501. Si el portal se detiene, la ventana
  queda abierta con el error en lugar de cerrarse.
- El portal usa siempre el puerto 8501: si ya hay otro portal abierto, avisa en lugar de irse en silencio al 8502.

### Seguridad
- El portal solo acepta conexiones desde la misma computadora (`server.address = "localhost"`). Antes, mientras
  corría en una PC, cualquiera en la red podía abrirlo con la "Network URL".

## [1.5.0] - 2026-09-30

### Agregado
- **Portal web (en prueba)** con Streamlit (`portal/`, `EJECUTAR_PORTAL.bat`, `requirements-portal.txt`):
  subir TXT y XML (o `.zip`), Procesar, resumen por estatus y del SAT, lista "Qué hacer", Descargar Excel y
  "Confirmar y registrar en el historial". Páginas de consulta del catálogo y del historial.
- `honorarios/servicio.py`: la conciliación completa sin preguntas en pantalla. La usan el `.bat` y el portal,
  así que siempre dan el mismo resultado.

### Cambiado
- El `.bat` ahora usa `servicio.py` por dentro: mismas preguntas y mismo resultado.
- `entradas.rutas_en()`: la revisión de carpetas funciona con cualquier carpeta, no solo `<Año>/<Mes>`.
- La consulta al SAT reporta su avance con funciones (barra de progreso en el portal, texto en la consola).

## [1.4.1] - 2026-09-30

### Cambiado
- Preguntas más claras: en lugar de `Año [2026]:` dice `Año (Enter = 2026, año actual):`, el mes muestra su
  nombre (`Enter = 9, septiembre`) y las preguntas sí/no dicen `(s = sí, n = no; Enter = no)`.
- Al capturar el periodo se confirma: "Periodo a conciliar: septiembre 2026".
- Un año que no tenga 4 dígitos (p. ej. "26") se rechaza con un mensaje claro.

## [1.4.0] - 2026-09-30

### Cambiado
- Tolerancias ajustadas a petición de Contabilidad: `TOLERANCIA` (importe TXT vs factura) y
  `TOLERANCIA_IMPUESTOS` (impuestos, total y conceptos del XML) bajan de 1.00 a **0.50**.
  `UMBRAL_OK` sigue en 0.05: la etiqueta "Redondeo" cubre de 0.05 a 0.50.
- Los textos del Excel que decían "1.00" ahora toman el valor de la configuración.

## [1.3.0] - 2026-09-30

### Agregado
- Alerta EFOS (artículo 69-B del CFF): con la misma consulta al SAT se revisa si el emisor podría estar en la
  lista de operaciones simuladas. Si es así, la línea queda "Cruzado con observaciones" con "Alerta EFOS".
  Nueva columna `EFOS` en Detalle y en XML sin TXT, indicador en el Resumen y conteo en la hoja Info.
  Códigos del campo `ValidacionEFOS` según la documentación del SAT (v1.4): 100 = el emisor está en la lista
  (alerta); 200 = no está; 201 = ni el emisor ni los terceros están. Un código no documentado también es alerta.

### Cambiado
- Se valida el código postal del receptor: `CP_SOCIEDAD = "66450"`. Una factura con otro
  `DomicilioFiscalReceptor` se rechaza.

## [1.2.0] - 2026-09-30

### Agregado
- Regla obligatoria: en cada corrida se verifica el estatus de cada factura en el SAT (Vigente / Cancelado /
  No encontrado) con el servicio web oficial `ConsultaCFDIService`, que no tiene captcha. No se pregunta.
  - Cancelada: "Rechazar", con la observación "Verificar si el médico envió la factura que sustituye a esta,
    que está cancelada en el SAT".
  - Si el médico reemitió por el mismo importe, se cruza la nueva ("Cruzado", sin marcarla como facturación
    doble) y la cancelada queda en XML sin TXT indicando cuál la sustituyó.
  - Nuevo indicador en el Resumen: "Líneas con factura cancelada en el SAT".
  - No encontrada, o con cancelación en proceso: "Detenida".
  - Sin respuesta del SAT: ninguna línea queda "Cruzado"; salen "Cruzado con observaciones" con
    "Validaciones no realizadas: estatus del CFDI en el SAT".
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
