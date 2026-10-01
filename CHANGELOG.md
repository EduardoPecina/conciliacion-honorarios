# Historial de cambios

## [1.12.0] - 2026-10-01

### Cambiado
- **Reusar las consultas a SAP ahora viene desmarcado por defecto** (portal y `.bat`). Volver a correr un mes casi
  siempre es porque se corrigió algo, muchas veces en SAP por una línea Detenida; reusar daría lo que SAP tenía
  antes de la corrección. Consultar todo de nuevo tarda unos minutos más, pero es lo seguro.
- Texto más claro: *"Reusar lo consultado a SAP el 15/09/2026 a las 10:42 (más rápido)"*, con la explicación en
  el ⓘ (cuándo marcarlo y cuándo no, con ejemplos: se agregó IH, RS↔RH).

### Corregido
- `CodigosSAP_AAAA-MM.xlsx`: al reusar, cada proveedor conserva la fecha en que realmente se consultó (antes se
  sobrescribía con la fecha de la corrida, como si se hubiera consultado de nuevo).

## [1.11.4] - 2026-10-01

### Cambiado
- Se regresa a la **licencia de la empresa** (uso interno), que sustituye a la MIT de la 1.11.3: el programa
  tiene reglas de negocio específicas de la empresa y así se evitan ambigüedades. Se agrega el crédito del autor
  (Eduardo Pecina) en `LICENSE` y en el README.

## [1.11.3] - 2026-10-01

### Cambiado
- **Licencia MIT** (antes, licencia propietaria de uso interno), con autorización del apoderado legal de la
  empresa. El repositorio sigue siendo privado. El logo de la empresa no se incluye en la licencia.

## [1.11.2] - 2026-10-01

### Cambiado
- "Quién actúa" de las líneas **Detenida** (tabla "Qué hacer" del portal y Excel): ahora dice solo
  **Customer Services** (antes "Sistemas / Customer Services", que hacía pensar que era una falla del programa).

## [1.11.1] - 2026-10-01

### Corregido
- **Doble clic varias veces al abrir el portal ("jitter click").** Antes, cada clic hecho mientras el portal
  arrancaba lanzaba su propio servidor: en una prueba con 5 clics llegaron a correr 3 a la vez (los de más
  terminaban con "Port 8501 is not available"), se abrían 5 pestañas, `portal.log` quedaba revuelto y, con mala
  suerte, salía una ventana de error falsa. Ahora solo un lanzador a la vez trabaja (candado en la carpeta temporal,
  que el sistema suelta solo aunque el lanzador se caiga); los clics de más se retiran sin hacer nada y se abre una
  sola pestaña. Misma prueba: 1 servidor, 1 pestaña, 0 errores.
- Si el servidor que arrancó el lanzador termina porque otro portal ya tiene el puerto, ya no se reporta como error.
- Un lanzador nunca vive más de 3.5 minutos (suelta el candado aunque algo se atore) y suelta el candado antes de
  mostrar una ventana de error, para que otro clic pueda reintentar.

### Agregado
- **Portal trabado:** si el portal ocupa el puerto pero no contesta (4 intentos de 5 s), el doble clic detiene sus
  procesos (solo los de este portal) y arranca uno nuevo. Probado: un portal ocupado conciliando contesta en
  centésimas de segundo, así que no se reinicia a la mitad.

### Documentación
- README: qué es `CodigosSAP_AAAA-MM.xlsx` (consultas a SAP guardadas por mes), cuándo reusarlas y cuándo no.

## [1.11.0] - 2026-10-01

### Cambiado
- **"Cerrar el portal" detiene todo lo del portal**, como un Ctrl + C: este portal, portales viejos que se hayan
  quedado abiertos y el lanzador (antes solo apagaba el portal en uso y podían quedar procesos de Python
  corriendo). Reconoce los procesos por su línea de comando (Python corriendo `portal/app.py` o
  `portal/iniciar.py` de esta instalación): otros Python, como el de VS Code, no se tocan. Usa `psutil`
  (agregado a `requirements-portal.txt`; el instalador lo instala).
- Al cerrar, la pestaña muestra "El portal se cerró" en lugar del "Connection error… restart it in your terminal"
  de Streamlit.

### Documentación
- README: "Publicar una corrección" (pasos para mandar una versión nueva a los usuarios).

## [1.10.1] - 2026-10-01

### Retirado
- Se quita el desinstalador (`DESINSTALAR.bat`, `instalador/desinstalar.py`) agregado en la 1.10.0: para
  actualizar o corregir un bug basta con correr el `INSTALAR.bat` de la versión nueva, que reemplaza `_Programa`
  completo, conserva catálogo, historial y reportes, y si algo falla deja la versión anterior.
- Al actualizar, el instalador borra el acceso directo "Desinstalar Portal de Honorarios" que dejó la 1.10.0.

## [1.10.0] - 2026-09-30

### Agregado
- **Desinstalador** (`DESINSTALAR.bat` + `instalador/desinstalar.py`) y acceso directo "Desinstalar Portal de
  Honorarios" en la carpeta `Honorarios`:
  - Confirma la desinstalación y pregunta si conservar el catálogo, el historial de UUID y los reportes
    (recomendado); si se eligen borrar, pide una segunda confirmación.
  - Borra solo lo que puso el instalador (`_Programa`, `LEEME.txt`, accesos directos); nunca archivos ajenos.
    No desinstala Python ni las librerías.
  - Se ejecuta desde una copia en la carpeta temporal de Windows para poder borrar `_Programa`. Si el portal
    está abierto o Windows no deja borrar algo, lo avisa.
- `LEEME.txt` explica cómo desinstalar.

## [1.9.0] - 2026-09-30

### Agregado
- **Instalador para las computadoras de Contabilidad** (`INSTALAR.bat` + `instalador/instalar.py`). Sirve para
  instalar y para actualizar:
  - Python 3.10+: si falta, lo instala solo para el usuario, sin permisos de administrador (winget o
    python.org); si ya está, se lo salta.
  - Librerías: `pip install -r requirements-portal.txt` (reintenta ante cortes de la red).
  - Crea `Honorarios` en el Escritorio del usuario (se le pregunta a Windows, así funciona con OneDrive) con
    `_Configuracion`, `Reportes`, `_Programa` y `LEEME.txt`, sin carpetas de meses.
  - Acceso directo "Portal de Honorarios" con el logo de Nova, que abre el portal sin ventana negra.
  - Al actualizar reemplaza solo `_Programa`; el catálogo, el historial y los reportes no se tocan. Si el
    portal está abierto, pide cerrarlo.

### Cambiado
- El portal guarda los Excel en `Honorarios/Reportes/<Año>/<MM Mes>/` (antes, en `<Año>/<MM Mes>/Salida`, mezclado
  con las carpetas de entrada del programa de consola).

## [1.8.1] - 2026-09-30

### Cambiado
- Portal: el aviso mientras se procesa ya no bloquea la página (no dejaba hacer scroll para ver las barras de
  avance). Ahora es un aviso informativo; solo si se da clic en otra página del menú o en "Cerrar el portal", el
  navegador pregunta "¿Quieres salir de todas formas?" (Aceptar sale, Cancelar se queda y el proceso sigue). Al
  cerrar o recargar la pestaña, el navegador muestra su aviso de "¿Salir del sitio?".

## [1.8.0] - 2026-09-30

### Cambiado
- `EJECUTAR_PORTAL.bat` ya no muestra la ventana negra: `portal/iniciar.py` (con `pythonw`) arranca el portal en
  segundo plano y abre el navegador predeterminado de Windows (pestaña nueva, o lo abre si estaba cerrado). Antes
  el navegador podía no abrirse porque el `.bat` dependía de PowerShell, que las políticas de la empresa pueden
  limitar.
- Doble clic con el portal ya abierto: solo abre otra pestaña (no arranca un segundo portal).

### Agregado
- Botón **"Cerrar el portal"** en el menú, y apagado automático a los 15 minutos sin pestañas abiertas
  (`portal/apagado.py`).
- Si el portal no puede iniciar, una ventana de Windows muestra el motivo; el detalle queda en `portal/portal.log`.

## [1.7.0] - 2026-09-30

### Agregado
- No se concilian periodos del futuro (portal y `.bat`). La fecha de "hoy" se toma de internet (encabezado `Date`
  del servicio del SAT); sin conexión, del reloj de la PC. El portal solo ofrece el mes actual y los anteriores
  (y el año anterior completo).
- Portal: mientras se procesa, la página queda bloqueada con un aviso para no detener la conciliación al cambiar
  de página o de control. Se libera sola al terminar.
- Portal: logo de Nova (`portal/recursos/logo.png`), arriba del título "Conciliación de Honorarios" y en la
  parte superior del menú (recortado al recuadro naranja, sin el margen claro que se notaba en el menú gris).
  Para cambiarlo basta con reemplazar ese archivo.

### Cambiado
- Portal: "Conciliar" pasa a "Conciliador Automático"; "Catálogo de proveedores" a "Catálogo de Proveedores
  TM26"; título "Conciliación de Honorarios".

## [1.6.1] - 2026-09-30

### Cambiado
- Instrucciones de SAP más claras y menos alarmantes (portal y `.bat`): lista de requisitos (SAP abierto en
  productivo, sesión iniciada, acceso a FK03, scripting de SAP GUI habilitado) y aviso de que se usa la primera
  ventana de SAP. Solo hay que evitar usar SAP mientras se consulta; el resto de la computadora sí se puede usar.

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
