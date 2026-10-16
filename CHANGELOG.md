# Historial de cambios

## [1.21.0] - 2026-10-16

### Agregado
- **Facturas pendientes (segunda vuelta del mes)**, en Conciliador Automático: se suben solo los XML nuevos
  (refacturas o XML que no habían llegado) y se buscan **solo entre las líneas que no cruzaron**. Lo cruzado no se
  toca: ni sus facturas ni sus copias pueden cruzar otra línea. Solo se consultan al SAT los XML nuevos; SAP se
  reusa. El cruce prefiere la factura válida a la rechazada. Un XML que ya estaba en el mes se ignora. El portal
  muestra qué líneas cambiaron, actualiza el Excel (la hoja Info registra cada vuelta) y guarda el trabajo.
  - Una factura agregada como pendiente puede venir fechada hasta el **mes siguiente** (un CFDI no se puede fechar
    hacia atrás) y queda **Cruzado con observaciones**. En la primera vuelta la regla del periodo no cambia.
- **"TXT generado"** en el Excel (hojas TXT para SAP y Detalle): fecha y Excel con que cada línea ya salió en un TXT
  de Generar TXT. En TXT para SAP, dentro de cada estatus va primero lo que todavía no tiene TXT. El Excel se vuelve
  a generar solo cuando se descarga un TXT nuevo.

### Cambiado
- **Generar TXT no deja descargar líneas que ya salieron en un TXT anterior** (posible pago doble), salvo que se
  marque que la carga anterior a SAP falló. Antes solo se avisaba.

## [1.20.1] - 2026-10-16

### Corregido
- **La Referencia medía 25 caracteres** (`26185113K7Q3MXP4_Q18-2026`) y SAP acepta 16. Ahora mide 16 como máximo: los
  últimos 7 caracteres del folio nuevo (los que lo distinguen) y la quincena: folio `26185113K7Q3MXP4` → Referencia
  `7Q3MXP4_Q18-2026`. El folio sigue siendo de 16 caracteres y no cambia.
- Candado: antes de descargar, ningún Folio ni Referencia puede pasar de 16 caracteres (`LONGITUD_FOLIO` y
  `LONGITUD_REFERENCIA` en `honorarios/config.py`). Si SAP confirma que la Referencia acepta más, se sube
  `LONGITUD_REFERENCIA` y la Referencia vuelve a llevar el folio completo.

## [1.20.0] - 2026-10-16

### Cambiado
- **Generar TXT: la Referencia ya no se captura, la arma el portal.** El Excel de la conciliación no la mostraba, así
  que al armar el Excel con la plantilla quedaba vacía y la Referencia salía solo con el folio. Ahora "Formatear
  folios" la arma siempre: el folio nuevo y la quincena de la fecha, `26185113K7Q3MXP4_Q18-2026`, con la misma
  estructura que la del TXT original (`A1893      Q18-2026`).
  - El Excel lleva solo las columnas **A a J** (Fecha … PM). La plantilla ya no trae Referencia; si un Excel la trae
    en la K, no se lee. En el TXT sigue saliendo la Referencia como columna 11, porque SAP espera las 11 columnas.
  - Hoja "TXT para SAP" del Excel de la conciliación: columnas A a J y, de apoyo, Estatus, Médico y Línea (K a M).
  - La vista previa muestra la quincena y avisa si las fechas del Excel son de más de una.
  - El folio también toma la quincena de la fecha. Para un mismo Excel, los folios son los mismos que antes.

## [1.19.1] - 2026-10-16

### Cambiado
- **INSTALAR.bat ya funciona aunque se abra directo dentro del ZIP**, sin "Extraer todo". Antes el instalador se
  detenía y pedía extraerlo. Ahora copia el programa a `%LOCALAPPDATA%\Honorarios\Instalador` (fuera de la
  carpeta temporal, que Windows puede borrar, y fuera de OneDrive), revisa que la copia quedó completa y sigue
  desde ahí, en la misma ventana. `instalacion.log` lo anota. Solo si el ZIP no se extrajo completo pide extraerlo
  a mano.

## [1.19.0] - 2026-10-16

### Cambiado
- **Edición TXT se reemplaza por Generar TXT: el TXT para SAP ya no se edita en el portal.** El TXT es lo que va a SAP
  y genera los pagos, así que los cambios manuales se quitan. Ahora el analista arma un Excel solo con las líneas
  listas para pagar, con las 11 columnas del TXT en su orden (A a K). Lo sube en **Generar TXT**, revisa la vista
  previa (con el médico, el total de líneas y el monto total), da clic en **Formatear folios** y descarga el TXT. Lo
  que sigue en revisión se queda fuera y se paga después, sin detener a lo demás.
  - **El portal revisa, no corrige:** cada celda que SAP rechazaría se reporta con su fila y columna (fecha, monto,
    Tax E0/F2, X/1/PM, campos vacíos, cuenta convertida en número por Excel, tabuladores, caracteres raros). No se
    genera el TXT hasta corregirla en el Excel.
  - **Avisos sin bloquear:** IDProveedor que no está en el catálogo, y dos filas con la misma factura (mismo ID,
    folio, monto y fecha: ¿pago doble?).
  - **Mismo formato que el TXT original:** tabulador, sin encabezado, CRLF, ANSI. Lo que Excel cambia se regresa a
    como va en el TXT: fecha `29.09.2026`, monto `9096.00`, columna 1 `001`.
  - Al descargar se guardan una copia del TXT y el **control de folios** en la carpeta del mes. El control (de qué
    folio original salió cada folio nuevo) también se puede descargar.
  - Plantilla vacía descargable, con las columnas en formato texto.
- **Folios de 16 caracteres, rastreables:** año + quincena + últimos 4 del IDProveedor + 8 caracteres calculados con
  la factura (`26185113K7Q3MXP4`). Hay 1.1 × 10^12 combinaciones por proveedor y quincena, y se comprueban contra el
  archivo y contra `HistorialFolios.csv`. Misma factura → mismo folio: si una línea ya cargada se vuelve a mandar, SAP
  la rechaza (no se paga dos veces) y el portal avisa que ese folio ya se había generado. `HistorialFolios.csv`
  registra además el folio original, el monto, la fecha, el Excel, el usuario y la fecha de generación.
- La Referencia sigue al folio nuevo: `A1893      Q18-2026` → `26185113K7Q3MXP4_Q18-2026`.

### Agregado
- Excel de la conciliación: hoja **TXT para SAP**, con las líneas del TXT tal cual (A a K, como texto) y, de apoyo,
  Estatus, Médico y Línea. Se deja solo lo aprobado y se sube en Generar TXT. Conciliador Automático tiene un enlace
  directo.

### Quitado
- Tabla editable, ventana "Editar línea", sugerencias de ID, filtros y bitácora de Edición TXT; la hoja "Cambios TXT"
  del Excel.

## [1.18.1] - 2026-10-06

### Corregido
- **Una actualización dejó `_Programa` vacío y el acceso directo sin funcionar**, y aun así el instalador dijo
  "terminó bien". Ahora el instalador **revisa la copia**: compara archivo por archivo `_Programa` contra el ZIP.
  - Si falta algo, **regresa la versión anterior** (el respaldo `_Programa_anterior` solo se borra cuando la versión
    nueva ya está revisada) y lo avisa.
  - Al final vuelve a revisar unos segundos después (por si OneDrive tocó la carpeta) y, si faltan archivos, pide
    volver a correr INSTALAR.bat en lugar de decir que terminó bien.
  - `instalacion.log` registra cuántos archivos se copiaron y el resultado de cada revisión.
- **INSTALAR.bat abierto directo dentro del ZIP** (sin "Extraer todo"): Windows lo corre desde una carpeta temporal
  (`...\Temp\...zip.123\...`). Ahora el instalador lo detecta y pide extraer el ZIP primero.

## [1.18.0] - 2026-10-06

### Seguridad
- **Librerías con huella SHA-256:** nuevo `requirements-instalar.txt` (generado con `instalador/generar_lock.py` desde
  `requirements-lock.txt`) con la huella de cada instalador publicado en PyPI. El instalador usa `--require-hashes`:
  si un archivo descargado no coincide byte por byte, la instalación se detiene. Verificado con los 42 instaladores
  reales de Windows + Python 3.13, y comprobado que una huella alterada detiene la instalación.
- **Trabajo guardado firmado (HMAC-SHA256):** "Continuar donde te quedaste" solo abre archivos firmados en esa misma
  computadora, con una clave de 256 bits guardada fuera de OneDrive (`%LOCALAPPDATA%\Honorarios`). La firma se
  comprueba antes de leer el contenido: un archivo modificado, sin firma o de otra computadora no se abre. Los
  trabajos guardados con versiones anteriores (sin firma) no se pueden continuar: hay que volver a procesar el mes.
- README, "Datos y seguridad": resumen de las conexiones (todo por HTTPS verificado; el portal solo en localhost) y
  de las protecciones.

## [1.17.1] - 2026-10-05

### Corregido
- Edición TXT: **a veces un filtro no se podía quitar o "revivía"**. Dos causas:
  - La lista de estatus del panel solo ofrecía los estatus presentes en la tabla: "Pendientes de revisar" en un mes
    sin líneas Cruzado dejaba un filtro activo pero invisible en el panel. Ahora ofrece siempre todos los estatus,
    con su número de líneas.
  - Si el panel se había abierto, sus controles conservaban los valores de ese momento y, en la siguiente acción,
    volvían a poner filtros ya quitados con los botones. Ahora los filtros se guardan aparte y el panel se vuelve a
    crear con los valores actuales cada vez que un botón los cambia. Además se conservan al cambiar de página.

### Agregado
- Barra de filtros activos arriba de la tabla: cada filtro es un botón con ✕ para quitar solo ese, y **Quitar todos
  los filtros** siempre visible.

## [1.17.0] - 2026-10-03

### Agregado
- **Bitácora de instalación `Honorarios\instalacion.log`:** todo lo que muestra la ventana del instalador (pasos de
  `INSTALAR.bat`, salida de pip, mensajes, errores inesperados con su detalle) queda en un archivo, con un encabezado
  por instalación (fecha, versión, Python, Windows, usuario) y el resultado final. Los mensajes de error piden mandar
  ese archivo. Se agrega cada instalación (historial) y se recorta solo si pasa de 2 MB.
- `INSTALAR.bat` anota sus pasos (búsqueda e instalación de Python, códigos de winget/curl); si no logra instalar
  Python, abre la carpeta con su registro.

## [1.16.1] - 2026-10-03

### Corregido
- **`_Programa_anterior` ya no se queda a la vista** después de actualizar. El instalador reintenta borrarlo unos
  segundos (OneDrive o el antivirus pueden tener archivos abiertos) y quita el atributo de solo lectura; si aun así
  no puede, el portal lo borra en segundo plano la próxima vez que se abre (solo si es una copia del programa).
- Si un respaldo viejo seguía bloqueado, la siguiente actualización fallaba al renombrar `_Programa`; ahora usa otro
  nombre (`_Programa_anterior_2`) y lo limpia después.

## [1.16.0] - 2026-10-03

### Agregado
- Edición TXT: **filtros y orden** para la tabla.
  - Botones rápidos **Pendientes de revisar** (oculta lo Cruzado, ordena por gravedad) y **Para corregir** (problemas en
    el TXT o posible ID equivocado).
  - Botón **Filtros y orden** (embudo): estatus en modo Mostrar solo / Ocultar, líneas con problemas / posible ID
    equivocado / editadas / folio sin formatear, Tax, rango de monto y orden (archivo, gravedad, médico, monto,
    problemas primero).
  - Leyenda con los filtros activos ("Mostrando X de N líneas · …") y Quitar filtros. La descarga aclara que el TXT
    lleva todas las líneas, sin importar los filtros.

## [1.15.3] - 2026-10-03

### Cambiado
- Edición TXT: **se quita el botón "Aplicar"** de las sugerencias de ID (cambiaba el ID en un clic). Ahora son
  "Posibles ID equivocados · solo apoyo", con el recordatorio de verificar en SAP, y el botón **Editar línea** abre una
  ventana con los valores actuales de IDProveedor, Nombre (se busca en el catálogo al escribir), Tax y Monto. Quien
  edita decide; el cambio queda en la bitácora.

## [1.15.2] - 2026-10-03

### Cambiado
- Formatear folios: la Referencia queda **sin espacios**, con guion bajo entre el folio y el resto
  (`A1893      Q18-2026` → `7KQ2M9XHRT_Q18-2026`). Con el folio de 10 caracteres mide lo mismo que antes (19).

## [1.15.1] - 2026-10-03

### Cambiado
- Edición TXT: **la Referencia cambia junto con el folio** al presionar "Formatear folios" (solo la parte del folio:
  `A1893      Q18-2026` → `7KQ2M9XHRT      Q18-2026`; espacios y resto intactos) y ya no se edita a mano.
- La columna **Tax** se edita con una lista (E0 / F2), para no teclear códigos inválidos.

## [1.15.0] - 2026-10-03

### Agregado
- **Trabajo de varios días:** todo lo del mes (resultado, consultas al SAT y a SAP, TXT original y cambios de
  Edición TXT) se guarda solo en la carpeta del mes. Al abrir el portal, **"Continuar donde te quedaste"** lo
  recupera con un clic, sin volver a consultar nada; el selector se mueve al mes recuperado.
- **Actualizar SAT / Actualizar SAP** en el resultado, con la fecha de la última consulta; si tiene 1 día o más, un
  aviso pide actualizar antes de confirmar o descargar (una factura pudo cancelarse o SAP pudo corregirse).
- **Excel con las ediciones:** "Descargar Excel" genera el reporte actual con la hoja **Cambios TXT** (bitácora); el
  TXT editado se guarda junto al Excel y su huella queda en la hoja Info. Enlace desde Edición TXT.

## [1.14.1] - 2026-10-03

### Cambiado
- **Versiones fijas, para que el programa funcione igual por años en cualquier PC:**
  - `INSTALAR.bat` usa siempre **Python 3.13** (si no está, lo instala para el usuario aunque haya otro Python); los
    `.bat` y los accesos directos usan ese mismo Python. El instalador se niega a correr con otra versión.
  - Nuevo `requirements-lock.txt` con las versiones **exactas** de las 40 librerías que se instalan (incluidas las
    dependencias de las dependencias), verificadas para Windows de 64 bits + Python 3.13. El instalador las usa como
    restricción y solo acepta instaladores precompilados (nunca compila).
  - Prueba que falla si el entorno tiene otras versiones o si se pide una librería sin fijar.
- Se reemplazan las funciones de Streamlit marcadas para retirarse: `st.components.v1.html` → `st.html` (con
  JavaScript) y `use_container_width` → `width="stretch"`. Ya no aparecen avisos de funciones obsoletas.

### Corregido
- La página "El portal se cerró" se armaba con etiquetas HTML dentro del script; `st.html` borra en silencio esos
  scripts, así que se reescribió sin etiquetas literales (y una prueba lo vigila).

## [1.14.0] - 2026-10-03

### Agregado
- **Edición TXT** (nueva página): el TXT procesado como tabla editable, con columnas de apoyo que no van al archivo
  (Estatus del cruce, Médico, Problemas del TXT, Sugerencia, Motivo).
  - Se editan IDProveedor, Monto, Tax, Centro de Costo, Cuenta Contable y Referencia. Fecha, X, 1 y PM no; el
    Folio solo con "Formatear folios".
  - Al editar, el médico y el estatus se actualizan solos: se vuelve a cruzar con lo ya consultado al SAT y a SAP.
    Botón para consultar en SAP solo los IDs nuevos.
  - Problemas por celda (rojo); los que SAP rechazaría bloquean la descarga. Celdas editadas en amarillo.
  - Sugerencia de ID correcto cuando una línea sin factura coincide en importe con una factura sin línea de otro
    proveedor, con botón Aplicar.
  - Bitácora de cambios (valor anterior/nuevo, usuario, fecha) y borrador que se recupera si se cierra el portal.
  - Descarga con el formato exacto del TXT original (11 columnas, tabulador, sin encabezados, misma codificación y
    saltos de línea); se guarda copia y bitácora en la carpeta del mes.
- **Folios únicos** (`honorarios/folios.py`): huella SHA-256 del UUID de la factura, 10 caracteres de un alfabeto de
  32 símbolos; parte por línea en pagos quebrados; aleatorio criptográfico para líneas sin factura; comprobación
  contra `HistorialFolios.csv`.
- `honorarios/txt_sap.py`: lectura y escritura fiel del TXT y revisión de problemas. `entradas.leer_linea_txt`
  (la misma validación por línea para la conciliación y la Edición TXT). `servicio.revalidar`.

### Corregido
- Al terminar de procesar, el aviso "La conciliación sigue en proceso" se quedaba activo: cambiar de página o cerrar
  el portal preguntaba si salir aunque ya había terminado.

## [1.13.0] - 2026-10-03

### Agregado
- **Aviso de mes no vigente:** si el mes elegido no es el actual, el portal lo avisa y pide confirmar
  ("Sí, quiero conciliar …") antes de habilitar Procesar. Recordatorio fijo: el TXT, los XML y el mes elegido deben
  coincidir.
- **Detección de archivos de otro mes:** si la mayoría de las fechas del TXT y de los XML no son del mes elegido, la
  conciliación se detiene antes de consultar al SAT y a SAP, y dice de qué mes parecen ser (portal y `.bat`). Antes,
  todas las facturas salían "Rechazar" por estar fuera del periodo, lo que hacía pensar que había que pedir
  refacturación.
- **Candado del historial:** confirmar ya no reemplaza un mes ya registrado con una conciliación que tenga menos de la
  mitad de sus UUID (se perdía la protección contra doble pago de ese mes).
- Prueba que simula el cambio de año (2027, 2031): opciones de año y mes dinámicas.

### Cambiado
- El menú del portal solo muestra **Conciliador Automático**. Catálogo de Proveedores e Historial de UUID se quitan
  del menú (Contabilidad no los usa); el motor sigue usando el catálogo y el historial igual.

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
