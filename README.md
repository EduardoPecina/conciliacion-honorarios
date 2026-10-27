# Conciliación de Honorarios Médicos

Herramienta que automatiza la revisión mensual de los pagos de honorarios a médicos y profesionales de la salud. Cruza el archivo de pagos que se carga a SAP (TXT) contra el catálogo de proveedores y las facturas electrónicas (XML CFDI 4.0), aplica las reglas fiscales y de negocio, y genera un reporte en Excel con el estatus de cada pago.

> **Este repositorio contiene solo código.** Ningún dato real (XML, TXT, catálogos, reportes) se sube aquí. Ver [Datos y seguridad](#datos-y-seguridad).

---

## Qué resuelve

Antes, cada factura se revisaba a mano contra el pago correspondiente. El programa hace en segundos lo siguiente:

1. **Cruza** cada línea del TXT con su factura: `IDProveedor → RFC (catálogo) → XML`, incluyendo pagos quebrados (varias líneas contra una factura, o una línea contra varias facturas).
2. **Valida** cada factura: receptor, periodo, método y forma de pago, timbre, impuestos recalculados, historial de UUID ya pagados.
3. **Consulta al SAT** si cada factura sigue vigente o si el médico la canceló después de enviarla (en cada corrida).
4. **Consulta SAP** (transacción FK03/BP, obligatorio) para confirmar que las retenciones configuradas del proveedor coinciden con la factura.
5. **Genera un Excel** donde cada número se puede rastrear hasta su fórmula.

```
TXT de pagos ──► Catálogo (ID → RFC) ──► XML de la factura
     │                                        │
     └──────────── reglas de negocio ◄────────┘
                          │
                 SAT (vigente / cancelada)
                          │
                   SAP (retenciones)
                          │
                   Reporte en Excel
```

## Estructura de carpetas

El programa vive en `_Programa` y ubica todo lo demás a partir de ahí, así que funciona igual en cualquier computadora.

Con el **portal** (lo que usa Contabilidad), la estructura es simple; el `INSTALAR.bat` la crea:

```
Escritorio/
├── Portal de Honorarios            ← acceso directo (abre el portal en el navegador)
└── Honorarios/
    ├── LEEME.txt
    ├── _Configuracion/
    │   ├── CatalogoProveedoresTM26.xlsx   (lo pone el usuario; tabla de Excel llamada "Catalogo")
    │   └── HistorialUUID.csv              (se crea solo)
    ├── Reportes/                          (Excel de cada conciliación: <Año>/<MM Mes>/, se crean solas)
    └── _Programa/                         ← este repositorio
```

El **programa de consola** (`EJECUTAR.bat`) además usa carpetas por periodo con las entradas:

```
Honorarios/
└── 2026/
    └── 09 Sep/
        ├── Entrada/
        │   ├── XML/                   (facturas .xml, pueden ir en subcarpetas)
        │   └── Honorarios.txt         (un solo .txt, el nombre no importa)
        └── Salida/                    (aquí se guarda el reporte)
```

## Instalación en las computadoras de Contabilidad

El portal corre en la computadora de cada persona (necesita su sesión de SAP GUI). Para instalarlo:

1. **Quien mantiene el programa** descarga el ZIP de la rama `main` (GitHub → **Code → Download ZIP**) y lo comparte (correo, Teams, carpeta compartida). El repositorio es privado: Contabilidad no necesita cuenta de GitHub.
2. **El usuario** da doble clic en **`INSTALAR.bat`**. Lo mejor es descomprimir antes el ZIP (clic derecho → **Extraer todo...**), pero si lo abre directo dentro del ZIP también funciona: Windows lo corre desde una carpeta temporal que puede borrar en cualquier momento, así que el instalador primero copia el programa a `%LOCALAPPDATA%\Honorarios\Instalador` (fuera de Temp y de OneDrive), revisa que la copia quedó completa y sigue desde ahí, en la misma ventana. Si el ZIP no se extrajo completo, pide extraerlo a mano.
   - Si no tiene **Python 3.13** (la versión con la que el programa está probado), lo instala **solo para su usuario, sin permisos de administrador** (con `winget` o, si no está o no pudo instalarlo, con el instalador oficial de python.org, que **solo se ejecuta si trae la firma digital válida de la Python Software Foundation**), junto a cualquier otro Python que ya tenga la PC (ese no se toca). Si ya lo tiene, se lo salta.
   - Instala las librerías (pandas, openpyxl, streamlit, pywin32…) **en las versiones exactas probadas, comprobando la huella SHA-256 de cada archivo** (`requirements-instalar.txt`); las que ya estén, se las salta. Ver [Versiones fijas](#versiones-fijas).
   - Crea `Honorarios` en su **Escritorio** (el de OneDrive si Windows lo tiene ahí; la ruta se le pregunta a Windows, no se escribe a mano), con `_Configuracion`, `Reportes`, `_Programa` y `LEEME.txt`.
   - Copia el programa a `_Programa` y **revisa que quedó completo** (archivo por archivo contra el ZIP). Si algo falta, deja la versión anterior y lo avisa.
   - Crea el acceso directo **"Portal de Honorarios"** (con el logo de Nova) en el Escritorio y en `Honorarios`. Abre el portal sin ventana negra.
3. **El usuario copia el Excel del catálogo** (`CatalogoProveedoresTM26.xlsx`), que le pasa quien mantiene el programa, a `Honorarios\_Configuracion`. **Ese catálogo lo mantiene cada usuario**: los médicos nuevos se dan de alta en su Excel.
4. Doble clic en **"Portal de Honorarios"**.

**Si algo falla al instalar:** todo lo que muestra la ventana negra (incluido lo que imprime pip y los pasos de `INSTALAR.bat` para buscar o instalar Python) queda en **`Honorarios\instalacion.log`**, con un encabezado por instalación (fecha, versión del programa, Python, Windows, usuario) y una línea final `RESULTADO: terminó bien` o `RESULTADO: NO se completó`. Los mensajes de error le piden al usuario mandar ese archivo; así se diagnostica a distancia, aunque haya cerrado la ventana. Cada instalación se agrega al final (sirve también de historial de actualizaciones) y el archivo se recorta solo si pasa de 2 MB. Si falla antes de tener Python, el `.bat` abre la carpeta temporal con su propio registro (`InstalacionHonorarios_bat.log`).

**Actualizar (por ejemplo, para corregir un bug):** se comparte el ZIP de la versión nueva y se corre su `INSTALAR.bat` (con el portal cerrado). No hace falta borrar nada antes:
- Reemplaza `_Programa` **completo** (no quedan archivos de la versión anterior); el catálogo, el historial y los reportes no se tocan.
- Si algo falla a la mitad, deja la versión anterior como estaba: el usuario nunca se queda sin programa.
- Mientras copia, la versión anterior queda como respaldo en `_Programa_anterior` y al terminar se borra (reintenta unos segundos, por si OneDrive la está sincronizando). Si aun así no se puede, **el portal la borra solo la próxima vez que se abre**, así que el usuario no ve carpetas de más. El portal solo borra una carpeta `_Programa_anterior…` junto a `_Programa` cuyo contenido sea una copia del programa; nunca otra cosa.

> ⚠️ Cada computadora tiene su propio catálogo e historial de UUID. Si dos personas concilian en computadoras distintas, un pago doble entre ellas no se detecta. Pendiente: una `_Configuracion` compartida (carpeta de red o SharePoint) o la base de datos.

## Publicar una corrección (quien mantiene el programa)

1. Corrige en una rama (en tu PC o máquina virtual) y corre las pruebas: `python -m unittest discover -s tests`. Prueba también el portal con datos reales.
2. Sube la versión en `honorarios/__init__.py` y anota el cambio en `CHANGELOG.md`. Ese número aparece abajo del menú del portal: así se sabe qué versión tiene cada usuario.
3. Pull Request → merge a `main`.
4. Descarga el ZIP de `main` (GitHub → **Code → Download ZIP**) y compártelo.
5. Cada usuario cierra el portal (botón "Cerrar el portal") y corre el `INSTALAR.bat` del ZIP. Al abrir el portal debe ver la versión nueva en el menú.

## Versiones fijas

Para que una PC instalada dentro de años funcione igual que una instalada hoy, **nada se instala "en la versión más nueva"**:

- **Python 3.13.** `INSTALAR.bat` busca Python 3.13 y, si no está, lo instala para el usuario (aunque la PC tenga otro Python más nuevo). Los accesos directos y los `.bat` usan ese Python. El instalador se niega a correr con otra versión.
- **Librerías:** `requirements-lock.txt` fija las versiones **exactas** de todo lo que se instala (42 librerías, incluidas las dependencias de las dependencias). De ahí se genera `requirements-instalar.txt`, que agrega la **huella SHA-256** de cada instalador publicado en PyPI. El instalador usa `pip install -r requirements-instalar.txt --require-hashes --only-binary=:all:`: solo instaladores precompilados de esas versiones (que PyPI conserva para siempre), y si un archivo descargado no coincide byte por byte con su huella (alterado en el camino o sustituido), la instalación se detiene. Todas se verificaron para Windows de 64 bits con Python 3.13.
- **Pruebas:** `tests/test_instalador.py` falla si el entorno de pruebas tiene versiones distintas a las fijadas, o si se pide una librería que no está fijada.

**Actualizar versiones (a propósito, nunca solo).** Python 3.13 recibe correcciones de seguridad hasta octubre de 2029; seguir usándolo después no rompe nada, pero conviene planear el cambio:
1. En un entorno limpio con la versión nueva de Python, instala las librerías nuevas y corre todas las pruebas y el portal con datos reales.
2. Regenera `requirements-lock.txt` con esas versiones (comprobando que existan para Windows: `pip download --platform win_amd64 --python-version 3.X --only-binary=:all: -r requirements-lock.txt`) y después las huellas: `python instalador/generar_lock.py` (escribe `requirements-instalar.txt`; una prueba falla si quedan desfasados).
3. Cambia `PYTHON_PROBADO` en `instalador/instalar.py`, la versión en `INSTALAR.bat`, `EJECUTAR_PORTAL.bat` y `EJECUTAR.bat`, y publica la versión como cualquier corrección.

## Instalación para desarrollo (con git)

Requiere Python 3.13 (con otra versión el programa funciona, pero las pruebas de versiones fijas avisan).

El repositorio debe quedar en una carpeta llamada **`_Programa`** dentro de `Honorarios/`, porque el programa busca `_Configuracion` y las carpetas de los años un nivel arriba de ella:

```
cd ruta\a\Honorarios
git clone <url-del-repositorio> _Programa
cd _Programa
pip install -r requirements.txt -c requirements-lock.txt
```

`pywin32` solo se necesita para la validación con SAP (Windows).

## Uso

**Doble clic en `EJECUTAR.bat`**, o desde una terminal:

```
python _Programa.py
```

El programa pregunta:

1. **Año y mes** del periodo. Cada pregunta dice qué valor se usa si solo presionas Enter (por ejemplo, `Año (Enter = 2026, año actual)`); al terminar, el programa confirma el periodo: "Periodo a conciliar: septiembre 2026".
2. **Abrir SAP.** La validación contra SAP es **obligatoria** (a petición de Contabilidad): el programa pide abrir SAP e iniciar sesión y espera a que presiones Enter. Si ese mes ya se procesó en esa computadora, pregunta si reusar lo consultado a SAP en esa fecha (por defecto, no: se consulta todo de nuevo).

   **Requisitos:** SAP abierto en el sistema productivo con la sesión iniciada, acceso a la transacción **FK03** y el **scripting de SAP GUI habilitado** (ambos los otorga IT; la primera vez SAP puede preguntar si permites que un script se conecte: responde que sí). El programa usa la **primera ventana de SAP** que abriste y navega ahí a la FK03, así que guarda lo que tengas en esa ventana. Mientras consulta, **no uses SAP**; el resto de la computadora sí se puede usar (el programa controla SAP por dentro, no mueve el mouse ni teclea).

Si SAP no está abierto (o el SAT no responde), el reporte se genera, pero **ninguna línea queda "Cruzado"** y el historial de UUID **no se registra**: hay que volver a correrlo con SAP abierto y conexión.

Después de leer las facturas, el programa **siempre** consulta su estatus en el SAT (requiere internet; ver [Estatus en el SAT](#estatus-en-el-sat)).

El reporte queda en `<Año>/<Mes>/Salida/Conciliacion_AAAA-MM_fecha_hora.xlsx`. Cada corrida genera un archivo nuevo; nunca se sobrescribe uno anterior.

## Portal web (en prueba)

La misma conciliación, desde el navegador y sin carpetas: se elige el mes, se arrastran el TXT y los XML (o un `.zip`), se da clic en **Procesar** y se descarga el Excel. El resultado muestra cuántas líneas hay en cada estatus, el resumen del SAT y la lista **"Qué hacer"** con quién debe actuar.

El historial de UUID **no** se registra al procesar: solo cuando alguien da clic en **"Confirmar y registrar en el historial"**. Así, una corrida de prueba no deja rastro.

Para probarlo en tu computadora:

```
pip install -r requirements-portal.txt -c requirements-lock.txt
```

y doble clic en **`EJECUTAR_PORTAL.bat`**. **No aparece ninguna ventana negra:** el portal arranca en segundo plano y se abre en el navegador predeterminado de Windows (en una pestaña nueva si el navegador ya estaba abierto; si estaba cerrado, lo abre). La primera vez puede tardar unos segundos. Si el portal ya estaba abierto, el doble clic solo abre otra pestaña.

**Clics de más:** si alguien da doble clic varias veces mientras el portal arranca, solo el primero trabaja; los demás se retiran sin hacer nada (nunca se arranca un segundo portal ni salen ventanas de error falsas), y se abre **una sola** pestaña. Para lograrlo, el lanzador toma un candado (`portal-honorarios-8501.lock` en la carpeta temporal de Windows, fuera de OneDrive) que Windows suelta solo cuando el lanzador termina, aunque se caiga; además, un lanzador nunca vive más de 3.5 minutos, así que un lanzador atorado no deja el acceso sin funcionar.

**Portal trabado:** si el portal ocupa el puerto pero no contesta durante unos 20 segundos, el doble clic detiene sus procesos (solo los de este portal) y arranca uno nuevo. Un portal ocupado conciliando sí contesta, así que no se reinicia a la mitad.

**Para cerrarlo:** botón **"Cerrar el portal"** en el menú de la izquierda. Detiene **todo** lo del portal, como un Ctrl + C: este portal, cualquier portal viejo que se haya quedado abierto y el lanzador. Los reconoce porque son Python corriendo `portal\app.py` o `portal\iniciar.py` de esta instalación; **cualquier otro Python (por ejemplo, el de VS Code) no se toca**. Después muestra la página "El portal se cerró" (en lugar del "Connection error… restart it in your terminal" de Streamlit). Si se cierran todas sus pestañas sin usar el botón, el portal se apaga solo a los 15 minutos. Antes de actualizar el programa, ciérralo.

Si el portal no puede iniciar, aparece una ventana de Windows con el motivo; el detalle completo queda en `%LOCALAPPDATA%\Honorarios\portal.log` (fuera de OneDrive; se reescribe cada vez que se abre el portal). Para ver los mensajes en una consola (útil para diagnosticar), desde la carpeta `_Programa`: `python portal\iniciar.py --consola` (mismo arranque y mismas protecciones que el acceso directo; `python -m streamlit run portal/app.py` no las aplica).

Si el navegador dice **"localhost rechazó la conexión"** (`ERR_CONNECTION_REFUSED`), el portal no está corriendo: vuelve a dar doble clic en `EJECUTAR_PORTAL.bat`. No es el firewall; `localhost` es la misma computadora.

El portal solo acepta conexiones desde la misma computadora: los compañeros de la red no pueden abrirlo. Está fijo en `.streamlit/config.toml` y además en el lanzador (`OPCIONES_STREAMLIT` en `portal/iniciar.py`), así que cambiar `config.toml` no lo abre a la red. Un servidor interno no es un cambio de configuración: el portal no tiene usuario, contraseña ni HTTPS, y antes necesita autenticación, cifrado y revisión de Ciberseguridad (ver [Datos y seguridad](#datos-y-seguridad)).

| Página | Qué hace |
|---|---|
| Conciliador Automático | Subir archivos → Procesar → resumen → Descargar Excel → Confirmar |
| Generar TXT | Subir el Excel con las líneas aprobadas → revisar → Formatear folios → Descargar el TXT para SAP |

Las páginas de Catálogo de Proveedores e Historial de UUID (`portal/paginas/catalogo.py` e `historial.py`) ya no se muestran en el menú porque Contabilidad no las usa. El catálogo y el historial se siguen usando igual en cada conciliación.

- **Periodo:** solo se ofrecen el mes actual y los anteriores (y los 12 meses del año anterior, para conciliar diciembre en enero). Un periodo que todavía no llega no se puede conciliar, ni en el portal ni en el `.bat`. La fecha de "hoy" se toma de internet (encabezado `Date` del servicio del SAT, al que la red ya tiene acceso) para no depender del reloj de la PC; sin conexión se usa el reloj de la PC y el portal lo indica. Nada depende de un año fijo: al cambiar de año, las opciones y las carpetas `Reportes\<Año>\<MM Mes>` se ajustan solas (hay una prueba que simula 2027 y 2031).
- **Mes que no es el vigente:** el portal avisa y pide marcar *"Sí, quiero conciliar &lt;mes&gt;"* antes de habilitar Procesar. Un recordatorio fijo indica que el TXT, los XML y el mes elegido deben coincidir.
- **Archivos de otro mes:** si la mayoría de las fechas del TXT y de los XML no son del mes elegido, la conciliación se detiene **antes** de consultar al SAT y a SAP con *"Los archivos parecen ser de enero 2026 y elegiste febrero 2026…"* (portal y `.bat`). Sin esto, todas las facturas saldrían "Rechazar" por estar fuera del periodo.
- **Confirmar no borra un mes ya registrado:** confirmar reemplaza lo que el historial tenía de ese mes. Si la nueva conciliación tiene menos de la mitad de los UUID ya registrados (por ejemplo, archivos equivocados), el historial no se toca y se avisa.
- **Mientras procesa**, un aviso azul recuerda que cambiar de página o cerrar la pestaña detiene la conciliación. No se bloquea nada (se puede hacer scroll para ver el avance): si alguien da clic en otra página del menú o en "Cerrar el portal", el navegador pregunta si de verdad quiere salir; al cerrar o recargar la pestaña, el navegador muestra su aviso de "¿Salir del sitio?". La decisión es del usuario.
- **Logo:** guarda el logo de la empresa como `portal/recursos/logo.png` (ver `portal/recursos/LEEME.md`); aparece arriba del título y del menú.

El portal corre en la computadora de cada usuario y usa **su** sesión de SAP GUI, igual que el `.bat`: antes de dar clic en **Procesar**, abre SAP e inicia sesión. Sin SAP (o sin respuesta del SAT) el botón "Confirmar" queda bloqueado.

### Trabajo de varios días

La conciliación de un mes puede llevar días (revisar errores, esperar correcciones). **El trabajo se guarda solo** en la carpeta del mes (`Reportes\<Año>\<MM Mes>\Trabajo_AAAA-MM.pkl`) cada vez que se procesa o se actualizan las consultas.

- Al abrir el portal aparece **"Continuar donde te quedaste"** con los últimos meses trabajados. Un clic recupera todo: el resultado, lo consultado al SAT y a SAP (sin volver a consultar) y el TXT. El selector se mueve al mes recuperado.
- **Si pasaron días**, el SAT o SAP pudieron cambiar (una factura cancelada, una corrección de Customer Services). En el resultado, **"Consultas al SAT y a SAP"** muestra la fecha de la última consulta y los botones **Actualizar SAT** y **Actualizar SAP**; si tienen 1 día o más, un aviso pide actualizarlas antes de confirmar o descargar el TXT.
- Para empezar el mes de cero, basta con subir los archivos y procesar (reemplaza el trabajo guardado de ese mes).
- El archivo lo escribe y lo lee solo el programa; si está dañado o es de una versión muy distinta, el portal lo dice y se vuelve a procesar.

### Hoja "XML sin TXT": las facturas que ninguna línea usó

Cada fila es un XML que se leyó bien pero que no cruzó con ninguna línea del TXT. Su **MotivoRechazo** dice qué tiene mal **esa factura**, no la línea: si el médico mandó varias, la línea puede estar Cruzado con otra factura suya. Para que no parezca que se cruzó la equivocada:

- La columna **Qué pasó** (justo después del nombre del archivo) lo explica en palabras. Por ejemplo: *"El médico mandó 2 XML y tiene 1 línea(s) en el TXT. Su línea 45 ya cruzó con otra factura: B1234.xml (del 2026-10-12). Esta factura no se usó por lo que dice MotivoRechazo."* Otros casos: el médico no tiene línea en el TXT, su RFC no está en el catálogo, su línea quedó "Diferencia de importe" (el TXT paga tanto y la factura es por tanto), o la factura es válida pero no hay otra línea para pagarla (¿falta la línea o facturó de más?).
- En **Detalle**, la columna **Otros XML del médico** avisa en la línea cuando a ese médico le sobró alguna factura: *"1 XML más de este médico no se usó (hoja XML sin TXT): …"*.

### Excel actualizado

El botón **Descargar Excel** de Conciliador Automático siempre da el reporte **actual**: si se actualizó el SAT o SAP, genera un Excel nuevo con los estatus actualizados (la hoja Info conserva el nombre y la huella SHA-256 del TXT original).

### Facturas pendientes: refacturas y XML que faltaban (segunda vuelta)

Las líneas que no cruzaron (Sin XML, Rechazar, Diferencia de importe, Detenida…) no detienen a las demás: lo que cruzó se paga con su TXT y lo pendiente se resuelve después, en el mismo mes. Cuando llegan las refacturas o los XML que faltaban:

1. En **Conciliador Automático**, con el mes abierto (o con "Continuar donde te quedaste"), sección **Facturas pendientes**: se suben **solo los XML nuevos** (sueltos o en .zip) y se da clic en **Buscar en las líneas pendientes**.
2. **Lo cruzado no se toca:** las líneas Cruzado, Cruzado (pago quebrado) y Cruzado con observaciones quedan igual, con su misma factura. Ni sus facturas ni copias de ellas (mismo UUID) pueden cruzar otra línea.
3. Los XML nuevos se consultan en el SAT (solo esos; SAP se reusa) y se buscan, con las mismas reglas, entre las líneas pendientes, junto con las facturas del mes que ninguna línea cruzada usa. Si el médico mandó una factura nueva, el cruce prefiere la válida a la que se rechazó. Un XML que ya estaba en el mes (mismo UUID) se ignora. Sin conexión con el SAT no se agrega nada.
   - La página dice cuántas líneas están pendientes y de qué estatus (por ejemplo: *156 líneas pendientes: 137 Sin XML, 11 Rechazar, 4 Diferencia de importe, 3 ID no está en catálogo, 1 Línea sin IDProveedor*).
   - **El catálogo se vuelve a leer en cada búsqueda:** un médico dado de alta después de procesar el mes ya cuenta. Si su XML ya había llegado (estaba en "XML sin TXT"), basta con dar clic en **Buscar en las líneas pendientes** sin subir nada. Los proveedores que nunca se consultaron en SAP (los nuevos en el catálogo) se consultan en ese momento si SAP está abierto; si no, su línea queda Cruzado con observaciones ("Validaciones no realizadas").
   - Una línea sin IDProveedor no puede cruzar: el TXT no dice a quién pagarle (se corrige con quien manda el TXT).
4. **Fecha de la refactura:** un CFDI no se puede fechar hacia atrás, así que la refactura de septiembre llega con fecha de octubre. Una factura agregada como pendiente se acepta si es del periodo o del mes siguiente, y queda **Cruzado con observaciones** ("posterior al periodo: se aceptó porque se agregó como pendiente el …"). Más tarde, o en la primera vuelta, la regla es la de siempre: la factura debe ser del mes.
5. El portal muestra qué líneas cambiaron (antes → ahora), el Excel se actualiza (la hoja Info registra cada vuelta) y el trabajo del mes queda guardado. El nombre del XML lleva la carpeta de la vuelta (`Pendientes 16-10-2026 1140/...`).
6. **Para pagar lo nuevo:** en la hoja **TXT para SAP**, la columna **TXT generado** marca las líneas que ya salieron en un TXT (fecha y Excel). Dentro de cada estatus, primero va lo que todavía no tiene TXT. Se arma el Excel con las líneas Cruzado sin TXT generado y se sube en Generar TXT.
7. Al terminar, **Confirmar conciliación** registra en el historial también los UUID nuevos.

Generar TXT **no deja descargar** líneas que ya salieron en un TXT anterior (posible pago doble), salvo que se marque "La carga anterior a SAP falló y hay que mandarlas otra vez". Además, llevan el mismo folio: si en realidad sí se cargaron, SAP las rechaza.

### Generar TXT (el TXT que se carga a SAP)

El TXT para SAP **ya no se edita en el portal**: se genera desde un Excel con **solo las líneas listas para pagar**. Lo que sigue en revisión no detiene a lo demás: se queda fuera y se paga después, cuando esté listo. El analista lleva su propio control de lo pendiente, en coordinación con las áreas involucradas.

1. **Excel de la conciliación → hoja "TXT para SAP".** Trae las líneas del TXT original con sus columnas tal cual (A a J, como texto) y, de apoyo, Estatus, Médico y Línea (K a M). El analista filtra por Estatus y deja solo lo que se va a pagar: Cruzado, Cruzado (pago quebrado), Cruzado con observaciones… lo que él considere correcto.
2. **Generar TXT → subir ese Excel.** Puede ser el mismo Excel de la conciliación (se lee la hoja "TXT para SAP") o uno nuevo con la plantilla de la página. Se leen solo las columnas **A a J**, en el orden del TXT: Fecha, IDProveedor, Folio, Monto, X, Tax, 1, Centro de Costo, Cuenta Contable, PM. El encabezado es opcional. **La Referencia no se captura:** la arma el portal (ver abajo).
3. **El portal revisa, no edita.** Si una celda no sirve, dice en qué fila y columna, y no deja generar el TXT hasta que se corrija **en el Excel**: fecha inválida, monto vacío, negativo o con más de 2 decimales, Tax distinto de E0/F2, X, 1 o PM distintos de X, 1 y T, IDProveedor, folio, centro de costo o cuenta vacíos, una cuenta que Excel convirtió en número (5E+31), tabuladores o caracteres raros. Avisa, sin bloquear, si un IDProveedor no está en el catálogo o si dos filas son la misma factura (mismo ID, folio, monto y fecha: ¿pago doble?).
4. **Vista previa** con el médico (del catálogo), el total de líneas y el monto total, para compararlos con el Excel.
5. **Formatear folios** (obligatorio antes de descargar): pone el folio único y arma la Referencia de cada línea. La vista previa muestra la quincena (de la fecha) y avisa si las fechas son de más de una quincena.
6. **Descargar TXT:** el mismo formato que el TXT original: 11 columnas separadas por tabulador, **sin encabezados**, saltos de línea de Windows (CRLF), ANSI. Lo que Excel cambia al abrir un TXT se regresa a como va en el TXT: fecha `29.09.2026`, monto `9096.00` (2 decimales, sin comas), columna 1 como `001`. Al descargar, los folios quedan en `_Configuracion\HistorialFolios.csv` y se guardan una copia del TXT y el **control de folios** (Excel: de qué folio original salió cada folio nuevo) en `Reportes\<Año>\<MM Mes>\`. El control también se puede descargar.

#### Folios únicos y rastreables ("Formatear folios")

SAP rechaza una línea si el proveedor repite un folio que ya usó, aunque sea de otro mes, y acepta hasta 16 caracteres. El folio que pone el portal mide **16** y se lee:

```
26 18 5113 K7Q3MXP4   →   26185113K7Q3MXP4
│  │   │      └─ 8 caracteres calculados con la factura (huella SHA-256 de IDProveedor + folio original + monto + fecha)
│  │   └─ últimos 4 caracteres del IDProveedor
│  └─ quincena de la fecha (Q18)
└─ año (2026)
```

- **Rastreable:** con solo ver el folio se sabe la quincena y el proveedor. `HistorialFolios.csv` y el control de folios guardan además el folio original, el monto, la fecha, el Excel de donde salió, quién lo generó y cuándo.
- **Misma factura → mismo folio**, siempre y en cualquier computadora. Si una línea que ya se cargó a SAP se vuelve a mandar, SAP la rechaza por folio repetido: **no se paga dos veces**. Si el portal ve que ese folio ya se había generado, lo avisa (con la fecha y el Excel). Un folio que ya tiene este formato se conserva: volver a subir un Excel ya formateado da el mismo TXT.
- Pago quebrado en partes iguales (líneas idénticas): cada repetición agrega su número de parte, así no chocan entre sí.
- 32^8 ≈ 1.1 × 10^12 combinaciones **por proveedor y por quincena** (32 símbolos que no se confunden: sin 0/O ni 1/I). Además, antes de aceptar un folio se comprueba contra los demás del archivo y contra `HistorialFolios.csv`; si chocara con el de otra factura, se recalcula.
- **La Referencia la arma el portal** (la columna 11 del TXT), **en 16 caracteres como máximo**: los últimos 7 caracteres del folio nuevo (los que lo distinguen; año, quincena e ID ya están en la quincena y en la línea) y la quincena de la fecha, unidos con guion bajo: folio `26185113K7Q3MXP4` → Referencia `7Q3MXP4_Q18-2026`. Es la estructura de la Referencia del TXT original (`A1001      Q18-2026`: folio, espacios y quincena), con el folio nuevo. Antes de descargar, el portal comprueba que ningún Folio ni Referencia pase de 16 (`LONGITUD_FOLIO` y `LONGITUD_REFERENCIA` en `honorarios/config.py`). Quincena: del 1 al 15 de cada mes es la primera, del 16 al fin la segunda (Q01 = 1 al 15 de enero, Q18 = 16 al 30 de septiembre).
- El largo está en `LONGITUD_FOLIO` (`honorarios/config.py`).

### Arquitectura: una regla

```
honorarios/   el motor: reglas, cruce, SAT, Excel  <- TODAS las reglas de negocio van aquí
portal/       el portal web: solo archivos, botones y pantallas
_Programa.py  el programa de consola (.bat)
```

**El portal no tiene reglas de negocio.** Solo recibe archivos, llama a `honorarios/servicio.py` (la misma conciliación que usa el `.bat`) y muestra el resultado. Así el portal y el `.bat` siempre dan el mismo resultado.

## Estatus de cada línea

| Estatus | Significa | Quién actúa |
|---|---|---|
| **Cruzado** | Pasó todas las validaciones | Nadie, se puede pagar |
| **Cruzado (pago quebrado)** | Igual que Cruzado, pero varias líneas y facturas se combinan | Nadie |
| **Cruzado con observaciones** | Cruzó, pero hay algo que revisar (ver columna Observaciones) | Revisor |
| **Detenida** | La factura puede estar bien, pero algo interno impide pagarla (SAP sin IH, UUID ya pagado, impuestos que no cuadran…) | Customer Services |
| **Rechazar** | La factura está mal hecha: se pide refacturar | Médico |
| **Diferencia de importe** | Hay XML del proveedor, pero ningún importe coincide | Revisor |
| **Sin XML** | No llegó la factura | Responsables de envío de XML |
| **Proveedor extranjero** | El proveedor tiene ID fiscal de otro país: no emite CFDI, se revisa a mano con su factura del extranjero (y la retención RE en SAP) | Revisor |
| **Proveedor sin RFC / RFC inválido / ID no está en catálogo / Línea sin IDProveedor** | Falla el catálogo o el TXT | Administración del catálogo |

## Reglas de negocio

### Motivos de rechazo (la factura se refactura)

- No está timbrada o no es CFDI versión 4.0.
- Receptor distinto de **SIN070530EI5 / SERVICIOS INTEGRALES NOVA DE MONTERREY / régimen 601 / código postal 66450**.
- Uso de CFDI distinto de **G03 Gastos en general**.
- Fecha de la factura fuera del mes que se concilia.
- Tipo de comprobante distinto de **I (Ingreso)**.
- Método de pago distinto de **PUE** o **PPD**.
- Forma de pago distinta de **03 Transferencia** o **99 Por definir**.
- PPD con una forma de pago distinta de 99.
- **Con consulta al SAT:** la factura está **cancelada**. En Observaciones se pide verificar si el médico envió la factura que la sustituye.
- RFC del emisor con formato inválido, o factura sin conceptos.

### Motivos para detener (se corrige internamente antes de pagar)

- El emisor es persona moral: no le aplican las retenciones de honorarios de persona física.
- El UUID ya se concilió en un periodo anterior (**evita pagos dobles**).
- El total, los conceptos o los impuestos del XML no cuadran con su propio subtotal.
- Moneda distinta de MXN.
- **Con consulta al SAT:** el SAT no encuentra la factura con el UUID, RFC y total del XML (XML alterado, apócrifo o recién timbrado), o el médico pidió cancelarla y la cancelación está en proceso.
- **Con SAP:** la factura lleva IVA pero el proveedor no tiene **IH** (retención de IVA 2/3); el régimen del XML no coincide con RS/RH; hay indicadores de ISR ambiguos.

### Cálculo de impuestos esperado

| Tipo (TXT) | Régimen (XML) | IVA | Ret. IVA | Ret. ISR |
|---|---|---|---|---|
| **E0** médico | 612 No RESICO | — | — | 10% |
| **E0** médico | 626 RESICO | — | — | 1.25% |
| **F2** otro profesional | 612 No RESICO | 16% | 2/3 del IVA | 10% |
| **F2** otro profesional | 626 RESICO | 16% | 2/3 del IVA | 1.25% |

En los F2, el importe del TXT ya incluye el IVA.

### Códigos de retención en SAP (tabla LFBW)

| Código | Significado |
|---|---|
| IH | Retención de IVA honorarios 2/3 (obligatorio en todo F2) |
| RS | Retención ISR RESICO 1.25% |
| RH / RM | Retención ISR honorarios / servicios médicos 10% |
| RE | Retención ISR extranjeros |
| IP | Retención IVA Pitex 100% (no corresponde a honorarios) |

#### Las consultas guardadas (`CodigosSAP_AAAA-MM.xlsx`)

Consultar SAP es lo más lento de la conciliación: SAP GUI se consulta **un proveedor a la vez** (FK03), y con unos 250 médicos son varios minutos. Por eso cada consulta se guarda en `CodigosSAP_AAAA-MM.xlsx`, junto a los reportes de ese mes (`Reportes\<Año>\<MM Mes>\` en el portal; `Salida\` en el `.bat`):

| Columna | Contenido |
|---|---|
| IDProveedor | ID del proveedor en SAP |
| Codigos | Códigos de retención que SAP tenía en ese momento (p. ej. `IH, RS`) |
| Error | Vacío si la consulta salió bien; si no, el motivo |
| Consulta | Fecha y hora de la consulta |

- **Para qué sirve:** si se vuelve a procesar **el mismo mes** (llegó una factura que faltaba, se corrigió el TXT…), se reusan las consultas que salieron bien y **solo se consulta a SAP lo que falta o falló**. Si SAP se cae a media consulta, la siguiente corrida sigue donde se quedó.
- **Por defecto NO se reusan.** La casilla *"Reusar lo consultado a SAP el dd/mm/aaaa a las hh:mm (más rápido)"* solo aparece si el mes ya se procesó en esa computadora, y viene **desmarcada**: cuando se vuelve a correr un mes casi siempre es porque se corrigió algo, y muchas veces es en SAP (por una línea Detenida: se agregó IH, se cambió RS↔RH). Márcala solo si desde esa fecha no se cambió nada en SAP. En el `.bat`, la pregunta equivalente tiene "no" como respuesta por defecto.
- La fecha que se muestra es la de la consulta **más antigua** que se reusaría; cada proveedor conserva en el archivo la fecha en que realmente se consultó.
- **Es por mes:** un mes nunca usa las consultas de otro, así que si un médico pasa de RH a RS (o al revés) se detecta en el mes en que cambió.
- **Evidencia:** queda como constancia de qué retenciones tenía SAP cuando se concilió.
- **Se puede borrar** sin perder nada: la siguiente corrida vuelve a consultar todo a SAP (solo tarda más).

## Estatus en el SAT

El portal de verificación del SAT tiene captcha porque es para personas. El programa no lo usa: consulta el **servicio web oficial `ConsultaCFDIService`**, el que el SAT publica para consultas automatizadas. No tiene captcha, no pide contraseña ni e.firma y responde lo mismo que el portal.

- Por cada factura solo se envía la expresión de su código QR: UUID, RFC emisor, RFC receptor, total y los últimos 8 caracteres del sello.
- Se consultan 4 facturas a la vez (`CONSULTAS_SAT_SIMULTANEAS`); cada UUID una sola vez.
- Si no hay internet, se sabe con la primera consulta y el programa termina el reporte, pero **ninguna línea queda "Cruzado"**: las que cruzaron salen "Cruzado con observaciones" con "Validaciones no realizadas: estatus del CFDI en el SAT". Hay que volver a correrlo con conexión antes de pagar.
- Si el médico canceló una factura y mandó la que la sustituye por el mismo importe, se cruza la nueva (queda "Cruzado", con la nota "Sustituye a la factura cancelada…"; no se toma como facturación doble) y la cancelada aparece en "XML sin TXT" indicando cuál la sustituyó.
- Si no llegó la sustituta, la línea sale "Rechazar" con la observación "Verificar si el médico envió la factura que sustituye a esta, que está cancelada en el SAT".
- El Resumen trae el indicador "Líneas con factura cancelada en el SAT".
- **Alerta EFOS (artículo 69-B):** la misma respuesta del SAT trae el campo `ValidacionEFOS`, que indica si el emisor está en la lista de EFOS (Empresas que Facturan Operaciones Simuladas, listado definitivo del artículo 69-B del CFF). Si hay alerta, la línea sale "Cruzado con observaciones" con "Alerta EFOS…: revisar antes de pagar". El Resumen trae el indicador "Líneas con alerta EFOS (69-B)".

  | Código `ValidacionEFOS` | Significado (Documentación del Servicio de Consulta de CFDI del SAT, v1.4) | Columna **EFOS** | Resultado |
  |---|---|---|---|
  | 100 | El RFC del emisor **sí** está en la lista de EFOS | Revisar: en lista EFOS (código 100) | Alerta |
  | 200 | El RFC del emisor **no** está en la lista | No listado | Sin alerta |
  | 201 | Ni el emisor ni los RFC "a cuenta de terceros" de la factura están en la lista | No listado | Sin alerta |
  | Otro | Código no documentado (p. ej. uno que el SAT agregue después) | Revisar: código … no documentado por el SAT | Alerta (se prefiere revisar de más) |
- Si todas las facturas están vigentes, el resultado es exactamente el mismo que sin consultar al SAT (hay una prueba automática que lo verifica).
- La columna **EstadoSAT** del Detalle (y de "XML sin TXT") dice Vigente, **En proceso de cancelación**, Cancelado, No encontrado o Sin respuesta, con color: rojo cancelada, naranja en proceso de cancelación, amarillo no encontrada o sin respuesta. "En proceso de cancelación" es un estatus aparte (sigue vigente, pero el médico pidió cancelarla y el SAT espera que la empresa acepte o rechace; si no responde en 3 días hábiles, se cancela sola): al filtrar por "Vigente" no se mezcla. Su línea queda Detenida. El Resumen trae el indicador "Líneas con factura en proceso de cancelación" y el portal, "SAT: en proceso de cancelación".
- Es una regla obligatoria: se consulta en cada corrida, sin preguntar.

## Tolerancias

Acordadas con Contabilidad (en `config.py`; la hoja **Tasas** del Excel muestra los valores de la corrida):

| Parámetro | Valor | Qué controla |
|---|---|---|
| `UMBRAL_OK` | 0.05 | Hasta aquí la diferencia se etiqueta "OK" (centavos de redondeo) |
| `TOLERANCIA_IMPUESTOS` | 0.50 | Variación de impuestos aceptada como "Redondeo"; arriba sale "Revisar" y la línea se detiene. También es el margen para que el Total y los conceptos del XML cuadren con su subtotal |
| `TOLERANCIA` | 0.50 | Diferencia máxima entre el TXT y la factura para cruzarlos. Entre 0.05 y 0.50 cruza con observación; más de 0.50 sale "Diferencia de importe" |

## Configuración

Todo lo ajustable está en **`honorarios/config.py`**: datos del receptor, tolerancias, tasas, formas y métodos de pago aceptados, usos de CFDI y mensajes. Cambiar una regla de negocio normalmente es cambiar una línea de ese archivo. Ejemplo: para aceptar otra forma de pago, se agrega su clave en `FORMAS_PAGO_OK`.

## Estructura del código

| Archivo | Contenido |
|---|---|
| `_Programa.py` | Punto de entrada |
| `honorarios/config.py` | Configuración y reglas ajustables |
| `honorarios/utilidades.py` | Lectura de importes, normalización de IDs, RFC y nombres |
| `honorarios/entradas.py` | Carpetas del periodo y lectura de catálogo, TXT y XML |
| `honorarios/historial.py` | Historial de UUID conciliados |
| `honorarios/reglas.py` | Validaciones por factura (CFDI, pago, SAP, impuestos) |
| `honorarios/cruce.py` | Cruce uno a uno y pagos quebrados; estatus final |
| `honorarios/reporte.py` | Excel de salida con fórmulas |
| `honorarios/sap.py` | Consulta obligatoria a SAP GUI (códigos de retención) |
| `honorarios/estado_sat.py` | Consulta del estatus del CFDI en el SAT |
| `honorarios/servicio.py` | La conciliación completa sin preguntas en pantalla (la usan el `.bat` y el portal) |
| `honorarios/principal.py` | Programa de consola: preguntas (periodo, abrir SAP) |
| `portal/` | Portal web (Streamlit): `app.py`, `paginas/` y `archivos.py` (guarda lo que se sube) |
| `tests/` | Pruebas automáticas con datos inventados |

## Pruebas

```
python -m unittest discover -s tests -v
```

Las pruebas generan una carpeta temporal con datos **inventados** (`tests/datos_prueba.py`), un caso por regla de negocio, y verifican el estatus de cada uno. Hay que correrlas antes de subir cualquier cambio. Si una regla cambia a propósito, se actualiza `CASOS_ESPERADOS` en `tests/datos_prueba.py`.

Para ver los datos de prueba en el explorador de archivos:

```
python tests/datos_prueba.py C:\ruta\destino
```

## Datos y seguridad

- El repositorio es **privado**, con acceso para el equipo de Sistemas, según lo autorizado por Ciberseguridad.
- El `.gitignore` bloquea `.xml`, `.txt`, `.xlsx`, `.csv` y demás archivos de datos. **Nunca subas facturas, TXT, catálogos ni reportes reales.**
- Antes de cada commit revisa `git status`: solo deben aparecer archivos `.py`, `.md`, `.bat`, `.toml`, `LICENSE` o `.txt` de configuración del repo.
- El portal no envía estadísticas de uso a Streamlit ni muestra en los errores los enlaces "Ask Google" / "Ask ChatGPT" (que mandarían el detalle del error a esos sitios).
- **Qué se guarda y dónde.** Los archivos que se suben (TXT, XML, ZIP) se copian a una carpeta temporal que se borra al terminar, pero **su contenido sí se conserva**: el Excel de la conciliación y `Trabajo_AAAA-MM.pkl` (que incluye el TXT original y los datos de cada XML) quedan en `Honorarios/Reportes/<Año>/<MM Mes>/`, junto con las copias del TXT para SAP y del control de folios. Si el Escritorio está sincronizado con OneDrive, la carpeta `Honorarios` también lo está (OneDrive corporativo: Sistemas debe confirmar que es un lugar aprobado para estos datos). El programa no cifra esos archivos: los protegen los permisos de Windows, el cifrado del disco y OneDrive, igual que cualquier otro Excel de Contabilidad. Lo que se descarga desde el navegador queda además en la carpeta Descargas.
- La bitácora del portal (`%LOCALAPPDATA%\Honorarios\portal.log`) queda fuera de OneDrive y no anota RFC: los avisos del catálogo solo dan el `IDProveedor`.
- **Todo lo que sale a internet va por HTTPS con el certificado verificado** (Python lo comprueba en cada conexión): el SAT, PyPI y python.org. El portal usa `http://localhost`, que nunca sale de la computadora: solo escucha en `127.0.0.1`, así que nadie de la red puede abrirlo. Esa configuración (`address = "localhost"`, protección XSRF y CORS, `allowedHosts` contra "DNS rebinding", sin estadísticas) está en `.streamlit/config.toml` **y además** `portal/iniciar.py` la pasa en la línea de comandos, que gana sobre cualquier otra configuración o variable `STREAMLIT_*`. Si una página web abierta en el navegador intenta conectarse al portal desde otro sitio, se rechaza sin consultar nada en internet (Streamlit preguntaría a `checkip.amazonaws.com` la IP pública de la PC; el lanzador se la da ya resuelta). El portal **no tiene usuario ni contraseña**: es seguro solo porque es local. Ponerlo en la red (un "servidor interno") requiere antes autenticación, HTTPS y revisión de Ciberseguridad.
- Los XML se leen con un lector protegido contra XML malicioso (sin entidades externas, protección contra "billion laughs"). Un ZIP subido no puede escribir fuera de su carpeta: se quitan `..`, la raíz y los prefijos de rutas de red (`\\servidor\...`); los caracteres que Windows no acepta (`:` `|` `?` `*` `<` `>` `"`) pasan a `_`, así que `D:` no cambia de unidad y `Factura 15:09:2026.xml` se guarda como `Factura 15_09_2026.xml`; los nombres de dispositivo (`CON`, `NUL`, `COM1`…) llevan `_` delante y se quitan los puntos y espacios del final. Además se comprueba que cada archivo quede dentro de la carpeta, y si un XML no se puede guardar, el portal dice cuál. El Excel de la conciliación y el control de folios guardan como texto cualquier valor que empiece con `=` (por ejemplo, un nombre del catálogo o de un XML): Excel no lo ejecuta como fórmula.
- **Trabajo guardado firmado:** `Trabajo_AAAA-MM.pkl` lleva una firma HMAC-SHA256 con una clave de 256 bits que solo existe en esa computadora (`%LOCALAPPDATA%\Honorarios\clave_trabajo.bin`, fuera de OneDrive). Antes de leerlo se comprueba la firma; si el archivo se modificó, lo puso otra persona o viene de otra computadora, no se abre (el formato podría ejecutar código si alguien lo alterara).
- **Librerías con huella:** el instalador rechaza cualquier archivo que no coincida con su huella SHA-256 (ver [Versiones fijas](#versiones-fijas)). El instalador de Python de respaldo (si no hay `winget` o no pudo instalarlo) solo se ejecuta con la firma digital válida de la Python Software Foundation.
- No se guardan contraseñas ni credenciales en ningún lado; SAP se usa en solo lectura con la sesión de quien concilia. Para eso Sistemas debe tener habilitado **SAP GUI Scripting** para ese usuario: el programa solo consulta (transacción de visualización FK03), pero el scripting trabaja con los permisos SAP del usuario.
- Los datos de prueba (`tests/`) son inventados: una prueba revisa que ningún RFC de las pruebas pase el dígito verificador del SAT (salvo el de la sociedad, que es público).
- La única conexión a internet durante la conciliación es al servicio oficial del SAT (`consultaqr.facturaelectronica.sat.gob.mx`), en cada corrida, y envía únicamente los datos del código QR de cada factura (ver [Estatus en el SAT](#estatus-en-el-sat)); nada del TXT ni del catálogo sale de la computadora. La otra integración es la consulta obligatoria, de solo lectura, a SAP GUI con la sesión de quien concilia.

## Licencia

Software propietario de Servicios Integrales Nova de Monterrey, S.A. de C.V. Todos los derechos reservados; uso interno únicamente. Ver [LICENSE](LICENSE).

**Autor:** Eduardo Pecina, quien diseñó y desarrolló el programa para la empresa.

## Historial de cambios

Ver [CHANGELOG.md](CHANGELOG.md).
