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

El programa vive en `_Programa` y ubica todo lo demás a partir de ahí, así que funciona igual en cualquier computadora:

```
Honorarios/
├── _Programa/                      ← este repositorio
├── _Configuracion/
│   ├── CatalogoProveedoresTM26.xlsx   (tabla de Excel llamada "Catalogo")
│   └── HistorialUUID.csv              (se crea solo)
└── 2026/
    └── 09 Sep/
        ├── Entrada/
        │   ├── XML/                   (facturas .xml, pueden ir en subcarpetas)
        │   └── Honorarios.txt         (un solo .txt, el nombre no importa)
        └── Salida/                    (aquí se guarda el reporte)
```

## Instalación

Requiere Python 3.10 o superior.

El repositorio debe quedar en una carpeta llamada **`_Programa`** dentro de `Honorarios/`, porque el programa busca `_Configuracion` y las carpetas de los años un nivel arriba de ella:

```
cd ruta\a\Honorarios
git clone <url-del-repositorio> _Programa
cd _Programa
pip install -r requirements.txt
```

`pywin32` solo se necesita para la validación con SAP (Windows).

## Uso

**Doble clic en `EJECUTAR.bat`**, o desde una terminal:

```
python _Programa.py
```

El programa pregunta:

1. **Año y mes** del periodo. Cada pregunta dice qué valor se usa si solo presionas Enter (por ejemplo, `Año (Enter = 2026, año actual)`); al terminar, el programa confirma el periodo: "Periodo a conciliar: septiembre 2026".
2. **Abrir SAP.** La validación contra SAP es **obligatoria** (a petición de Contabilidad): el programa pide abrir SAP e iniciar sesión y espera a que presiones Enter. Si hay consultas SAP guardadas de ese mes, pregunta si reusarlas.

   **Requisitos:** SAP abierto en el sistema productivo con la sesión iniciada, acceso a la transacción **FK03** y el **scripting de SAP GUI habilitado** (ambos los otorga IT; la primera vez SAP puede preguntar si permites que un script se conecte: responde que sí). El programa usa la **primera ventana de SAP** que abriste y navega ahí a la FK03, así que guarda lo que tengas en esa ventana. Mientras consulta, **no uses SAP**; el resto de la computadora sí se puede usar (el programa controla SAP por dentro, no mueve el mouse ni teclea).

Si SAP no está abierto (o el SAT no responde), el reporte se genera, pero **ninguna línea queda "Cruzado"** y el historial de UUID **no se registra**: hay que volver a correrlo con SAP abierto y conexión.

Después de leer las facturas, el programa **siempre** consulta su estatus en el SAT (requiere internet; ver [Estatus en el SAT](#estatus-en-el-sat)).

El reporte queda en `<Año>/<Mes>/Salida/Conciliacion_AAAA-MM_fecha_hora.xlsx`. Cada corrida genera un archivo nuevo; nunca se sobrescribe uno anterior.

## Portal web (en prueba)

La misma conciliación, desde el navegador y sin carpetas: se elige el mes, se arrastran el TXT y los XML (o un `.zip`), se da clic en **Procesar** y se descarga el Excel. El resultado muestra cuántas líneas hay en cada estatus, el resumen del SAT y la lista **"Qué hacer"** con quién debe actuar.

El historial de UUID **no** se registra al procesar: solo cuando alguien da clic en **"Confirmar y registrar en el historial"**. Así, una corrida de prueba no deja rastro.

Para probarlo en tu computadora:

```
pip install -r requirements-portal.txt
```

y doble clic en **`EJECUTAR_PORTAL.bat`**: el navegador se abre solo en http://localhost:8501 en cuanto el portal está listo (la primera vez puede tardar). Deja la ventana negra abierta mientras lo usas; si el portal se detiene, la ventana se queda abierta mostrando el error. O desde una terminal: `python -m streamlit run portal/app.py`.

Si el navegador dice **"localhost rechazó la conexión"** (`ERR_CONNECTION_REFUSED`), el portal no está corriendo: revisa la ventana negra. No es el firewall; `localhost` es la misma computadora.

Si la ventana negra dice **"Port 8501 is not available"**, ya hay otro portal abierto: usa ese (http://localhost:8501) o cierra su ventana negra y vuelve a abrirlo.

Por ahora el portal solo acepta conexiones desde la misma computadora (`address = "localhost"` en `.streamlit/config.toml`): los compañeros de la red no pueden abrirlo. Para el servidor interno, Sistemas cambia esa línea.

| Página | Qué hace |
|---|---|
| Conciliador Automático | Subir archivos → Procesar → resumen → Descargar Excel → Confirmar |
| Catálogo de Proveedores TM26 | Consulta y búsqueda (por ahora se edita en el Excel de `_Configuracion`) |
| Historial de UUID | UUID ya conciliados, por periodo |

- **Periodo:** solo se ofrecen el mes actual y los anteriores (y los 12 meses del año anterior, para conciliar diciembre en enero). Un periodo que todavía no llega no se puede conciliar, ni en el portal ni en el `.bat`. La fecha de "hoy" se toma de internet (encabezado `Date` del servicio del SAT, al que la red ya tiene acceso) para no depender del reloj de la PC; sin conexión se usa el reloj de la PC y el portal lo indica.
- **Mientras procesa**, la página completa queda bloqueada con un aviso: cambiar de página, cambiar un control o cerrar la pestaña detendría la conciliación. Al terminar se libera sola.
- **Logo:** guarda el logo de la empresa como `portal/recursos/logo.png` (ver `portal/recursos/LEEME.md`); aparece arriba del título y del menú.

El portal corre en la computadora de cada usuario y usa **su** sesión de SAP GUI, igual que el `.bat`: antes de dar clic en **Procesar**, abre SAP e inicia sesión. Sin SAP (o sin respuesta del SAT) el botón "Confirmar" queda bloqueado.

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
| **Detenida** | La factura puede estar bien, pero algo interno impide pagarla (SAP sin IH, UUID ya pagado, impuestos que no cuadran…) | Sistemas / Customer Services |
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
- La columna **EstadoSAT** del Detalle dice Vigente, Vigente (cancelación en proceso), Cancelado, No encontrado o Sin respuesta.
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
- El portal (`.streamlit/config.toml`) no envía estadísticas de uso a Streamlit. Lo que se sube se guarda en una carpeta temporal que se borra al terminar; el Excel queda en `<Año>/<Mes>/Salida`, como con el `.bat`.
- La única conexión a internet es al servicio oficial del SAT (`consultaqr.facturaelectronica.sat.gob.mx`), en cada corrida, y envía únicamente los datos del código QR de cada factura (ver [Estatus en el SAT](#estatus-en-el-sat)); nada del TXT ni del catálogo sale de la computadora. La otra integración es la consulta obligatoria, de solo lectura, a SAP GUI con la sesión de quien concilia.

## Licencia

Software propietario de Servicios Integrales Nova de Monterrey, S.A. de C.V. Todos los derechos reservados; uso interno únicamente. Ver [LICENSE](LICENSE).

## Historial de cambios

Ver [CHANGELOG.md](CHANGELOG.md).
