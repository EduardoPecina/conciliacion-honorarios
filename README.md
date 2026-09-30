# Conciliación de Honorarios Médicos

Herramienta que automatiza la revisión mensual de los pagos de honorarios a médicos y profesionales de la salud. Cruza el archivo de pagos que se carga a SAP (TXT) contra el catálogo de proveedores y las facturas electrónicas (XML CFDI 4.0), aplica las reglas fiscales y de negocio, y genera un reporte en Excel con el estatus de cada pago.

> **Este repositorio contiene solo código.** Ningún dato real (XML, TXT, catálogos, reportes) se sube aquí. Ver [Datos y seguridad](#datos-y-seguridad).

---

## Qué resuelve

Antes, cada factura se revisaba a mano contra el pago correspondiente. El programa hace en segundos lo siguiente:

1. **Cruza** cada línea del TXT con su factura: `IDProveedor → RFC (catálogo) → XML`, incluyendo pagos quebrados (varias líneas contra una factura, o una línea contra varias facturas).
2. **Valida** cada factura: receptor, periodo, método y forma de pago, timbre, impuestos recalculados, historial de UUID ya pagados.
3. **Opcionalmente consulta SAP** (transacción FK03/BP) para confirmar que las retenciones configuradas del proveedor coinciden con la factura.
4. **Genera un Excel** donde cada número se puede rastrear hasta su fórmula.

```
TXT de pagos ──► Catálogo (ID → RFC) ──► XML de la factura
     │                                        │
     └──────────── reglas de negocio ◄────────┘
                          │
                   (opcional) SAP
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

1. **Año y mes** del periodo (Enter = mes actual).
2. **¿Validar contra SAP?** Responde `s` solo si SAP está abierto con tu sesión iniciada. Mientras consulta, no toques el mouse ni el teclado.

El reporte queda en `<Año>/<Mes>/Salida/Conciliacion_AAAA-MM_fecha_hora.xlsx`. Cada corrida genera un archivo nuevo; nunca se sobrescribe uno anterior.

## Estatus de cada línea

| Estatus | Significa | Quién actúa |
|---|---|---|
| **Cruzado** | Pasó todas las validaciones | Nadie, se puede pagar |
| **Cruzado (pago quebrado)** | Igual que Cruzado, pero varias líneas y facturas se combinan | Nadie |
| **Cruzado con observaciones** | Cruzó, pero hay algo que revisar (ver columna Observaciones) | Revisor |
| **Detenida** | La factura puede estar bien, pero algo interno impide pagarla (proveedor bloqueado, SAP sin IH, UUID ya pagado…) | Sistemas / Customer Services |
| **Rechazar** | La factura está mal hecha: se pide refacturar | Médico |
| **Diferencia de importe** | Hay XML del proveedor, pero ningún importe coincide | Revisor |
| **Sin XML** | No llegó la factura | Responsables de envío de XML |
| **Proveedor extranjero** | El proveedor tiene ID fiscal de otro país: no emite CFDI, se revisa a mano con su factura del extranjero (y la retención RE en SAP) | Revisor |
| **Proveedor sin RFC / RFC inválido / ID no está en catálogo / Línea sin IDProveedor** | Falla el catálogo o el TXT | Administración del catálogo |

## Reglas de negocio

### Motivos de rechazo (la factura se refactura)

- No está timbrada o no es CFDI versión 4.0.
- Receptor distinto de **SIN070530EI5 / SERVICIOS INTEGRALES NOVA DE MONTERREY / régimen 601**.
- Uso de CFDI distinto de **G03 Gastos en general**.
- Fecha de la factura fuera del mes que se concilia.
- Tipo de comprobante distinto de **I (Ingreso)**.
- Método de pago distinto de **PUE** o **PPD**.
- Forma de pago distinta de **03 Transferencia** o **99 Por definir**.
- PPD con una forma de pago distinta de 99.
- RFC del emisor con formato inválido, o factura sin conceptos.

### Motivos para detener (se corrige internamente antes de pagar)

- Proveedor bloqueado en el catálogo.
- El emisor es persona moral: no le aplican las retenciones de honorarios de persona física.
- El UUID ya se concilió en un periodo anterior (**evita pagos dobles**).
- El total, los conceptos o los impuestos del XML no cuadran con su propio subtotal.
- Moneda distinta de MXN.
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
| `honorarios/sap.py` | Consulta opcional a SAP GUI |
| `honorarios/principal.py` | Flujo principal |
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
- Antes de cada commit revisa `git status`: solo deben aparecer archivos `.py`, `.md`, `.bat`, `LICENSE` o `.txt` de configuración del repo.
- El programa no se conecta a internet ni a servicios externos. La única integración es la consulta opcional, de solo lectura, a SAP GUI.

## Licencia

Software propietario de Servicios Integrales Nova de Monterrey, S.A. de C.V. Todos los derechos reservados; uso interno únicamente. Ver [LICENSE](LICENSE).

## Historial de cambios

Ver [CHANGELOG.md](CHANGELOG.md).
