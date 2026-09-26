# Mapa de campos por formato

Los campos marcados **FASE 2** son los únicos que dependen de la planilla para escribirse.
`C42` es la excepción: se escribe en fase 1 (igual a `C41`) pero la fase 2 lo valida contra la
planilla. Todo lo demás se diligencia en la fase 1, antes de tenerla.

Fuente de cada dato: `CONTRATO` (PDF en `Contratos/`), `ANEXO` (PDF en `Anexos/`),
`PERFIL` (`registro/perfil.json`), `CALCULADO` por la skill, `PREGUNTAR` al usuario.

---

## 1. Excel — `MSCC_U-FT-12.010.069` · hoja `Certificación Mensual`

Solo se diligencian las celdas azules. Todo lo demás es fórmula y lo recalcula Excel.
Máximo **5 contratos**: filas 23 a 27.

`excel_fill.py` limpia `B23:I27`, `D36:E36`, `C41`, `C42` y `E45:E48` antes de escribir, porque
el archivo de `Formatos/` **viene con un ejemplo ya diligenciado**. `D36` y `E45` son celdas
combinadas: limpiar solo su ancla no borra nada, hay que nombrar el rango completo.
En la fase 2 se corre con `--no-clear` para no borrar lo de la fase 1.

| Celda | Campo | Fuente |
|---|---|---|
| `D17` | Nombre del contratista | PERFIL |
| `D18` | Número de documento | PERFIL |
| `D19` | Correo institucional | PERFIL |
| `I17` | Fecha de diligenciamiento | CALCULADO (hoy) |
| `I18` | ¿Es pensionado? `SI`/`NO` | PERFIL |
| `B23` | Nombre o razón social del contratante | Fijo `UNIVERSIDAD NACIONAL DE COLOMBIA`. Celda **combinada** `B23:B27`: se escribe **una sola vez**, no por fila |
| `C23:C27` | Empresa en QUIPU | **PREGUNTAR** (no está en el contrato) |
| `D23:D27` | Tipo orden contractual | CONTRATO (encabezado: `OSE No. 22`, `OPS No. 487`) |
| `E23:E27` | Número orden contractual | CONTRATO |
| `F23:F27` | Valor total antes de IVA | CONTRATO (campo `VALOR`, sin la contribución especial) |
| `G23:G27` | Fecha de inicio | CONTRATO cláusula `PLAZO` si trae fechas; si solo trae días → **PREGUNTAR** |
| `H23:H27` | Fecha de terminación | igual que la anterior |
| `I23:I27` | Clase de riesgos laborales | ANEXO certificado ARL (puede diferir por contrato); si falta o hay varios → **PREGUNTAR** |
| `D36` | Contrato al que aplican deducciones | PREGUNTAR, solo si `B57 > 95 UVT` |
| `C37` `C38` `C39` `E37` `E38` `E39` | 6 anexos de deducción `SI`/`NO` | PREGUNTAR (por defecto `NO`) |
| `C41` | Periodo de solicitud de pago | Input de la skill (día 1 del mes) |
| `C42` | Periodo de la planilla | **FASE 1** · mismo valor y formato que `C41` · FASE 2 la **valida** contra el periodo de **cobertura** de la planilla del ANEXO y corrige + avisa si difiere |
| `E45` | Declaración disminución base retención | **FASE 2** · `  SI  ` solo si `C41 == C42`; si no, la lista solo ofrece `NO` |

Listas cerradas:

- `D23:D27` → `OPS, OSE, OCO, OFS, OSF, OEF, OFO, ODO, OOF, RAG, RVF, RCO, RPS, RVE`
- `I23:I27` → `Riesgo 1` … `Riesgo 5`
- `E45` → `  SI  ` / `  NO  ` (con los dos espacios a cada lado)

Celdas de control que devuelve `excel_fill.py`:

| Clave | Celda | Uso |
|---|---|---|
| `aporte_salud` | `I36` | informativo |
| `aporte_pension` | `I37` | informativo |
| `fondo_solidaridad` | `I38` | informativo |
| `aporte_arl` | `I39` | informativo |
| `total_aportes` | `I40` | **debe ser ≤ lo pagado en la planilla** |
| `valor_mensualizado_total` | `B57` | umbral de 95 UVT para deducciones |
| `ibc_consolidado` | `D57` | informativo |
| `base_retencion_total` | `I57` | informativo |

---

## 2. Word — `U.FT.12.010.053` Constancia de cumplimiento contractual

Coordenadas **después** de aplanar los content controls, que es lo que hace
`docx_fill.py` al cargar. Todas en `scope: body`.

| Op | Destino | Campo | Fuente |
|---|---|---|---|
| `replace` | `[SEDE]` | Sede | CONTRATO (encabezado) |
| `replace` | `[NOMBRE DE LA DEPENDENCIA]` | Dependencia | CONTRATO (encabezado / `DESTINO`) |
| `replace` | `[NOMBRE INTERVENTOR O SUPERVISOR]` | **Quien firma** | `firmante` — por defecto el del contrato, confirmado con el usuario |
| `replace` | `[No. Identificación]` | Cédula de **quien firma** | `firmante` |
| `replace` | `en la ciudad de _______` | Ciudad de expedición | CONTRATO (`LUGAR DE EJECUCION`) |
| `replace` | `Correo electrónico:` | Correo de **quien firma** | **PREGUNTAR** |
| `replace` | `Teléfono:` | Teléfono de **quien firma** | **PREGUNTAR** |
| `replace` | `Haga clic aquí o pulse para escribir una fecha.` | Fecha de expedición | PREGUNTAR / CALCULADO |
| `set_cell` | tabla 0 `(1,1)` | Modalidad | CONTRATO — usar la etiqueta completa de la lista de abajo |
| `set_cell` | tabla 0 `(1,3)` | Número / Año | CONTRATO |
| `set_cell` | tabla 0 `(1,5)` | Adición No. / Año | PREGUNTAR (normalmente `N/A`) |
| `set_cell` | tabla 0 `(1,7)` | Otrosí No. / Año | PREGUNTAR (normalmente `N/A`) |
| `set_cell` | tabla 0 `(2,1)` | Código empresa SGF–QUIPU | **PREGUNTAR** |
| `set_cell` | tabla 0 `(3,1)` | Contratista | PERFIL |
| `set_cell` | tabla 0 `(4,1)` | Identificación | PERFIL |
| `blanks` | tabla 0 `(5,0)` — 3 espacios | nº planilla · fecha de pago · periodo de cobertura | **FASE 2** · ANEXO planilla |
| `checkbox` | tabla 0 `(6,0)` índices `0` parcial · `1` final · `2` único | Tipo de pago | CONTRATO (`FORMA DE PAGO`) |
| `blanks` | tabla 0 `(6,0)` — 4 espacios: nº parcial · valor parcial · valor final · valor único | Valor autorizado | CALCULADO |
| `checkbox` | tabla 0 `(6,1)` índices `0` excelente · `1` bueno · `2` aceptable | Nivel de satisfacción | PREGUNTAR (por defecto `0`) |
| `checkbox` | tabla 0 `(6,2)` índices `0` SI · `1` N/A | ¿Se recibió informe? | `0` (siempre se adjunta) |

Etiquetas válidas de modalidad (tomadas de la lista del propio formato):
`OCA - Orden contractual de arrendamiento`, `OCO - Orden contractual de consultoría`,
`ODC - Orden contractual de compra`, `ODO -Orden contractual de obra`,
`OPS - Orden contractual de prestación de servicios personales de apoyo a la gestión`,
`OSE - Orden contractual de servicios`, `OSU - Orden contractual de suministros`,
`CCO`, `CDO`, `CDA`, `CDC`, `CPS`, `CSE`, `CSU`, `CIS` (contratos equivalentes).

> El nº de planilla, su fecha de pago y su periodo de cobertura salen **de la planilla misma**.
> La fecha de pago no tiene por qué coincidir con el mes cobrado; el **periodo de cobertura sí**.

---

## 3. Word — `U.FT.12.011.020` Informe de ejecución de actividades

| Op | Destino | Campo | Fuente |
|---|---|---|---|
| `set_cell` | header tabla 0 `(0,1)` | Orden contractual No. | CONTRATO |
| `set_cell` | header tabla 0 `(0,3)` | Del año | CONTRATO |
| `set_cell` | header tabla 0 `(1,1)` | Contratista | PERFIL |
| `set_cell` | header tabla 0 `(1,3)` | C.C. / C.E. | PERFIL |
| `set_cell` | header tabla 0 `(3,2)` `(3,3)` `(3,4)` | Desde: día · mes · año | CALCULADO — `max(primer día del mes, fecha_inicio)` |
| `set_cell` | header tabla 0 `(3,6)` `(3,7)` `(3,8)` | Hasta: día · mes · año | CALCULADO — `min(último día del mes, fecha_terminacion)` |
| `set_cell` | body tabla 0 `(0,0)` | `OBJETO: …` | CONTRATO (`OBJETO GENERAL`) |
| `set_cell` | body tabla 0 `(2,2)` `(2,3)` `(2,4)` | Fecha inicio del contrato | CONTRATO |
| `set_cell` | body tabla 0 `(2,6)` `(2,7)` `(2,8)` | Fecha terminación | CONTRATO |
| `ensure_rows` | body tabla 1, `count = 2 + nº obligaciones` | — | — |
| `repeat_header` | body tabla 1, `rows: 2` | — | — repite las 2 filas de encabezado en cada página cuando el informe cruza de página |
| `set_cell` | body tabla 1 filas `2…`, columnas `0..4` | No. · obligación · actividades · % periodo · % acumulado | CONTRATO / **derivado de la obligación** / CALCULADO |
| `set_cell` | body tabla 2 `(0,1)` `(1,1)` `(2,1)` | Productos entregados | CONTRATO (cláusula `NOTA`) si aplica |
| `replace` | `se firma el presente informe el día de mes de año.` | Fecha de firma | CALCULADO |
| `set_cell` | body tabla 3 `(1,0)` | Nombre del contratista | PERFIL |
| `set_cell` | body tabla 3 `(1,2)` | Nombre del `Vo.Bo.` | `firmante` — **el mismo** que firma la constancia |

Las obligaciones se copian **textuales** del contrato y los porcentajes son calculados.

**`ACTIVIDADES EJECUTADAS` se redacta desde la obligación de esa misma fila**, expresada como
ejecutada en el periodo informado. No se le pregunta al usuario en blanco: el formato define esa
columna como "para el cumplimiento de la obligación en el periodo informado", y la obligación ya
está en el contrato.

Límite: sin hechos concretos que el contrato no respalde — ni cantidades, ni fechas puntuales, ni
entregables inventados. El borrador se le muestra al usuario, que lo ajusta antes de firmar.
