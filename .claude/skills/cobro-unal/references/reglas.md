# Reglas de negocio del cobro

## Qué determina el monto a pagar

`total_aportes` depende de **cuatro cosas y nada más**:

| Dato | De dónde sale |
|---|---|
| Valor del contrato (`F23:F27`) | Contrato |
| Fechas de inicio y terminación (`G`, `H`) | Contrato, o se pregunta si no las trae |
| Clase de riesgo ARL (`I23:I27`) | Certificado ARL, o se pregunta |
| ¿Pensionado? (`I18`) | Perfil |

Verificado corriendo el libro: con solo esos datos devuelve los aportes correctos, **sin** nombre,
correo, código QUIPU, tipo ni número de orden. Esas celdas son de identificación y no entran en
ninguna fórmula de la sección 4.

Por eso el monto se entrega antes de completar el formato: es lo que el usuario necesita para
pagar la planilla, y no cambia después.

## Las dos fases

El cobro se parte en dos porque el Excel es el que **dice cuánto pagar** de seguridad social:
no se puede tener la planilla antes de diligenciarlo.

**Fase 1 — sin planilla.** Se parte en dos:
- **1A**: se pregunta solo lo que bloquea el cálculo y se entrega **el monto a pagar**.
- **1B**: se completa el resto — Excel salvo `E45` (`C42` ya queda con el mismo valor y formato
  que `C41`), informe de ejecución completo, y constancia completa salvo el punto 2. Termina
  esperando.

**Fase 2 — con planilla y ARL.** Se valida el monto, se completa `E45` y la constancia, se
**valida** `C42` contra la planilla y se arman los paquetes.

La frontera es exacta: **solo tres cosas dependen de la planilla.**

| Dónde | Qué | Fase |
|---|---|---|
| Excel `C42` | Periodo de la planilla | Se escribe en 1 (igual a `C41`), se **valida** en 2 |
| Excel `E45` | Declaración disminución base retención | 2 |
| Constancia `(5,0)` | Nº de planilla · fecha de pago · periodo de cobertura | 2 |

`E45` no afecta los aportes (`I36:I40`), solo la base de retención (`I57`). Por eso
`total_aportes` de la fase 1 ya es definitivo.

## Estructura de carpetas

```
Cobros/
├── registro/                  perfil.json · contratos.json   (persiste entre periodos)
└── AAAA-MM/
    ├── certificacion-cedular.xlsx            libro único del periodo · todos los contratos
    ├── ESTADO.md                             fase actual, monto a pagar, pendientes
    ├── TIPO-NNN-AAAA/
    │   ├── constancia.docx                   editable
    │   ├── constancia.pdf                    suelto
    │   ├── informe.docx                      editable
    │   ├── informe.pdf                       suelto
    │   ├── certificacion-cedular.pdf         suelto · mismo libro, D36 = este contrato
    │   └── PAQUETE-TIPO-NNN-AAAA-MM.pdf        agrupado · se crea en la fase 2
    └── TIPO-MMM-AAAA/
        └── …
```

Los formatos Word son únicos por contrato, por eso viven en su subcarpeta.

### El cedular: uno solo, copiado en cada paquete

El Excel se diligencia **una vez por periodo**, no por contrato, y esto es estructural:

- La nota (a) del formato exige relacionar **todas** las órdenes vigentes, una a una.
- Las fórmulas consolidan: `D16 = MAX(SMMLV, MIN(SUM(D11:D15), 25 × SMMLV))`. El piso de
  1 SMMLV y el techo de 25 se aplican al **total**, no a cada contrato. Uno por contrato
  aplicaría el piso varias veces y **pagarías de más**.
- `D36` ("Solicito aplicar mis deducciones al contrato No.") toma su lista desplegable de
  `Datos!$B$11:$B$15`, que son las filas 23 a 27 de la misma hoja. Es un selector **entre** los
  contratos listados ahí, y por eso exige que convivan en un único documento.
- La clase de riesgo se cotiza por la **más alta** entre todos: una sola cotización.

Ahora bien, cada contrato se radica con su propia solicitud de pago. Se exporta **un PDF por
contrato cobrado**, todos desde el mismo libro y **variando únicamente `D36`**, que apunta al
contrato de esa radicación.

Verificado: `D36` no aparece en **ninguna** fórmula del libro. Es declarativo. Cambiarlo no mueve
el agregado — tres corridas cambiando solo ese campo devolvieron `total_aportes` idéntico. Los
contratos relacionados, el IBC, los aportes y la planilla son los mismos en todas las copias.

**Cuidado cuando sí se piden deducciones.** Si alguno de los seis checkboxes está en `SI`, la
instrucción manda: *"Las deducciones se aplicarán únicamente a un solo contrato."* Ahí `D36` debe
ser **el mismo contrato elegido en todas las copias**, no el de cada radicación, o estarías
pidiendo la deducción dos veces. Sin deducciones, `D36` es un campo obligatorio e inerte y lleva
el contrato que se cobra.

### Sueltos y agrupados a la vez

Cada documento diligenciado queda en **tres formas**, y las tres se conservan:

| Forma | Para qué sirve |
|---|---|
| `.docx` | Corregir sin rehacer el proceso, y firmar |
| `.pdf` individual | Enviarlo solo cuando piden un documento puntual, o mandarlo a firma del supervisor |
| Dentro del `PAQUETE-*.pdf` | Radicar la solicitud de pago completa |

Unir **nunca** consume los originales: `merge_pdf.py` solo lee. Después de armar el paquete, los
sueltos siguen ahí y así deben quedar. No hay paso de limpieza, y no debe agregarse uno.

## Vigencia en el periodo

Un contrato entra al Excel del periodo `mm/yyyy` cuando su rango
`[fecha_inicio, fecha_terminacion]` **se solapa** con ese mes, aunque sea un día.
Se relacionan **todos** los contratos vigentes, no solo el que se cobra.

## Periodo del informe

El `PERIODO DEL INFORME` del formato `U.FT.12.011.020` **no es el mes calendario**: es la
intersección entre el mes cobrado y la vigencia del contrato.

```
desde = max(primer_dia_del_mes, fecha_inicio)
hasta = min(ultimo_dia_del_mes, fecha_terminacion)
```

Ninguno de los dos extremos puede salirse de `[fecha_inicio, fecha_terminacion]`: un informe que
reporta días anteriores al inicio o posteriores a la terminación declara ejecución fuera del
contrato, y la UNAL lo devuelve.

| Caso | Mes cobrado | Vigencia | Periodo del informe |
|---|---|---|---|
| Mes intermedio | 09/2026 | 25/05 – 22/12/2026 | 01/09/2026 – 30/09/2026 |
| **Último mes** | 08/2026 | 25/05 – **22/08**/2026 | 01/08/2026 – **22/08**/2026 |
| **Primer mes** | 08/2026 | **14/08** – 30/11/2026 | **14/08**/2026 – 31/08/2026 |
| Mes único | 09/2026 | 04/09 – 19/09/2026 | 04/09/2026 – 19/09/2026 |

El campo vive en el **header** del formato, así que una sola op corrige las tres páginas. La fecha
de firma del informe es otra cosa y sí puede ser posterior a la terminación: certifica cuándo se
suscribe el documento, no qué se ejecutó.

## Meses de ejecución

```
meses = (año_fin - año_inicio) * 12 + mes_fin - mes_inicio + 1
```

Cuenta meses calendario tocados, no días. Del 04/09 al 19/09 es **1** mes.

## Avance del informe de ejecución

```
avance_periodo   = 100 / meses
avance_acumulado = numero_de_mes * (100 / meses)
```

`numero_de_mes` es la posición del mes cobrado dentro del contrato (1-based).
El **último mes siempre cierra en 100 %**: si el redondeo no llega, se ajusta a 100.

## Quién firma cada documento

| Documento | Firma |
|---|---|
| Constancia de cumplimiento | **Solo el supervisor o interventor**: nombre, cédula, correo y teléfono |
| Informe de ejecución | **Ambos**: el contratista firma, el supervisor da el `Vo.Bo.` |
| Certificación cedular | El contratista. Por Central de Pagos no se requiere firma |

### El supervisor del contrato no siempre es quien firma

Los contratos designan al supervisor **"o quien haga sus veces"**. Si ese día firma un encargado,
la constancia debe llevar **los datos de quien firma**, no los del contrato: es el documento que
autoriza el pago y va con nombre y cédula de quien lo suscribe.

Por eso el registro guarda dos cosas distintas:

| Campo | Qué es | De dónde sale |
|---|---|---|
| `supervisor` | El designado en el contrato | Contrato — nombre y cédula |
| `firmante` | Quien realmente firma este periodo | **Se pregunta y se confirma cada periodo** |

Regla: preguntar una vez por periodo si firma el supervisor del contrato o alguien más. Es una
sola pregunta y evita un documento con el nombre equivocado. Si firma otro, pedir sus cuatro
datos: nombre, cédula, correo y teléfono.

El nombre del `Vo.Bo.` del informe y el de la firma de la constancia son **la misma persona**:
el `firmante`. Nunca uno del contrato y el otro preguntado.

## ARL

La clase de riesgo y el certificado son dos cosas distintas y se necesitan en momentos distintos:

- **Clase de riesgo** (`Riesgo 1`…`Riesgo 5`) → hace falta en la **fase 1**, va al Excel `I23:I27`.
- **Certificado en PDF** → hace falta en la **fase 2**, es una parte del paquete.

Puede ser **diferente por contrato**. Si `Anexos/` tiene más de un certificado, no adivines a cuál
contrato corresponde cada uno: **preguntá**. Si no hay ninguno, preguntá la clase de riesgo
contrato por contrato.

Cuando los contratos tienen riesgos distintos, el aporte se paga **por el más alto**, y hay que
decírselo al usuario porque cambia lo que debe pagar. Riesgo 4 y 5 los asume la Universidad:
el Excel deja ese aporte en `0`.

## El valor del cobro se deriva, no se pregunta

El valor que va en el punto 4 de la constancia **nunca se le pregunta al usuario**: sale del
contrato o de un número que la skill ya calculó para liquidar la planilla. Cascada, en orden:

| # | Caso | Valor del cobro |
|---|---|---|
| 1 | El contrato fija el monto de **este** pago | Ese monto, textual |
| 2 | `PAGO ÚNICO` | El valor total del contrato |
| 3 | Pagos por producto con monto | El del producto entregado en el periodo |
| 4 | Es el **pago final** y hay historial | Valor total − suma de lo ya cobrado |
| 5 | Pagos parciales sin montos | **El valor mensualizado**: `valor / meses` |

El paso 5 es el que más aplica y es gratis: `valor / meses` es exactamente `B52:B56` del Excel,
el mismo número con el que se calculó el IBC y los aportes. Si se liquidó la planilla, ese valor
**ya existe**. Preguntarlo es pedir un dato que ya está sobre la mesa.

Ejemplos reales:

```
"EN LA SIGUIENTE FORMA: PAGO ÚNICO"                    → caso 2 → el valor total
"EN LA SIGUIENTE FORMA: PAGOS PARCIALES"  (sin montos) → caso 5 → valor / meses
"UN PAGO DE $X A LA ENTREGA DEL PRODUCTO N° n"         → caso 3 → $X
```

Preguntar solo si ninguno de los cinco casos resuelve, y aun así **proponer el mensualizado** como
valor por defecto en vez de dejar la pregunta abierta.

### El número del pago

`Parcial No. ___` tampoco se pregunta: es la posición del cobro dentro del contrato. Sale de
`numero_de_mes`, el mismo que calcula el avance del informe. En contratos que pagan por producto,
es el número del producto entregado.

## Validación planilla vs. Excel — bloqueante

`total_aportes` (`I40`, sección 4 del Excel) **debe ser ≤** el total de aportes efectivamente
pagados en la planilla adjunta.

Si la planilla queda por debajo, el paquete **no se arma**: hay que corregir la planilla.
Reportar ambos valores y la diferencia.

## Periodo de la planilla

La planilla tiene dos fechas distintas y solo una debe coincidir:

- **Fecha de pago** → espacio 2 de la constancia. No tiene que coincidir con nada.
- **Periodo de cobertura** → espacio 3 de la constancia, y **debe ser igual al mes que se está
  cobrando**.

`C42` ya se escribió en la fase 1 con el mismo valor que `C41`. La fase 2 lo compara contra el
periodo de cobertura real de la planilla: si coinciden, `C42` queda igual; si difieren, se corrige
`C42` con el valor real y se avisa al usuario. Esa diferencia es justo la que hace que `E45` solo
admita `NO` y se pierda la disminución de base de retención.

## Parámetros de la vigencia

Viven en la hoja oculta `Datos` del propio libro; no se replican ni se recalculan a mano.
Para 2026: SMMLV `1.750.905`, UVT `52.374`, IBC `40 %`, salud `12,5 %`, pensión `16 %`.

- El IBC consolidado tiene piso de 1 SMMLV y techo de 25 SMMLV.
- Deducciones solo si `valor_mensualizado_total > 95 UVT` (para 2026: `4.975.530`).

## Contenido del paquete por contrato

En este orden, todo en un único PDF:

1. Constancia de cumplimiento contractual
2. Informe de ejecución de actividades
3. Planilla de seguridad social
4. Certificado de ARL **de ese contrato**
5. Contrato
6. Certificación cedular — la copia **de esa subcarpeta**, con su propio `D36`
7. **Soportes adicionales — opcional**, al final de todo

### Qué manda la parte 7

La parte 7 **no depende del tipo de pago** (parcial, final o único). Depende de una sola cosa:

> **la sección 3 del informe de ejecución — "PRODUCTOS ENTREGADOS A LA FECHA
> (En caso de haber sido pactados)".**

Esa tabla es la declaración formal de qué se entregó. La parte 7 es su evidencia física.

| Sección 3 del informe | Parte 7 del paquete |
|---|---|
| Vacía (el contrato no pactó productos) | No existe. Ni se pregunta |
| Con productos declarados | Se pregunta al usuario dónde está el soporte de cada uno |

Los soportes no se generan: los tiene el usuario. Si declara productos pero no tiene el archivo,
se deja constancia en el reporte y el paquete se arma igual — no bloquea.

No inventes la sección 3: se llena solo con los productos pactados en la cláusula `NOTA` del
contrato. Si el contrato no pactó ninguno, queda vacía y ahí se acaba el asunto.

> Ojo con los contratos que pagan **por producto entregado** ("un pago de $X a la entrega del
> producto N° n"). Ahí el `Parcial No.___` de la constancia es el número del producto, y la
> sección 3 debe declarar justamente ese.

## Contratos escaneados

Algunos contratos no tienen capa de texto (`pdftotext` devuelve vacío). Esos se leen con la
herramienta de lectura de PDF por visión, nunca se adivinan sus datos.
