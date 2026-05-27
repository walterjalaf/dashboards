# Estructura Google Sheets para Dashboard Itala

## Convenciones generales

- **Una pestaña por colección.** El parser lee por nombre de tab, no por índice.
- **Fila 1 = headers**, sin filas merged, sin sub-headers. Headers en `snake_case`, sin espacios al final.
- **Fechas:** formato `YYYY-MM-DD` (texto) o `Date` nativo de Sheets — nunca texto con `/` ambiguo.
- **Montos:** número positivo siempre (el signo lo da la pestaña: `egresos` resta, `ingresos` suma).
- **Caja canónica:** sólo los 9 valores oficiales (ver tab `cajas`). Para entradas heredadas/inconsistentes, hay un tab `mapeo_cajas` que las traduce.
- **Validación de datos** (`Datos → Validación`): aplicar dropdown desde tabs catálogo en columnas `caja`, `clase`, `rubro`, `subrubro`.

---

## 12 pestañas

### 1. `movimientos` (unificada — recomendada en vez de `ingresos`+`egresos`)

Una sola tabla con signo del monto explícito. Más simple de mantener que dos tabs separados.

| col | header | tipo | requerido | validación |
|---|---|---|---|---|
| A | `id` | string | sí | UUID o `YYYYMM-NNNN` |
| B | `fecha` | date | sí | |
| C | `tipo` | enum | sí | `INGRESO` / `EGRESO` |
| D | `caja` | string | sí | dropdown de `cajas` |
| E | `monto` | number | sí | positivo |
| F | `clase` | string | sí | dropdown filtrado por tipo |
| G | `rubro` | string | sí | dropdown filtrado por clase |
| H | `subrubro` | string | no | dropdown filtrado por rubro |
| I | `detalle` | string | no | texto libre |
| J | `medio_pago` | string | no | (sólo egresos) |
| K | `observaciones` | string | no | |
| L | `origen_data1` | string | no | descripción cruda banco (DATA 1 del Excel) |
| M | `origen_data2` | string | no | grupo concepto banco |
| N | `origen_data3` | string | no | concepto detallado banco |
| O | `mes` | string | derivada | `=TEXT(B2,"YYYY-MM")` |

*Si preferís mantener Egresos/Ingresos separados como en el Excel original, mantené los mismos headers pero quitá col C y poné `monto` siempre positivo.*

### 2. `saldos`

| col | header | tipo | requerido |
|---|---|---|---|
| A | `caja` | string | sí (dropdown `cajas`) |
| B | `fecha` | date | sí |
| C | `saldo` | number | sí |
| D | `tipo_saldo` | enum | sí — `INICIAL` / `DIARIO` / `FINAL` |

### 3. `quiter`

| col | header | tipo | requerido |
|---|---|---|---|
| A | `id_quiter` | string | sí (ID del ERP) |
| B | `fecha` | date | sí |
| C | `caja` | string | sí (dropdown `cajas` — ya normalizada) |
| D | `caja_quiter_raw` | string | sí (nombre original Quiter, ej `BBVA BANCO FRANCES S.A.`) |
| E | `concepto` | string | sí |
| F | `ingreso` | number | no (0 si no aplica) |
| G | `egreso` | number | no (0 si no aplica) |
| H | `referencia` | string | no |

### 4. `cajas` (catálogo canónico)

| col | header | ejemplo |
|---|---|---|
| A | `caja` | `GALICIA` |
| B | `tipo` | `BANCO` / `EFECTIVO` / `BILLETERA` |
| C | `cuenta` | `072-329722/2` |
| D | `moneda` | `ARS` |
| E | `activo` | `TRUE`/`FALSE` |
| F | `orden_display` | número (para ordenar cards) |

### 5. `mapeo_cajas` (normalización de nombres Excel/legacy)

| `nombre_origen` | `caja_canonica` | `fuente` |
|---|---|---|
| GALICIA SRL | GALICIA | excel_dic25 |
| ICBC SRL | ICBC | excel_dic25 |
| SANTANDER SRL | SANTANDER | excel_dic25 |
| MP SRL | LIQUIDACIONES MP | excel_dic25 |
| MP | LIQUIDACIONES MP | excel_dic25,ene26,feb26 |
| MP LIQUIDACIONES | LIQUIDACIONES MP | excel_feb26 |
| LIQUIDAC MP | LIQUIDACIONES MP | excel_mar26+ |
| MP LIQUIDAC | LIQUIDACIONES MP | excel_mar26+ |
| Efectivo | EFECTIVO | excel_all |
| San Juan | SAN JUAN | excel_all |
| GALICIA  | GALICIA | trailing_space |

### 6. `mapeo_quiter` (Quiter raw → caja canónica)

| `nombre_quiter` | `caja_canonica` |
|---|---|
| BBVA BANCO FRANCES S.A. | FRANCES |
| BANCO GALICIA Y BUENOS AIRES S.A.U. | GALICIA |
| GALICIA MAS | GALICIA MAS |
| BANCO SANTANDER ARGENTINA S.A. | SANTANDER |
| BANCO DE SAN JUAN S.A. | SAN JUAN |
| BANCO ICBC ARGENTINA S.A. | ICBC |
| BANCO SUPERVIELLE S.A. | SUPERVILLE |
| MERCADO PAGO | LIQUIDACIONES MP |
| CAJA EFECTIVO | EFECTIVO |

*(Completar a medida que aparezcan otros nombres en el export de Quiter.)*

### 7. `rubreros_egresos` (catálogo válido)

| `clase` | `rubro` | `subrubro` |
|---|---|---|
| COMPRAS | PROVEEDORES | FIAT CREDITO |
| COMPRAS | PROVEEDORES | FCA AUTOMOVILES |
| COMPRAS | PROVEEDORES | NORAUTO |
| COMPRAS | SERVICIOS TERCERIZADOS | POLARIZADOS |
| GASTOS DE LA EMPRESA | GASTOS POSTVENTA | GASTOS POSTVENTA |
| GASTOS DE LA EMPRESA | ALQUILERES | ALQUILERES |
| GASTOS DE LA EMPRESA | GASTOS DE GESTORIA | RIOFRIO |
| IMPUESTOS | IMPUESTO A LOS DEBITOS Y CREDITOS | GALICIA |
| IMPUESTOS | OTROS IMPUESTOS | IVA |
| SUELDOS | SUELDOS | SUELDOS |
| MOVIMIENTOS DE FONDOS | INTERNO | INTERNO |
| ... | ... | ... |

### 8. `rubreros_ingresos`

| `clase` | `rubro` | `subrubro` |
|---|---|---|
| COBROS | VENTAS | TALLER |
| COBROS | VENTAS | REPUESTO |
| COBROS | VENTAS | MERCADO LIBRE |
| COBROS | VENTAS | QR/POINT |
| COBROS | PLAN DE AHORRO | 1° CUOTA |
| COBROS | PLAN DE AHORRO | GASTOS PLAN DE AHORRO |
| COBROS | ACREDITACIONES | FISERV |
| COBROS | CREDITO CHEQUES | GALICIA |
| COBROS | A CUENTA OKM | A CUENTA OKM |
| MOVIMIENTOS DE FONDOS | INTERNO | INTERNO |
| TRANSFERENCIAS | INTERNO | INTERNO |
| DEPOSITO EFECTIVO | CAJA | INTERNO |
| DEVOLUCION | DEVOLUCION | DEVOLUCION |

### 9. `clases_egresos` (dim aux para dropdowns en cascada)
`COMPRAS`, `FINANCIACIONES`, `GASTOS DE LA EMPRESA`, `IMPUESTOS`, `IMPUESTOS Y COMISIONES BANCARIAS`, `MOVIMIENTOS DE FONDOS`, `OTROS - EXTRAORDINARIOS`, `RETIRO DE SOCIOS`, `SERVICIOS`, `SUELDOS`.

### 10. `clases_ingresos`
`COBROS`, `DEPOSITO EFECTIVO`, `DEVOLUCION`, `MOVIMIENTOS DE FONDOS`, `TRANSFERENCIAS`.

### 11. `aliases_clase` (normalización de variantes históricas)

| `alias` | `valor_canonico` | `dimension` |
|---|---|---|
| 1) COMPRAS | COMPRAS | clase_egreso |
| 2) GASTOS DE LA EMPRESA | GASTOS DE LA EMPRESA | clase_egreso |
| LIQUIDACION DE DINERO | COBROS | clase_ingreso |
| TRANSFERENCIAS  | TRANSFERENCIAS | clase_ingreso (trailing_space) |
| PAGO  | PAGO | header_egreso (trailing_space) |

### 12. `config` (key-value general)

| `key` | `value` |
|---|---|
| version_schema | 1.0 |
| fecha_corte_saldos | 2026-05-31 |
| moneda_default | ARS |
| tolerancia_quiter | 0.5 |
| zona_horaria | America/Argentina/Buenos_Aires |

---

## Dataset de normalización (pasteable)

### `cajas` — 9 filas

```tsv
caja	tipo	cuenta	moneda	activo	orden_display
GALICIA	BANCO	049-9876543	ARS	TRUE	1
GALICIA MAS	BANCO	049-9876544	ARS	TRUE	2
SANTANDER	BANCO	072-329721	ARS	TRUE	3
SAN JUAN	BANCO	05005607679	ARS	TRUE	4
FRANCES	BANCO	072-329722/2	ARS	TRUE	5
ICBC	BANCO	0810/xxxxx	ARS	TRUE	6
SUPERVILLE	BANCO	xxx	ARS	TRUE	7
LIQUIDACIONES MP	BILLETERA	mp-itala	ARS	TRUE	8
EFECTIVO	EFECTIVO	-	ARS	TRUE	9
```

### `mapeo_cajas` — 14 filas (cubre todas las variantes vistas en los 6 Excel)

```tsv
nombre_origen	caja_canonica	fuente
GALICIA	GALICIA	excel
GALICIA SRL	GALICIA	excel_dic25
GALICIA MAS	GALICIA MAS	excel
SANTANDER	SANTANDER	excel
SANTANDER SRL	SANTANDER	excel_dic25
SAN JUAN	SAN JUAN	excel
San Juan	SAN JUAN	excel
FRANCES	FRANCES	excel
ICBC	ICBC	excel
ICBC SRL	ICBC	excel_dic25
SUPERVILLE	SUPERVILLE	excel
SUPERVIELLE	SUPERVILLE	typo
EFECTIVO	EFECTIVO	excel
Efectivo	EFECTIVO	excel
MP	LIQUIDACIONES MP	excel_dic25_ene26_feb26
MP SRL	LIQUIDACIONES MP	excel_dic25
MP LIQUIDACIONES	LIQUIDACIONES MP	excel_feb26
LIQUIDAC MP	LIQUIDACIONES MP	excel_mar26_plus_egresos
MP LIQUIDAC	LIQUIDACIONES MP	excel_mar26_plus_ingresos
LIQUIDACIONES MP	LIQUIDACIONES MP	canonical
```

### `clases_egresos` — 10 filas

```tsv
clase
COMPRAS
FINANCIACIONES
GASTOS DE LA EMPRESA
IMPUESTOS
IMPUESTOS Y COMISIONES BANCARIAS
MOVIMIENTOS DE FONDOS
OTROS - EXTRAORDINARIOS
RETIRO DE SOCIOS
SERVICIOS
SUELDOS
```

### `clases_ingresos` — 5 filas

```tsv
clase
COBROS
DEPOSITO EFECTIVO
DEVOLUCION
MOVIMIENTOS DE FONDOS
TRANSFERENCIAS
```

### `rubreros_egresos` — extracto canónico (29 rubros, expandir subrubros) [TSV pasteable, copiado del DATA del Excel + ampliado con valores reales en TABLA GENERAL]

```tsv
clase	rubro	subrubro
COMPRAS	PROVEEDORES	FIAT CREDITO
COMPRAS	PROVEEDORES	FCA AUTOMOVILES
COMPRAS	PROVEEDORES	NORAUTO
COMPRAS	PROVEEDORES	EURO REPAR
COMPRAS	PROVEEDORES	RUTAS AUTOMOTORES
COMPRAS	PROVEEDORES	PROVEEDORES VARIOS
COMPRAS	PROVEEDORES	TITULOS PROPIOS
COMPRAS	PROVEEDORES	COMPRA MERCADO LIBRE
COMPRAS	SERVICIOS TERCERIZADOS	POLARIZADOS
COMPRAS	SERVICIOS TERCERIZADOS	ALARMAS
COMPRAS	SERVICIOS TERCERIZADOS	SACABOLLO
COMPRAS	SERVICIOS TERCERIZADOS	PARCHADURAS
COMPRAS	SERVICIOS TERCERIZADOS	OTROS
GASTOS DE LA EMPRESA	GASTOS POSTVENTA	GASTOS POSTVENTA
GASTOS DE LA EMPRESA	ALQUILERES	ALQUILERES
GASTOS DE LA EMPRESA	RECLAMOS	RECLAMOS
GASTOS DE LA EMPRESA	GASTOS DE ENTREGA	CORTESIA
GASTOS DE LA EMPRESA	GASTOS DE ENTREGA	RTO
GASTOS DE LA EMPRESA	GASTOS DE ENTREGA	VERIFICACION
GASTOS DE LA EMPRESA	GASTOS DE ENTREGA	LAVANDERIA
GASTOS DE LA EMPRESA	GASTOS OFICINA	ROPA DE TRABAJO
GASTOS DE LA EMPRESA	GASTOS OFICINA	MANTENIMIENTO/REPARACIONES
GASTOS DE LA EMPRESA	GASTOS OFICINA	INSUMOS
GASTOS DE LA EMPRESA	GASTOS OFICINA	SUPERMERCADO
GASTOS DE LA EMPRESA	GASTOS OFICINA	LIBRERIA
GASTOS DE LA EMPRESA	GASTOS OFICINA	AGASAJOS/OTROS
GASTOS DE LA EMPRESA	GASTOS USADOS	GASTOS USADOS
GASTOS DE LA EMPRESA	GASTOS USADOS	IMPUESTO AUTOMOTOR USADOS
GASTOS DE LA EMPRESA	VIATICOS	VIATICOS
GASTOS DE LA EMPRESA	VIATICOS	GASTOS DE ZONA
GASTOS DE LA EMPRESA	GASTOS ADMINISTRATIVOS	GASTOS ADMINISTRATIVOS
GASTOS DE LA EMPRESA	COMBUSTIBLE	COMBUSTIBLE
GASTOS DE LA EMPRESA	GASTOS DE TALLER	GASTOS DE TALLER
GASTOS DE LA EMPRESA	SISTEMA/PLATAFORMAS	QUITER
GASTOS DE LA EMPRESA	SISTEMA/PLATAFORMAS	LINKENTRY
GASTOS DE LA EMPRESA	SISTEMA/PLATAFORMAS	PILOT
GASTOS DE LA EMPRESA	SISTEMA/PLATAFORMAS	ALEPH
GASTOS DE LA EMPRESA	SISTEMA/PLATAFORMAS	SISTEMA CREDITICIO
GASTOS DE LA EMPRESA	SISTEMA/PLATAFORMAS	NOSIS
GASTOS DE LA EMPRESA	SISTEMA/PLATAFORMAS	SIOMA
GASTOS DE LA EMPRESA	GASTOS DE GESTORIA	YACANTE
GASTOS DE LA EMPRESA	GASTOS DE GESTORIA	CHIARAMONTE
GASTOS DE LA EMPRESA	GASTOS DE GESTORIA	RIOFRIO
GASTOS DE LA EMPRESA	GASTOS DE GESTORIA	MARTINEZ ELSA
GASTOS DE LA EMPRESA	GASTOS DE GESTORIA	GALLASTEGUI
GASTOS DE LA EMPRESA	GASTOS DE GESTORIA	MOYA
GASTOS DE LA EMPRESA	GASTOS DE GESTORIA	DAVILA
GASTOS DE LA EMPRESA	GASTOS DE GESTORIA	RAMIREZ
GASTOS DE LA EMPRESA	GASTOS DE GESTORIA	PEREZ
GASTOS DE LA EMPRESA	GASTOS DE VENTAS	GASTOS DE VENTAS
GASTOS DE LA EMPRESA	HONORARIOS PROFESIONALES	HONORARIOS PROFESIONALES
GASTOS DE LA EMPRESA	HONORARIOS PROFESIONALES	ABOGADO
GASTOS DE LA EMPRESA	HONORARIOS PROFESIONALES	CIAF
GASTOS DE LA EMPRESA	HONORARIOS PROFESIONALES	GABRIELA GUILLEN
GASTOS DE LA EMPRESA	HONORARIOS PROFESIONALES	HIGIENE Y SEGURIDAD
GASTOS DE LA EMPRESA	PUBLICIDAD	PUBLICIDAD
GASTOS DE LA EMPRESA	SEGUROS	SEGUROS
IMPUESTOS	IMPUESTO A LOS DEBITOS Y CREDITOS	GALICIA
IMPUESTOS	IMPUESTO A LOS DEBITOS Y CREDITOS	GALICIA MAS
IMPUESTOS	IMPUESTO A LOS DEBITOS Y CREDITOS	SANTANDER
IMPUESTOS	IMPUESTO A LOS DEBITOS Y CREDITOS	SAN JUAN
IMPUESTOS	IMPUESTO A LOS DEBITOS Y CREDITOS	FRANCES
IMPUESTOS	IMPUESTO A LOS DEBITOS Y CREDITOS	ICBC
IMPUESTOS	OTROS IMPUESTOS	IVA
IMPUESTOS	OTROS IMPUESTOS	IIBB
IMPUESTOS	OTROS IMPUESTOS	AUTONOMO
IMPUESTOS	OTROS IMPUESTOS	MONOTRIBUTO GESTORA
IMPUESTOS	OTROS IMPUESTOS	BIENES PERSONALES
IMPUESTOS	OTROS IMPUESTOS	AGENTE DE RETENCION
IMPUESTOS	OTROS IMPUESTOS	REG PROP AUTOMOTOR
IMPUESTOS	OTROS IMPUESTOS	IMPUESTO AUTOMOTOR
IMPUESTOS	OTROS IMPUESTOS	PLAN DE PAGO
IMPUESTOS	OTROS IMPUESTOS	MUNICIPALIDAD
IMPUESTOS	OTROS IMPUESTOS	F. 931
IMPUESTOS	OTROS IMPUESTOS	OBRA SOCIAL
IMPUESTOS	OTROS IMPUESTOS	SINDICATO
IMPUESTOS	OTROS IMPUESTOS	TRANSFERENCIA JUDICIAL
IMPUESTOS	OTROS IMPUESTOS	EMBARGO JUDICIAL
IMPUESTOS Y COMISIONES BANCARIAS	COMISIONES BANCARIAS	COMISIONES BANCARIAS
IMPUESTOS Y COMISIONES BANCARIAS	INTERESES POR DESCUBIERTO	INTERESES POR DESCUBIERTO
SERVICIOS	SERVICIOS	ENERGIA SAN JUAN
SERVICIOS	SERVICIOS	AGUA
SERVICIOS	SERVICIOS	GAS
SERVICIOS	SERVICIOS	INTERNET
SERVICIOS	SERVICIOS	TELEFONIA
SERVICIOS	SERVICIOS	CORREO
SERVICIOS	SERVICIOS	REVISTA
SERVICIOS	SERVICIOS	POSNETS
SERVICIOS	OTROS SERVICIOS	OTROS SERVICIOS
SUELDOS	SUELDOS	SUELDOS
SUELDOS	SUELDOS	ANTICIPOS
SUELDOS	SUELDOS	F931
RETIRO DE SOCIOS	RETIRO DE SOCIOS	RETIRO DE SOCIOS
FINANCIACIONES	TARJETA DE CREDITO	TARJETA DE CREDITO
FINANCIACIONES	PRESTAMO	PRESTAMO
MOVIMIENTOS DE FONDOS	INTERNO	GALICIA
MOVIMIENTOS DE FONDOS	INTERNO	GALICIA MAS
MOVIMIENTOS DE FONDOS	INTERNO	SANTANDER
MOVIMIENTOS DE FONDOS	INTERNO	SAN JUAN
MOVIMIENTOS DE FONDOS	INTERNO	FRANCES
MOVIMIENTOS DE FONDOS	INTERNO	ICBC
MOVIMIENTOS DE FONDOS	INTERNO	SUPERVILLE
MOVIMIENTOS DE FONDOS	INTERNO	LIQUIDACIONES MP
MOVIMIENTOS DE FONDOS	INTERNO	EFECTIVO
OTROS - EXTRAORDINARIOS	OTROS - EXTRAORDINARIOS	OTROS - EXTRAORDINARIOS
```

### `rubreros_ingresos`

```tsv
clase	rubro	subrubro
COBROS	VENTAS	QR/POINT
COBROS	VENTAS	MERCADO LIBRE
COBROS	VENTAS	TALLER
COBROS	VENTAS	REPUESTO
COBROS	PLAN DE AHORRO	1° CUOTA
COBROS	PLAN DE AHORRO	GASTOS PLAN DE AHORRO
COBROS	PLAN DE AHORRO	CANCELACION
COBROS	ACREDITACIONES	FISERV
COBROS	CREDITO CHEQUES	GALICIA
COBROS	CREDITO CHEQUES	GALICIA MAS
COBROS	CREDITO CHEQUES	FRANCES
COBROS	CREDITO CHEQUES	SAN JUAN
COBROS	CREDITO CHEQUES	SANTANDER
COBROS	A CUENTA OKM	A CUENTA OKM
COBROS	USADO	USADO
COBROS	PROVEEDORES	FIAT CREDITO
COBROS	PROVEEDORES	OTROS PROVEEDORES
COBROS	APORTE SOCIO	INTERNO
COBROS	OTROS	OTROS
TRANSFERENCIAS	INTERNO	INTERNO
MOVIMIENTOS DE FONDOS	INTERNO	INTERNO
DEPOSITO EFECTIVO	CAJA	INTERNO
DEPOSITO EFECTIVO	REINTEGRO IMP Y RECLAMOS	MERCADO PAGO
DEVOLUCION	DEVOLUCION	DEVOLUCION
```

### `aliases_clase`

```tsv
alias	valor_canonico	dimension
1) COMPRAS	COMPRAS	clase_egreso
2) GASTOS DE LA EMPRESA	GASTOS DE LA EMPRESA	clase_egreso
LIQUIDACION DE DINERO	COBROS	clase_ingreso
TRANSFERENCIAS 	TRANSFERENCIAS	clase_ingreso
PAGO 	PAGO	header
```

---

## Validación en cascada (Sheets nativo)

Para hacer dropdowns dependientes en `movimientos`:

```
F2 (clase):    rango = clases_egresos!A2:A   o clases_ingresos!A2:A según C2
G2 (rubro):    rango = FILTER(rubreros_egresos!B:B, rubreros_egresos!A:A = F2)
H2 (subrubro): rango = FILTER(rubreros_egresos!C:C, rubreros_egresos!A:A=F2, rubreros_egresos!B:B=G2)
```

Sheets nativo no soporta dropdowns dependientes con `FILTER` directo en validación de datos. Tres opciones:

1. **Apps Script** que repinte la validación de G/H cuando cambia F (más limpio, ~30 líneas).
2. **Columnas auxiliares ocultas** con `UNIQUE/FILTER` y referencia a esos rangos como named range.
3. **No validar en cascada**: dejar las 3 columnas como dropdowns planos contra `clases_egresos`, `unique(rubreros_egresos.B)`, `unique(rubreros_egresos.C)`. Más permisivo pero zero código.

---

## Acceso desde el parser

Si vas por **publicar como CSV** (rápido):
```
https://docs.google.com/spreadsheets/d/{ID}/gviz/tq?tqx=out:csv&sheet={tab_name}
```

Si vas por **API oficial** (recomendado para tokens y privacidad):
- Google Sheets API v4
- Una request por tab, o `batchGet` con array de rangos `cajas!A:Z`, `movimientos!A:Z`, etc.
