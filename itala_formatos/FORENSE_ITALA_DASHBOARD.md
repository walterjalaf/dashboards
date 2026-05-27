# Documento Maestro Forense — Excels Itala (Dashboard Rendición y Finanzas)

**Fecha de análisis:** 2026-05-26
**Archivos analizados:** 6 (DIC/25, ENE/26, FEB/26, MAR/26, ABR/26, MAY/26)
**Tipo:** `.xlsx` (sin macros VBA)
**Objetivo:** Mapear la estructura real para alimentar las 7 colecciones del dashboard (`saldos`, `ingresos`, `egresos`, `quiter`, `rubreros_ingresos`, `rubreros_egresos`, `quiter_mapeo`).

---

## 1. Resumen Ejecutivo

Los 6 archivos son **reportes mensuales** del flujo de caja de Itala SA (concesionaria automotor en San Juan). Cada archivo tiene entre 15 y 21 hojas, ~117k fórmulas, y consolida movimientos bancarios de ~9 cajas en dos tablas-pivote: `TABLA GENERAL - Egresos` y `TABLA GENERAL - Ingresos`. **No tienen macros**.

**Complejidad: alta.** Tres problemas serios para el dashboard:

1. **El schema de las TABLAS GENERAL evolucionó tres veces** (DIC/ENE → FEB → MAR+). El parser debe manejar tres versiones.
2. **Los nombres de caja son inconsistentes** entre meses y entre Egresos/Ingresos (ej: `'MP'` → `'LIQUIDAC MP'` / `'MP LIQUIDAC'` / `'MP SRL'`).
3. **Quiter no existe como fuente de datos.** No hay hojas `BANCOS QUITER` ni `MP QUITER` ni mapeo `quiter_mapeo` en ningún archivo. La vista `rQuiter` del dashboard quedará vacía hasta que se aporten esas fuentes.

Tampoco hay columna SALDO en San Juan ni FRANCES — la vista `rSaldos` queda incompleta para esas dos cajas si solo se leen las hojas-banco.

### Métricas

| Métrica | Valor |
|---|---|
| Archivos | 6 |
| Hojas totales (rango) | 15 – 21 |
| Hojas ocultas | 0 – 3 (`Borrador - Tomás`, `ECO-FERO 03.26`, `PARA ENVIAR - INGRESOS`) |
| Fórmulas (MAY) | 117.794 (240 simples / 105.347 intermedias / 5.429 avanzadas / 6.778 críticas) |
| Named ranges | 0 |
| Macros VBA | 0 |
| Cajas distintas | 9 (con variaciones de nombre) |
| Movs Egresos/mes (rango) | 794 (MAY) – 1.186 (ENE) |
| Movs Ingresos/mes (rango) | 360 (MAY) – 669 (FEB) |

---

## 2. Inventario de Hojas (consolidado entre meses)

| Hoja | Presente en | Rol | Cargar para |
|---|---|---|---|
| `GALICIA` | todos | Hoja-banco: extracto bruto + clasificado | `saldos`, fuente de Egresos/Ingresos |
| `GALICIA MAS` | ENE → MAY | Hoja-banco | `saldos`, fuente |
| `SANTANDER` | todos | Hoja-banco | `saldos`, fuente |
| `San Juan` | todos | Hoja-banco (sin SALDO) | fuente |
| `FRANCES` | todos | Hoja-banco (sin SALDO) | fuente |
| `ICBC` | todos | Hoja-banco | `saldos`, fuente |
| `SUPERVILLE` | todos | Hoja-banco | `saldos`, fuente |
| `Efectivo` | todos | Caja efectivo | `saldos`, fuente |
| `MP` | DIC, ENE, FEB | Mercado Pago (hasta FEB) | fuente MP |
| `COBROS MP` | sólo FEB | Transitorio | (ignorar) |
| `MP LIQUIDACIONES - PRUEBA` | sólo FEB | Transitorio | (ignorar) |
| `LIQUIDACIONES MP` | MAR → MAY | MP definitivo (dump bruto de la API) | fuente MP |
| `Borrador - Tomás` | FEB → MAY (hidden desde MAR) | Mismo schema que `LIQUIDACIONES MP` — copia bruta | (ignorar) |
| **`TABLA GENERAL - Egresos`** | **todos** | **Consolidado de egresos — fuente principal** | **`egresos`** |
| **`TABLA GENERAL - Ingresos`** | **todos** | **Consolidado de ingresos — fuente principal** | **`ingresos`** |
| `DATA` | todos | Catálogo de clases/rubros/subrubros + lista cajas | `rubreros_ingresos`, `rubreros_egresos` |
| `PARA ENVIAR - EGRESOS` | todos | Reporte impreso (no es tabla) | (ignorar) |
| `PARA ENVIAR - INGRESOS` | todos (oculta en ABR) | Reporte impreso | (ignorar) |
| `PARA ENVIAR -BALANCE SEMANAL` | FEB+ | Reporte impreso | (ignorar) |
| `ACREDITACIONES` | todos | Tabla pequeña actividad → banco (7 filas) | mapeo opcional |
| `PRESTAMO` | MAR+ | Cuotas pendientes | (ignorar para dashboard) |
| `RESUMEN` | ENE+ | Balance diario (datos mayormente 0) | (ignorar) |
| `Hoja2` | sólo DIC | Vacía | (ignorar) |
| `ECO-FERO 03.26` | MAR (visible), ABR/MAY (hidden) | Legacy de un cliente puntual | (ignorar) |

---

## 3. Hojas-banco (estructura interna)

Cada hoja-banco tiene **múltiples tablas pegadas horizontalmente** en la misma hoja:

| Hoja | Tabla bruta | Tabla limpia | Tabla Egresos | Tabla Ingresos | Col SALDO |
|---|---|---|---|---|---|
| `GALICIA` | A–P (16 cols, schema banco) | R–Y (`=IF(...)` espejo) | AA–AD+ | (combinada) | P, Y |
| `GALICIA MAS` | A–P | R–Y (espejo) | AA–AF | AH–AM | P, Y |
| `SANTANDER` | J–M (4 cols) | — | O–Q | S–U | M |
| `San Juan` | H–K (4 cols) | — | M–P (Débito) | R–U (Crédito) | **ninguna** |
| `FRANCES` | L–O (4 cols) | — | R–T (Débito) | V–X (Crédito) | **ninguna** (saldo en B4 como texto) |
| `Efectivo` | A–F (6 cols) | — | H–K (Débitos) | M–P (Créditos) | F |
| `SUPERVILLE` | A–F | I–P (limpio con saldo) | R–W (Débitos) | Y–AD (Créditos) | F, P |
| `ICBC` | A–J (10 cols) | L–P (limpio con saldo) | R–T (Débitos) | V–X (Créditos) | P |
| `LIQUIDACIONES MP` | A–AS (73 cols, dump API MP) | — | — | — | AS |

**Estrategia recomendada:** **NO** leer las hojas-banco para Egresos/Ingresos. Usar las `TABLA GENERAL` (ya consolidan y clasifican todo). Las hojas-banco sirven sólo para `saldos` (cuando tienen columna SALDO).

### Esquema "tabla bruta" (cols A–P de GALICIA / GALICIA MAS)

| Col | Header | Tipo |
|---|---|---|
| A | Fecha | datetime |
| B | Descripción | string |
| C | Origen | string |
| D | Débitos | number |
| E | Créditos | number |
| F | Grupo de Conceptos | string (ej "000907 - Transferencias") |
| G | Concepto | string (ej "907269 - TRF INMED PROVEED") |
| H | Número de Terminal | number |
| I | Observaciones Cliente | string |
| J | Número de Comprobante | number |
| K-N | Leyendas Adicionales 1-4 | string |
| O | Tipo de Movimiento | string ("Imputado") |
| P | **Saldo** | number |

### LIQUIDACIONES MP / Borrador - Tomás (dump bruto de Mercado Pago API)

73 columnas: `FECHA DE LIBERACIÓN` (timestamp ISO), `ID DE OPERACIÓN`, `TIPO DE REGISTRO`, `DESCRIPCIÓN`, `MONTO NETO ACREDITADO`, `MONTO NETO DEBITADO`, `MONTO BRUTO`, `COMISIONES`, `MEDIO DE PAGO`, `SUCURSAL` ("Fiat"), `SALDO` (col AS), etc. **Importante:** `MONTO BRUTO` viene con decimales irregulares y a veces se marca como fecha por error (los warnings de openpyxl son inocuos).

---

## 4. TABLA GENERAL — fuente principal del dashboard

### 4.1 Schema TABLA GENERAL - Egresos (evolución)

| Mes | Cols | Cambios |
|---|---|---|
| DIC/25 | 7 | `BANCO, _, FECHA, DATA 1, DATA 2, DATA 3, TOTAL, _, RUBRO` (sin DETALLE, sin SUBRUBRO, sin EGRESOS-clase, sin PAGO, sin OBSERVAC) |
| ENE/26 | 8 | + `DETALLE` (col H) |
| FEB/26 | 10 | + `SUBRUBRO` (col J) |
| MAR – MAY/26 | **13** | + `FECHA-` (col B), `EGRESOS` (clase, col I), `PAGO` (col L), `OBSERVAC` (col M). **Schema completo.** |

**Schema MAR+ definitivo:**

| # | Col | Header | Tipo | Descripción |
|---|---|---|---|---|
| 1 | A | `BANCO` | string | Caja (`GALICIA`, `GALICIA MAS`, `LIQUIDAC MP`, `EFECTIVO`, etc.) |
| 2 | B | `FECHA-` | datetime | Fecha original del extracto |
| 3 | C | `FECHA` | datetime | Fecha normalizada (igual a B en general) |
| 4 | D | `DATA 1` | string | Descripción del movimiento |
| 5 | E | `DATA 2` | string | Grupo de conceptos del banco |
| 6 | F | `DATA 3` | string | Concepto detallado del banco |
| 7 | G | `TOTAL` | number | **Monto egreso** (positivo) |
| 8 | H | `DETALLE` | string | Beneficiario / descripción manual |
| 9 | I | `EGRESOS` | string | **Clase** (`COMPRAS`, `GASTOS DE LA EMPRESA`, `IMPUESTOS`, etc.) |
| 10 | J | `RUBRO` | string | **Rubro** (`PROVEEDORES`, `GASTOS POSTVENTA`, etc.) |
| 11 | K | `SUBRUBRO` | string | **Subrubro** (`FIAT CREDITO`, `RIOFRIO`, etc.) |
| 12 | L | `PAGO ` (con espacio) | string | Forma de pago |
| 13 | M | `OBSERVAC` | string | Observaciones libres |

### 4.2 Schema TABLA GENERAL - Ingresos (evolución)

| Mes | Cols | Cambios |
|---|---|---|
| DIC/25 | 8 | `BANCO, _, FECHA, DATA 1-3, Leyenda Adicional, TOTAL` |
| ENE/26 | 10 | + `DETALLE`, `RUBRO` |
| FEB/26 | 10 | igual |
| MAR – MAY/26 | **13** | + `INGRESOS` (clase, col J), `SUBRUBRO` (K), `ORIGEN` (M) |

**Schema MAR+ definitivo:**

| # | Col | Header | Tipo |
|---|---|---|---|
| 1 | A | `BANCO` | string |
| 2 | B | `FECHA` | datetime |
| 3 | C | `F` | datetime (duplicado) |
| 4 | D | `DATA 1` | string |
| 5 | E | `DATA 2` | string |
| 6 | F | `DATA 3` | string |
| 7 | G | `Leyenda Adicional` | string |
| 8 | H | `TOTAL` | number — **monto ingreso** |
| 9 | I | `DETALLE` | string |
| 10 | J | `INGRESOS` | string — **clase** |
| 11 | K | `RUBRO` | string |
| 12 | L | `SUBRUBRO` | string |
| 13 | M | `ORIGEN` | string |

### 4.3 Cajas en `BANCO` (col A) — **inconsistencias entre meses**

| Caja canónica del dashboard | DIC/25 | ENE/26 | FEB/26 | MAR – MAY/26 |
|---|---|---|---|---|
| GALICIA | `GALICIA SRL` | `GALICIA` | `GALICIA` | `GALICIA` |
| GALICIA MAS | — | `GALICIA MAS` | `GALICIA MAS` | `GALICIA MAS` |
| SANTANDER | `SANTANDER SRL` | `SANTANDER` | `SANTANDER` | `SANTANDER` |
| SAN JUAN | `SAN JUAN` | `SAN JUAN` | `SAN JUAN` | `SAN JUAN` |
| FRANCES | `FRANCES` | `FRANCES` | `FRANCES` | `FRANCES` |
| ICBC | `ICBC SRL` | `ICBC` | `ICBC` | `ICBC` |
| SUPERVILLE | `SUPERVILLE` | `SUPERVILLE` | `SUPERVILLE` | `SUPERVILLE` |
| EFECTIVO | `EFECTIVO` | `EFECTIVO` | `EFECTIVO` | `EFECTIVO` |
| LIQUIDACIONES MP (Egresos) | `MP SRL` | `MP` | `MP` | `LIQUIDAC MP` |
| LIQUIDACIONES MP (Ingresos) | `MP SRL` | `MP` | `MP` + `MP LIQUIDACIONES` | **`MP LIQUIDAC`** |

**Acción obligatoria** en el parser: normalizar `BANCO` con un diccionario antes de cargar.

```python
NORMALIZAR_CAJA = {
    'GALICIA SRL': 'GALICIA',
    'ICBC SRL': 'ICBC',
    'SANTANDER SRL': 'SANTANDER',
    'MP SRL': 'LIQUIDACIONES MP',
    'MP': 'LIQUIDACIONES MP',
    'LIQUIDAC MP': 'LIQUIDACIONES MP',
    'MP LIQUIDAC': 'LIQUIDACIONES MP',
    'MP LIQUIDACIONES': 'LIQUIDACIONES MP',
    'EFECTIVO': 'EFECTIVO',
    # el resto pasa por sí mismo
}
```

### 4.4 Volúmenes por archivo

| Archivo | Egresos | Ingresos |
|---|---:|---:|
| DIC/25 | 916 | 429 |
| ENE/26 | 1.186 | 573 |
| FEB/26 | 1.134 | 669 |
| MAR/26 | 1.074 | 641 |
| ABR/26 | 934 | 506 |
| MAY/26 | 794 | 360 |
| **Total** | **6.038** | **3.178** |

---

## 5. Hoja DATA — catálogos

Layout de la hoja DATA (147 filas × 29 cols, **sparse**):

| Col | Header (fila 1) | Contenido | Uso para dashboard |
|---|---|---|---|
| A | (vacío) | Lista de **cajas válidas**: `GALICIA, SUPERVILLE, SAN JUAN, SANTANDER, FRANCES, ICBC, MP, MP LIQUIDAC` (8 cajas — **no incluye GALICIA MAS ni EFECTIVO**, advertir al cargar) | lista de cajas (extender) |
| C | `RUBROS` | Lista de rubros de egresos (37 items: ARCA, COBROS, COMISIONES BANCARIAS Y DE MP, COMPRAS, …) | — (legacy, usar Q/R) |
| D | `TOTALES` | Totales calculados por rubro (fórmulas SUMIFS) | reportes |
| H | `EGRESOS POR RUBRO` | Igual a col C (espejo) | — |
| I (fila 1) | `2026-02-12 00:00:00` | Fecha snapshot (header dinámico) | — |
| L | `SUBRUBROS` | Lista plana de 50+ subrubros | — |
| M | `TOTALES` | Totales por subrubro | reportes |
| **Q** | **`EGRESOS`** | **Clase (`1) COMPRAS`, `2) GASTOS DE LA EMPRESA`, …)** | **`rubreros_egresos.clase`** |
| **R** | **`RUBRO`** | **Rubro válido para esa clase** | **`rubreros_egresos.rubro`** |
| **S** | **`SUBRUBRO`** | **Subrubro válido para ese rubro** | **`rubreros_egresos.subrubro`** |
| V | `EGRESOS` | Clase variante (sin numeración) | — (redundante con Q) |
| W | `RUBRO` | Rubro variante | — |
| X | `SUBRUBRO` | Subrubro variante | — |
| **Z** | **`INGRESOS`** | **Clase ingreso (`LIQUIDACION DE DINERO`, `COBROS`, `TRANSFERENCIAS`, …)** | **`rubreros_ingresos.clase`** |
| **AA** | **`RUBRO`** | | **`rubreros_ingresos.rubro`** |
| **AB** | **`SUBRUBRO`** | | **`rubreros_ingresos.subrubro`** |

### 5.1 Catálogo de Egresos (valores reales en TABLA GENERAL)

**Clases (10):** `COMPRAS`, `FINANCIACIONES`, `GASTOS DE LA EMPRESA`, `IMPUESTOS`, `IMPUESTOS Y COMISIONES BANCARIAS`, `MOVIMIENTOS DE FONDOS`, `OTROS - EXTRAORDINARIOS`, `RETIRO DE SOCIOS`, `SERVICIOS`, `SUELDOS`.

**Rubros (29):** `COMBUSTIBLE, COMISIONES BANCARIAS, GASTOS ADMINISTRATIVOS, GASTOS DE ENTREGA, GASTOS DE GESTORIA, GASTOS DE TALLER, GASTOS DE VENTAS, GASTOS OFICINA, GASTOS POSTVENTA, GASTOS USADOS, HONORARIOS PROFESIONALES, IMPUESTO A LOS DEBITOS Y CREDITOS, IMPUESTO AUTOMOTOR, IMPUESTOS, INTERESES POR DESCUBIERTO, INTERNO, OTROS - EXTRAORDINARIOS, OTROS IMPUESTOS, OTROS SERVICIOS, PROVEEDORES, PUBLICIDAD, RETIRO DE SOCIOS, SEGUROS, SERVICIOS, SERVICIOS TERCERIZADOS, SISTEMA/PLATAFORMAS, SUELDOS, TARJETA DE CREDITO, VIATICOS`.

**Subrubros (64+):** abogado, autonomo, CIAF, FIAT CREDITO, NORAUTO, etc.

### 5.2 Catálogo de Ingresos

**Clases (5):** `COBROS, DEPOSITO EFECTIVO, DEVOLUCION, MOVIMIENTOS DE FONDOS, TRANSFERENCIAS`.

**Rubros (10):** `A CUENTA OKM, ACREDITACIONES, APORTE SOCIO, CAJA, CREDITO CHEQUES, DEVOLUCION, INTERNO, PLAN DE AHORRO, PROVEEDORES, VENTAS`.

**Subrubros (13):** `1° CUOTA, A CUENTA OKM, CANCELACION, DEVOLUCION, FIAT CREDITO, FISERV, GALICIA, GALICIA MAS, GASTOS PLAN DE AHORRO, INTERNO, MERCADO LIBRE, REPUESTO, TALLER`.

### 5.3 Advertencias del catálogo

- En la hoja DATA hay **dos variantes del catálogo de egresos** (cols Q-S y V-X) que no coinciden 100%. Usar el de Q-S (`1) COMPRAS`, `2) GASTOS DE LA EMPRESA`, etc.) y normalizar el prefijo `N) `.
- El catálogo Egresos cita rubros como `OTROS - EXTRAORDINARIOS` que también figura como **clase** en TABLA GENERAL. Hay solapamiento entre niveles.
- Algunos subrubros en TABLA GENERAL (ej `ABOGADO`, `CIAF`, `EMBARGO JUDICIAL`, `RAMIREZ`) **no figuran** en DATA → activan la barra ⚠ del dashboard.

---

## 6. Saldos

Las hojas-banco son la única fuente de saldos diarios. **Estructura inconsistente**:

| Caja | Columna saldo en hoja-banco | Disponible? |
|---|---|---|
| GALICIA | P (bruto) y Y (espejo) | ✓ |
| GALICIA MAS | P y Y | ✓ |
| SANTANDER | M | ✓ |
| ICBC | F (col bruta) y P (espejo) | ✓ |
| SUPERVILLE | F y P | ✓ |
| Efectivo | F | ✓ |
| LIQUIDACIONES MP | AS | ✓ |
| **San Juan** | (no hay header SALDO) | **✗** |
| **FRANCES** | (saldo de cabecera en B4 como texto: `'-105.403.752,44'`) | **✗ (parseable como string)** |

**Estrategia:** para cada hoja-banco, leer pares `(FECHA, SALDO)` de las columnas indicadas. Para San Juan y FRANCES: derivar saldo como `saldo_inicial + cumsum(crédito - débito)` por fecha, o reportar `null` en `rSaldos`.

### "Saldo inicial"
Ninguna hoja tiene fila explícita `saldo inic`. El dashboard espera registros tipo `{caja, fecha, saldo: "saldo inic"}` — habrá que **construirlos sintéticamente** tomando el primer saldo del mes anterior al cierre.

---

## 7. Quiter — **FALTANTE**

El dashboard espera 3 estructuras Quiter:

- `quiter` — `{caja, fecha, ingreso, egreso, concepto}` desde hojas `BANCOS QUITER` + `MP QUITER`
- `quiter_mapeo` — `{nombre_banco_quiter → caja_excel}`

**Ninguna de las tres existe en ningún archivo.** Búsqueda exhaustiva:

- No hay hojas con "QUITER" en el nombre.
- "Quiter" aparece sólo como **subrubro** en DATA (cols L25, S35, X37) — es decir, gasto del sistema, no fuente de datos.
- "Borrador - Tomás" (hidden) parece "fuente Quiter" pero es un mirror de `LIQUIDACIONES MP` (mismo header de 73 columnas).

**Conclusión:** la vista `rQuiter` del dashboard quedará vacía. Hay tres opciones:

1. **Esconder la vista** hasta que Itala aporte el export de Quiter.
2. **Mockear** con un loader vacío (`quiter = []`) y mostrar todos los Δ como `Excel - 0`.
3. **Pedir** los archivos `.xlsx` o `.csv` de Quiter (exportar manualmente desde el ERP) y definir un schema canónico antes de seguir.

---

## 8. Mapa de Fórmulas (MAY/26 como representativo)

| Complejidad | Cantidad | % |
|---|---:|---:|
| Simples (SUM, ROUND, IF directo) | 240 | 0.2 % |
| Intermedias (IFERROR, VLOOKUP, SUMIFS) | 105.347 | 89.4 % |
| Avanzadas (IFs anidados, INDEX/MATCH) | 5.429 | 4.6 % |
| Críticas (INDIRECT, refs externas, >3 niveles) | 6.778 | 5.8 % |

**Patrones principales:**

- **Espejo limpio** en hojas-banco: `=IF(A2="","",A2)` repetido miles de veces. No es "lógica" — es presentación.
- **Lookups de clasificación** en `TABLA GENERAL`: típicamente `IFERROR(VLOOKUP(...), "")` y SUMIFS por banco/clase/rubro.
- **Fórmulas críticas concentradas en**: `PARA ENVIAR -BALANCE SEMANAL` (referencias a múltiples hojas vía INDIRECT con nombres dinámicos) y `RESUMEN`.

**Para el parser del dashboard estas fórmulas son irrelevantes** — basta con leer valores (`data_only=True`).

---

## 9. Hojas ocultas

| Hoja | Estado | Contenido | Propósito |
|---|---|---|---|
| `Borrador - Tomás` | hidden (desde MAR) | Dump bruto MP (73 cols, mismo schema que LIQUIDACIONES MP) | Borrador de trabajo, no usar |
| `ECO-FERO 03.26` | hidden (ABR, MAY); visible MAR | Hoja chica (136 filas × 10 cols), referencia a un cliente "Eco-Fero" en marzo | Legacy de un caso puntual |
| `PARA ENVIAR - INGRESOS` | hidden (sólo en ABR) | Reporte impreso (suelto, dejado oculto) | Ignorar |

Ninguna afecta la integridad de las TABLAS GENERAL.

---

## 10. Mapeo final: Excel → Colecciones del dashboard

| Colección dashboard | Fuente Excel | Notas |
|---|---|---|
| **`saldos`** `{caja, fecha, saldo}` | Hojas-banco (col SALDO según tabla §6) | Faltan San Juan y FRANCES; sintetizar "saldo inic" |
| **`ingresos`** `{caja, fecha, monto, detalle, clase, rubro, subrubro}` | `TABLA GENERAL - Ingresos` cols A, B, H, I, J, K, L | Mapear: caja=A, fecha=B, monto=H, detalle=I, clase=J, rubro=K, subrubro=L. Normalizar caja con dict §4.3. Manejar schemas viejos (DIC/ENE) con defaults. |
| **`egresos`** `{caja, fecha, monto, detalle, clase, rubro, subrubro}` | `TABLA GENERAL - Egresos` cols A, B, G, H, I, J, K | Mapear: caja=A, fecha=B, monto=G, detalle=H, clase=I, rubro=J, subrubro=K. Idem normalización. |
| **`quiter`** | **NO EXISTE** | Devolver `[]`, ocultar vista o pedir fuente |
| **`rubreros_ingresos`** `{clase, rubro, subrubro}` | `DATA` cols Z, AA, AB (filas 2 hasta vacío) | |
| **`rubreros_egresos`** `{clase, rubro, subrubro}` | `DATA` cols Q, R, S (filas 2 hasta vacío) | Quitar prefijo `N) ` de la clase |
| **`quiter_mapeo`** | **NO EXISTE** | Devolver `{}` |

---

## 11. Hallazgos y Riesgos

### 🔴 Bloqueantes para el dashboard

1. **Quiter no tiene fuente.** Decisión de producto necesaria antes de implementar `rQuiter`.
2. **Schema de TABLA GENERAL evolucionó.** El parser debe detectar la versión por presencia/ausencia de columnas y mapear por header (no por índice fijo). Si se carga DIC/25 con un parser pensado para MAY/26, las clases `EGRESOS`/`INGRESOS` quedarán como `null`.
3. **Nombres de caja inconsistentes** entre meses y entre Egresos vs Ingresos. Sin normalizar, los filtros de la vista `Caja` mostrarán duplicados (`MP`, `LIQUIDAC MP`, `MP LIQUIDAC`).

### 🟡 Importantes

4. **San Juan y FRANCES** no tienen columna SALDO normalizada — la vista `rSaldos` quedará incompleta.
5. **El catálogo de DATA tiene huecos**: faltan `GALICIA MAS` y `EFECTIVO` en la lista de cajas (col A); además ~10% de los subrubros que aparecen en TABLA GENERAL no están en DATA → mucha barra ⚠.
6. **LIQUIDACIONES MP** trae numerosos warnings de "celda marcada como fecha pero valor fuera de límites" en `MONTO BRUTO` (cols BC, BI). Son strings de identificadores tratados como fechas por error de tipado de Excel. **No afectan los datos relevantes** (`MONTO NETO ACREDITADO`, `MONTO NETO DEBITADO` están en F y G y son numéricos limpios).

### 🟢 Observaciones

7. **0 macros VBA, 0 named ranges, 0 referencias externas a otros archivos.** Modelo cerrado, autocontenido.
8. La columna `PAGO ` en TABLA GENERAL - Egresos tiene un espacio al final del header. Cuidado con `df['PAGO']` vs `df['PAGO ']` en pandas.
9. Las dos columnas de fecha en TABLA GENERAL (`FECHA-` / `FECHA` en egresos, `FECHA` / `F` en ingresos) son idénticas en la práctica. Usar la primera.

---

## 12. Recomendaciones concretas para el parser del dashboard

```python
# pseudo-código del loader
SCHEMA_EGRESOS_V1 = ['BANCO','_','FECHA','DATA 1','DATA 2','DATA 3','TOTAL','_','RUBRO']         # DIC/25
SCHEMA_EGRESOS_V2 = [..., 'DETALLE','RUBRO','SUBRUBRO']                                          # FEB/26
SCHEMA_EGRESOS_V3 = ['BANCO','FECHA-','FECHA','DATA 1-3','TOTAL','DETALLE','EGRESOS','RUBRO',
                     'SUBRUBRO','PAGO ','OBSERVAC']                                              # MAR+/26 ← canónico

def detect_version(headers):
    if 'EGRESOS' in headers and 'PAGO ' in headers: return 'v3'
    if 'SUBRUBRO' in headers: return 'v2'
    return 'v1'

def normalizar_caja(banco):
    return NORMALIZAR_CAJA.get(banco.strip().upper(), banco.strip())

def cargar_egresos(wb):
    ws = wb['TABLA GENERAL - Egresos']
    headers = [c.value for c in ws[1]]
    v = detect_version(headers)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if v == 'v3':
            caja, _, fecha, _, _, _, monto, detalle, clase, rubro, subrubro, *_ = row
        elif v == 'v2':
            caja, _, fecha, _, _, _, monto, detalle, rubro, subrubro = row[:10]
            clase = None
        else:
            caja, _, fecha, _, _, _, monto, _, rubro = row[:9]
            detalle = clase = subrubro = None
        if not caja or not monto: continue
        yield {
            'caja': normalizar_caja(caja),
            'fecha': fecha,
            'monto': float(monto),
            'detalle': detalle,
            'clase': clase,
            'rubro': rubro,
            'subrubro': subrubro,
        }
```

### Prioridades sugeridas

1. **Implementar parser por header (no por índice)**, con detección de versión.
2. **Diccionario de normalización de cajas** antes de la carga.
3. **Vista Quiter en estado "vacía/pendiente"** hasta tener fuente.
4. **Saldo sintético** para San Juan y FRANCES.
5. **Validación silenciosa**: cargar también `Borrador - Tomás` y `LIQUIDACIONES MP` brutos y compararlos con lo que llega a `TABLA GENERAL - Egresos` para detectar si hay movimientos MP no clasificados.
