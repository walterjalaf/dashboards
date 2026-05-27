# Estructura de la Google Sheet "Resumen_Mensual_DonAmerico_2026"

## Una sola hoja, una sola tabla. 38 columnas, una fila por mes.

### Cómo armarla

1. Crear Google Sheet nueva. Nombre sugerido: `Resumen_Mensual_DonAmerico_2026`.
2. Renombrar `Hoja 1` → `datos`.
3. Pegar el contenido de `dataset_normalizado.csv` directo (File → Import → Replace current sheet, o `Ctrl+V` con la primera celda en A1).
4. Verificar que las celdas numéricas tienen formato Number (Google las detecta sola, pero confirmar).
5. **Publicar como CSV:**
   - `File → Share → Publish to the web`
   - Pestaña activa: `datos`
   - Formato: `Comma-separated values (.csv)`
   - Marcar "Automatically republish when changes are made"
   - Copiar la URL pública (queda en formato `https://docs.google.com/spreadsheets/d/e/<key>/pub?output=csv`)
6. Pegar esa URL en el HTML del dashboard, en la constante `SHEET_CSV_URL`.

---

## Esquema de columnas

### Metadata (3 columnas)

| Columna | Tipo | Ejemplo | Descripción |
|---|---|---|---|
| `mes_nombre` | string | `ENERO` | Nombre del mes en mayúsculas |
| `mes_num` | int | `1` | Número del mes (1-12) |
| `año` | int | `2026` | Año |

### KPIs principales (15 columnas)

| Columna | Tipo | Origen en Excel | Descripción |
|---|---|---|---|
| `ventas` | number | `DATA STUDIO!B2` | Ingresos totales del mes |
| `costo_ventas` | number | `DATA STUDIO!B3` | Costo de reposición |
| `gastos_empresa` | number | `DATA STUDIO!B4` | Gastos generales |
| `impuestos` | number | `DATA STUDIO!B5` | Impuestos del período |
| `servicios` | number | `DATA STUDIO!B6` | Luz, gas, alquiler, etc. |
| `sueldos` | number | `DATA STUDIO!B7` | Empleados |
| `gastos_bancarios` | number | `DATA STUDIO!B9` | Comisiones bancarias |
| `gastos_mercado_pago` | number | `DATA STUDIO!B10` | Comisiones MP |
| `gastos_posnet` | number | `DATA STUDIO!B11` | Comisiones posnet |
| `extraordinarios` | number | `DATA STUDIO!B12` | Gastos no recurrentes |
| `total_egresos` | number | Σ de las 8 anteriores | **Calculado, no leer de Excel** |
| `resultado_economico` | number | `DATA STUDIO!B13` | = ventas − total_egresos |
| `margen_neto` | number | calculado | = resultado / ventas × 100 |
| `compras` | number | `DATA STUDIO!B17` | Compra de mercadería del mes |
| `stock_cierre` | number | `DATA STUDIO!B16` | Inventario al cierre |

### Rubros principales (10 columnas)

Buscar por rótulo en col A de `ESTADOS` (ENE/FEB) o `ECONOMICO` (MAR/ABR), tomar valor de col B.

| Columna | Rótulo a buscar |
|---|---|
| `rubro_carnes` | `CARNES` |
| `rubro_fiambres` | `FIAMBRES` |
| `rubro_comestibles` | `COMESTIBLES` |
| `rubro_bebidas` | `BEBIDAS` |
| `rubro_verduras` | `VERDURAS` |
| `rubro_panificacion` | `PANIFICACION` |
| `rubro_pollo` | `POLLO` |
| `rubro_lacteos` | `LACTEOS` |
| `rubro_limpieza` | `LIMPIEZA` |
| `rubro_kiosco` | `KIOSCO` |

### Rubro "Otros" (1 columna agregada + 7 de desglose)

| Columna | Descripción |
|---|---|
| `rubro_otros` | Σ de los 7 siguientes |
| `rubro_otros_cerdo` | individual |
| `rubro_otros_rotiseria` | individual |
| `rubro_otros_perfumeria` | individual |
| `rubro_otros_frescos` | individual |
| `rubro_otros_congelados` | individual |
| `rubro_otros_varios` | individual |
| `rubro_otros_descartables` | individual (solo aparece desde marzo) |

### Saldos patrimoniales (opcional — 2 columnas)

| Columna | Origen |
|---|---|
| `saldo_clientes` | `DATA STUDIO!B14` (con huecos en los datos actuales) |
| `saldo_proveedores` | `DATA STUDIO!B15` (con huecos en los datos actuales) |

> Estas dos están en el CSV adjunto pero el dashboard del usuario no las consume. Quedan disponibles si en el futuro se agrega una vista patrimonial.

---

## Reglas de actualización (para meses futuros)

Cuando llegue el archivo de MAYO 2026:

1. Ejecutar el script `etl_consolida.py` (incluido en el repo) apuntando a la carpeta con los 5 archivos. El script regenera el CSV completo.
2. Reemplazar el contenido de la pestaña `datos` con el nuevo CSV (`Ctrl+A` → `Delete` → `Ctrl+V` con la nueva data).
3. La opción "Automatically republish when changes are made" garantiza que el dashboard ya está actualizado.

No es necesario tocar el HTML.

---

## Validaciones recomendadas dentro de Google Sheet

Pegar estas fórmulas en una hoja auxiliar `_validaciones` para que cualquier inconsistencia salte rápido:

```
=IF(ROUND(B2-(C2+D2+E2+F2+G2+I2+J2+K2+L2),2)=0, "OK", "DIFF: "&(B2-(C2+D2+E2+F2+G2+I2+J2+K2+L2)))
```
(verifica que ventas − Σ egresos individuales ≈ resultado_economico)

```
=IF(B2>0, ROUND(O2/B2*100,2)=ROUND(N2,2), "MARGEN MAL")
```
(verifica margen_neto)

```
=IF(SUM(Q2:Z2)+AA2>0, ROUND(SUM(Q2:Z2)+AA2,0)=ROUND(B2,0), "RUBROS NO SUMAN VENTAS")
```
(verifica que la suma de rubros ≈ ventas totales)

> Estos checks ya pasaron localmente en los 4 archivos (ver `00_FORENSE_DON_AMERICO_2026.md` §3). Replicarlos en la sheet evita que un mes nuevo entre roto.
