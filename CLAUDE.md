# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

A collection of standalone financial dashboards and ETL scripts for three clients:

- **Itala** — rendición de ingresos/egresos mensuales (archivos Excel por mes)
- **Don Américo** — dashboard de finanzas de supermercado con rubros (fuente Google Sheets o Excel local)
- **CIAF Carreras** — dashboard de inscripciones a carreras de running (fuente Google Sheets)

Plus a shared Python tool (`normalizar_bancos.py`) that normalizes bank statements from multiple Argentine banks into a single Excel.

## Running things

**Bank normalizer (GUI):**
```
python normalizar_bancos.py
```
Opens a tkinter dialog sequence — one per bank. Produces `Reportes_banco/reporte banco normalizado.xlsx`.

**Don Américo ETL (local Excels → CSV):**
```
python formatos/etl_consolida.py <carpeta_excels> <salida.csv>
```
Input files must match `RESUMEN_MENSUAL_*.xlsx`. Detects sheet format automatically (ESTADOS = legacy, ECONOMICO = new).

**Don Américo test extractor (Google Sheets API → JSON):**
```
python test_extract.py
```
Writes `piloto_unificado.json`. Has a hardcoded API key and folder ID.

**Dashboards:** open the `.html` file directly in a browser — no server needed.

## Dashboard architecture

All dashboards are **self-contained single-file HTML** with no build step. The pattern is always:

1. CDN imports at the top: SheetJS (`xlsx.full.min.js`), Chart.js, html2canvas, jsPDF
2. All CSS in a `<style>` block using CSS variables defined in `:root`
3. All JS in a `<script>` block at the bottom — no external modules

**Data flow:** user uploads an Excel file (or the file is fetched from Google Drive/Sheets API) → parsed client-side with SheetJS → transformed in JS → rendered with Chart.js and HTML tables.

## Design system

All dashboards share the CIAF dark-navy design system. The canonical tokens are:

```css
--ink:#0E1F33        /* page background */
--surface:#132239    /* card/panel background */
--surface-2:#0a1829  /* deeper surface, pre-header */
--brand:#1E4A82      /* primary blue */
--brand-2:#2D6CB8    /* accent blue (active borders, highlights) */
--pearl:#E8EAEC      /* primary text */
--accent:#F26522     /* orange warning/highlight */
--positive:#2ECC71   /* green for income/positive values */
--negative:#E74C3C   /* red for expenses/negative values */
```

Font stack: **Montserrat** (headings, labels, uppercase tags) · **Open Sans** (body) · **JetBrains Mono** (monetary values, numbers).

Reuse existing CSS classes rather than adding new ones. Key classes: `.kpi-card`, `.kpi-grid`, `.banner`, `.sum-tbl`, `.chart-wrap`, `.dt` (detail table), `.caja-hdr`, `.tab`/`.tab.active`, `.ms-wrap` (multi-select dropdown).

## Itala data format

Source files: `itala_excels/REPORTE EGRESOS-INGRESOS ITALA [MES]-[YY].xlsx` and `itala_excels_extracted/` (processed copies).

The dashboard reads these files client-side. Month names in filenames are in Spanish uppercase (ENERO, FEBRERO, etc.).

## Don Américo data format

Monthly Excels have three key sheets:
- **DATA STUDIO** (col A = label, col B = value) — stable layer for KPI extraction
- **ECONOMICO** or **ESTADOS** (legacy) — rubros breakdown, section bounded by `CLASIF DE VENTAS POR RUBROS` and `CLASIF DE VENTAS POR FORMAS DE COBROS`
- Sheet format changed from ESTADOS (Jan–Feb) to ECONOMICO (Mar onward)

`total_egresos` is always **calculated**, never read directly. `resultado_economico` is read and cross-validated.

## Bank normalizer architecture

`normalizar_bancos.py` follows a consistent two-function pattern per bank:
- `_load_<bank>(filepath)` → raw DataFrame (handles encoding, header detection, fake-XLS quirks)
- `_normalize_<bank>(df)` → tuple `(df_egresos, df_ingresos)` with canonical columns: `BANCO, FECHA, DATA 1, DATA 2, TOTAL` + type-specific columns (`EGRESOS/RUBRO/SUBRUBRO/PAGO/OBSERVAC` or `INGRESOS/RUBRO/SUBRUBRO/ORIGEN`)

To add a new bank: implement both functions and add an entry to the `BANCOS` list. Number parsing uses `_parse_arg_number` (Argentine format: `(1.234,56)` = −1234.56).

## Key files

| File | Purpose |
|---|---|
| `dashboard_itala.html` | Itala rendición dashboard |
| `dashboard_don_americo_json.html` | Don Américo dashboard reading from `piloto_unificado.json` |
| `dashboard_carreras.html` | CIAF running races dashboard |
| `normalizar_bancos.py` | Multi-bank statement normalizer (tkinter GUI) |
| `formatos/etl_consolida.py` | Don Américo local ETL |
| `test_extract.py` | Don Américo Google Sheets extractor |
| `ciaf-dashboard-design-system.skill` | Design system reference (zip with tokens.css, template.html, render-functions.js) |
| `KPIs_dashboard_gaps.md` | Audit of current carreras dashboard gaps and proposed KPI set |
