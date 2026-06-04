---
name: dashboards-context
description: Use ONLY when working on the Don Américo, Itala, or CIAF Carreras dashboards. Contains the repository structure, Excel data formats, and recent architectural changes to avoid re-analyzing the project.
---

# Contexto de Dashboards

Este repositorio contiene tres proyectos principales de dashboards Vanilla JS (HTML/CSS/JS sin frameworks pesados), los cuales leen datos financieros.

## 1. Don Américo (Supermercado)
- **Archivos:** `dashboard_don_americo.html` (legacy) y `dashboard_don_americo_json.html` (nuevo).
- **Contexto:** Se migró de leer múltiples Excels pesados directamente en el frontend a consumir un único `piloto_unificado.json`.
- **ETL:** Hay un script de Python (`test_extract.py` / `etl_consolida.py`) que se conecta a la API de Google Drive, escanea los Excels de cada mes, extrae solo las hojas "DATA STUDIO" y "ECONOMICO/ESTADOS", normaliza los rubros y calcula `total_egresos` y `margen_neto`, para escupir el JSON estructurado.
- **Reglas de trabajo:** El dashboard nuevo lee el JSON servido por http para evitar problemas de CORS. Ya no utiliza `xlsx.js` ni parsea Google Sheets en vivo en el navegador.

## 2. ITALA (Rendición Automotriz)
- **Archivos:** `dashboard_itala.html`
- **Contexto:** Datos altamente transaccionales. El Google Sheet funciona como base de datos relacional.
- **Estructura:** Utiliza hojas como `movimientos` (ingresos/egresos detallados), `saldos` (control de cajas) y `quiter` (sistema ERP). 
- **Reglas de trabajo:** Gran dependencia de catálogos de normalización (`cajas`, `mapeo_cajas`, `rubreros`) para evitar que los selectores y gráficas del dashboard se rompan por errores de tipeo en el Excel crudo.

## 3. CIAF Carreras (Eventos Deportivos)
- **Archivos:** `dashboard_carreras.html`, `KPIs_dashboard_gaps.md`
- **Contexto:** Actualmente lee un roll-up diario (`ingresos x carrera`) para mostrar únicamente la plata cobrada (Mercado Pago vs Efectivo).
- **Gaps a resolver:** Según la auditoría, sobra información en las hojas crudas por evento (`10k`, `DVL`, etc.) que no se explota. Próximas implementaciones deben agregar KPIs como: cantidad de inscriptos, ticket promedio, mix por distancia, margen neto (leyendo la hoja `egresos`), y velocidad de inscripción.

## Stack Tecnológico y Arquitectura
- Frontend puro Vanilla (HTML5, CSS, JS embebido en el mismo HTML).
- Renderizado de gráficos con `Chart.js`.
- Exportación a PDF con `html2canvas` y `jsPDF`.
- Lógica de lectura antigua (a deprecar): `xlsx.js` procesando binarios obtenidos de Drive API en el navegador del cliente.
- Lógica de lectura nueva (hacia donde vamos): `fetch` directo a JSON/CSV estáticos generados por ETLs (Python) en background.