import urllib.request
import urllib.parse
import json
import re

api_key = 'AIzaSyAduvvGyXxyntmoMJQR_pnczbqatXW8q0c'
folder_id_2026 = '18dbLxdOuAe8koxgPJ6IvmLpEZer4X7-e'

MES_MAP = {
    'ENERO':1, 'FEBRERO':2, 'MARZO':3, 'ABRIL':4, 'MAYO':5, 'JUNIO':6,
    'JULIO':7, 'AGOSTO':8, 'SEPTIEMBRE':9, 'OCTUBRE':10,
    'NOVIEMBRE':11, 'DICIEMBRE':12,
}
RUBROS_PRINCIPALES = ['CARNES','FIAMBRES','COMESTIBLES','BEBIDAS','VERDURAS',
                      'PANIFICACION','POLLO','LACTEOS','LIMPIEZA','KIOSCO']
RUBROS_OTROS = ['CERDO','ROTISERIA','PERFUMERIA','FRESCOS','CONGELADOS','VARIOS','DESCARTABLES']

def fetch_json(url):
    req = urllib.request.urlopen(url)
    return json.loads(req.read())

def get_files(folder_id):
    url = f"https://www.googleapis.com/drive/v3/files?q='{folder_id}'+in+parents+and+trashed=false&key={api_key}&fields=files(id,name,mimeType)"
    return fetch_json(url).get('files', [])

def parse_month_year(name):
    name_up = name.upper()
    mes_nombre = None
    mes_num = 0
    for m, n in MES_MAP.items():
        if m in name_up:
            mes_nombre = m
            mes_num = n
            break
    a_match = re.search(r'(20\d{2})', name_up)
    ano = int(a_match.group(1)) if a_match else 2026
    return mes_nombre, mes_num, ano

def extract_data():
    files = get_files(folder_id_2026)
    results = []
    
    for f in files:
        if f['mimeType'] != 'application/vnd.google-apps.spreadsheet': continue
        mes_nombre, mes_num, ano = parse_month_year(f['name'])
        if not mes_nombre: continue
        
        print(f"Procesando {f['name']}...")
        
        # Get sheet names
        meta_url = f"https://sheets.googleapis.com/v4/spreadsheets/{f['id']}?key={api_key}&fields=sheets.properties.title"
        meta = fetch_json(meta_url)
        sheet_titles = [s['properties']['title'] for s in meta.get('sheets', [])]
        
        target_eco = 'ECONOMICO' if 'ECONOMICO' in sheet_titles else 'ESTADOS' if 'ESTADOS' in sheet_titles else None
        
        if not target_eco or 'DATA STUDIO' not in sheet_titles:
            print("  Faltan hojas requeridas")
            continue
            
        ranges = "ranges=" + urllib.parse.quote("'DATA STUDIO'!A:B") + "&ranges=" + urllib.parse.quote(f"'{target_eco}'!A:B")
        data_url = f"https://sheets.googleapis.com/v4/spreadsheets/{f['id']}/values:batchGet?{ranges}&valueRenderOption=UNFORMATTED_VALUE&key={api_key}"
        
        batch_data = fetch_json(data_url).get('valueRanges', [])
        ds_rows = batch_data[0].get('values', [])
        eco_rows = batch_data[1].get('values', [])
        
        record = {
            'mes_nombre': mes_nombre,
            'mes_num': mes_num,
            'a\u00f1o': ano,
            'ventas': 0.0, 'costo_ventas': 0.0, 'gastos_empresa': 0.0, 'impuestos': 0.0,
            'servicios': 0.0, 'sueldos': 0.0, 'gastos_bancarios': 0.0,
            'gastos_mercado_pago': 0.0, 'gastos_posnet': 0.0, 'extraordinarios': 0.0,
            'compras': 0.0, 'stock_cierre': 0.0, 'saldo_proveedores': 0.0, 'saldo_clientes': 0.0,
            'resultado_economico': 0.0, 'total_egresos': 0.0, 'margen_neto': 0.0
        }
        
        DS_MAP = {
            'VENTAS': 'ventas', 'COSTO DE VENTAS': 'costo_ventas', 'GASTOS': 'gastos_empresa',
            'IMPUESTOS': 'impuestos', 'SERVICIOS': 'servicios', 'EMPLEADOS': 'sueldos',
            'GASTOS Y COMISIONES BANCARIAS': 'gastos_bancarios', 
            'GASTOS Y COMISIONES MERCADO PAGO': 'gastos_mercado_pago',
            'GASTOS Y COMISIONES POSNET': 'gastos_posnet', 'EXTRAORDINARIOS': 'extraordinarios',
            'COMPRAS': 'compras', 'STOCK AL CIERRE': 'stock_cierre',
            'SALDO PROVEEDORES': 'saldo_proveedores', 'SALDO CLIENTES': 'saldo_clientes',
            'RESULTADO ECONOMICO': 'resultado_economico'
        }
        
        # Parse DATA STUDIO
        for row in ds_rows:
            if len(row) >= 2:
                k = str(row[0]).strip().upper()
                v = row[1]
                if isinstance(v, (int, float)) and k in DS_MAP:
                    record[DS_MAP[k]] = float(v)
                    
        # Calcular total_egresos y margen_neto
        record['total_egresos'] = (record['costo_ventas'] + record['gastos_empresa'] + 
            record['impuestos'] + record['servicios'] + record['sueldos'] + 
            record['gastos_bancarios'] + record['gastos_mercado_pago'] + 
            record['gastos_posnet'] + record['extraordinarios'])
            
        if record['ventas'] > 0:
            record['margen_neto'] = (record['resultado_economico'] / record['ventas']) * 100
            
        # Parse Rubros
        for r in RUBROS_PRINCIPALES: record[f'rubro_{r.lower()}'] = 0.0
        for r in RUBROS_OTROS: record[f'rubro_otros_{r.lower()}'] = 0.0
        record['rubro_otros'] = 0.0
        
        for row in eco_rows:
            if len(row) >= 2:
                k = str(row[0]).strip().upper()
                v = row[1]
                if isinstance(v, (int, float)):
                    if k in RUBROS_PRINCIPALES:
                        record[f'rubro_{k.lower()}'] = float(v)
                    elif k in RUBROS_OTROS:
                        record[f'rubro_otros_{k.lower()}'] = float(v)
                        record['rubro_otros'] += float(v)
                        
        results.append(record)
        
    results.sort(key=lambda x: x['mes_num'])
    return results

data = extract_data()
with open('piloto_unificado.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, indent=2, ensure_ascii=False)

print(f"Exito. Se guardaron {len(data)} meses en piloto_unificado.json")
