"""Parsea los XLSX de IDECBA descargados y genera macro_data.json consolidado."""
import json, os, openpyxl, xlrd, requests
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(BASE, 'idecba')
OUT = os.path.join(BASE, 'macro_data.json')

def descargar_ipcba():
    if not os.path.exists(XLSX):
        os.makedirs(XLSX)
    url_api = "https://www.estadisticaciudad.gob.ar/eyc/wp-json/wp/v2/banco_datos?search=ipcba_mensual"
    try:
        r = requests.get(url_api).json()
        url_xlsx = r[0]['link_archivo']
        print(f"Descargando IPCBA desde: {url_xlsx}")
        res = requests.get(url_xlsx)
        with open(os.path.join(XLSX, 'ipcba_mensual.xlsx'), 'wb') as f:
            f.write(res.content)
    except Exception as e:
        print(f"Error descargando IPCBA: {e}")

def month_label(dt):
    meses = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
    return f"{meses[dt.month-1]}-{str(dt.year)[2:]}"

def parse_year_cell(v):
    if v is None: return None
    if isinstance(v,(int,float)):
        y = int(v)
        return y if 1900 < y < 2100 else None
    s = str(v).strip().rstrip('*').strip()
    if s.endswith('.0'): s = s[:-2]
    if s.isdigit():
        y = int(s)
        return y if 1900 < y < 2100 else None
    return None

def parse_ipcba():
    path = os.path.join(XLSX,'ipcba_mensual.xlsx')
    if not os.path.exists(path):
        return {"meses":[], "nivel_general":[], "var_mensual":[]}
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    meses, nivel, var_m = [], [], []
    for row in ws.iter_rows(min_row=2):
        y = parse_year_cell(row[0].value)
        m = row[1].value
        if y and isinstance(m, int) and 1 <= m <= 12:
            dt = datetime(y, m, 1)
            meses.append(month_label(dt))
            nivel.append(row[2].value)
            var_m.append(row[3].value)
    return {"meses": meses, "nivel_general": nivel, "var_mensual": var_m}

# (Omitimos el resto de funciones de parseo por brevedad, pero asegúrate de que el script termine con esto:)

if __name__ == "__main__":
    descargar_ipcba() # Primero descargamos el archivo
    data_ipcba = parse_ipcba()
    
    # Estructura mínima para que no falle el resto del proceso
    out = {
        'ipcba': data_ipcba,
        'iae': {'trimestres':[], 'indice':[]},
        'canastas': {'meses':[], 'ca':[], 'total':[]},
        'empleo': {'trimestres':[], 'desocupacion':[]},
        'locales': {'periodos':[], 'ejes':{}},
        'fuente': 'IDECBA · GCBA',
        'generado': datetime.now().strftime('%Y-%m-%d'),
    }
    
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',',':'))
    print(f'Wrote {OUT}')
