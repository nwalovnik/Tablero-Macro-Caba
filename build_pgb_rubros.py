"""Baja PGB sectorial + IPCBA por rubro desde IDECBA, los inyecta en macro_data.json
y reemplaza el const MACRO = {...} dentro de tablero-macro.html.

Ejecutar luego de build_macro_data.py + build_calendario.py.
Pensado para el Task Scheduler de Windows o cron.
"""
import json, os, re, subprocess, sys, time
import requests
from io import BytesIO
import openpyxl
from datetime import datetime, timedelta

BASE = os.path.dirname(os.path.abspath(__file__))
# PROJ ya no es necesario, usamos BASE para todo
MACRO_JSON = os.path.join(BASE, 'macro_data.json')
HTML_FILES = [os.path.join(BASE, 'tablero-macro.html')]

H = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
     'Accept': 'application/json,text/html,*/*'}
WP_REST = 'https://www.estadisticaciudad.gob.ar/eyc/wp-json/wp/v2/banco_datos'

# ─── GET con reintentos ────────────────────────────────────────────
# El sitio de IDECBA a veces no responde por unos minutos (timeouts de
# conexión/lectura). Se reintenta con espera creciente antes de rendirse.
def get(url, timeout=60, intentos=3, **kw):
    for i in range(1, intentos + 1):
        try:
            r = requests.get(url, headers=H, timeout=timeout, **kw)
            if r.status_code >= 500:
                raise requests.HTTPError(f'HTTP {r.status_code}')
            return r
        except (requests.ConnectionError, requests.Timeout, requests.HTTPError) as e:
            if i == intentos:
                raise
            espera = 30 * i
            print(f'  intento {i}/{intentos} falló ({e.__class__.__name__}); reintento en {espera}s', flush=True)
            time.sleep(espera)

# ─── Helpers de descubrimiento de URL XLSX ─────────────────────────
def find_xlsx_for_search(query):
    """Busca un dataset por query en banco_datos y devuelve la URL del XLSX adjunto.
    Devuelve (None,None) si la API falla — el caller usa la URL fallback hardcoded."""
    try:
        r = get(WP_REST, timeout=(20, 45), params={'search': query, 'per_page': 5})
        if r.status_code != 200:
            print(f'  WP REST devolvió {r.status_code}, usando fallback', flush=True)
            return None, None
        posts = r.json()
    except Exception as e:
        print(f'  WP REST falló ({e}), usando fallback', flush=True)
        return None, None
    for post in posts:
        link = post.get('link')
        if not link: continue
        try:
            html = get(link, timeout=(20, 45)).text
        except Exception:
            continue
        m = re.search(r'href="(https://www\.estadisticaciudad\.gob\.ar/eyc/wp-content/uploads/[^"]+\.xlsx)"', html, re.I)
        if m: return m.group(1), post.get('title', {}).get('rendered', '')
    return None, None

def download_xlsx(url):
    # allow_redirects=False: si el archivo ya no existe, IDECBA redirige a la
    # portada (http) en vez de dar 404; mejor fallar rápido y usar el respaldo.
    r = get(url, timeout=(20, 90), allow_redirects=False)
    if r.status_code != 200:
        raise RuntimeError(f'{url} devolvió {r.status_code} (archivo movido o inexistente)')
    return openpyxl.load_workbook(BytesIO(r.content), data_only=True)

# ─── PGB · variación porcentual i.a. por categoría ClaNAE ─────────
def parse_pgb_variacion():
    print('[PGB var] descubriendo URL...', flush=True)
    url, title = find_xlsx_for_search('variacion porcentual producto geografico bruto trimestral')
    if not url:
        url = 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2026/03/PGB_K_variacion_porcentual.xlsx'
        title = 'Variación porcentual i.a. del PGB Trimestral por ClaNAE (fallback)'
    print(f'[PGB var] {url}', flush=True)
    wb = download_xlsx(url)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[1] or []
    sub = rows[2] or []
    year_cols = []
    for c, h in enumerate(header):
        if h is None: continue
        m = re.search(r'(20\d{2})', str(h))
        if m: year_cols.append({'year': int(m.group(1)), 'startCol': c})
    trimestres = []
    for yc in year_cols:
        for q in range(4):
            col = yc['startCol'] + q
            if col >= len(sub): continue
            s = sub[col]
            if not s: continue
            qm = re.search(r'(\d)', str(s))
            qi = int(qm.group(1)) if qm else (q+1)
            trimestres.append({'col': col, 'year': yc['year'], 'q': qi, 'label': f"{yc['year']}-T{qi}"})
    pgb_total = None
    categorias = []
    for r in rows[3:]:
        name = r[0] if r else None
        if not name: continue
        s = str(name).strip()
        if not s or s.startswith('*') or re.match(r'^Fuente', s, re.I): break
        valores = []
        for tt in trimestres:
            v = r[tt['col']] if tt['col'] < len(r) else None
            valores.append(round(float(v), 4) if isinstance(v, (int, float)) else None)
        if all(v is None for v in valores): continue
        item = {'nombre': s, 'valores': valores}
        if re.search(r'^Producto\s+Geogr', s, re.I):
            pgb_total = item
        else:
            categorias.append(item)
    if not pgb_total:
        raise RuntimeError('PGB var: no se encontró fila "Producto Geográfico Bruto"')
    last_idx = len(trimestres) - 1
    while last_idx >= 0 and pgb_total['valores'][last_idx] is None:
        last_idx -= 1
    if last_idx < 0:
        raise RuntimeError('PGB var: sin valores')
    sectores_ultimo = sorted(
        [{'nombre': c['nombre'], 'var_ia': c['valores'][last_idx]} for c in categorias if c['valores'][last_idx] is not None],
        key=lambda x: -x['var_ia']
    )
    return {
        'fuente': 'IDECBA · ' + url.split('/')[-1],
        'fuente_url': url,
        'titulo_dataset': title,
        'trimestres': trimestres,
        'pgb_total': pgb_total,
        'categorias': categorias,
        'ultimo_trim': trimestres[last_idx]['label'],
        'ultimo_var_ia': pgb_total['valores'][last_idx],
        'prev_var_ia': pgb_total['valores'][last_idx-1] if last_idx > 0 else None,
        'sectores_ultimo': sectores_ultimo,
    }

# ─── PGB · nivel (precios constantes 2004) por categoría ──────────
def parse_pgb_nivel():
    print('[PGB nivel] descubriendo URL...', flush=True)
    url, title = find_xlsx_for_search('producto geografico bruto trimestral millones pesos 2004 ClaNAE')
    if not url:
        url = 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2026/03/PGB_K_Trimestral.xlsx'
        title = 'PGB Trimestral en millones de pesos a precios de 2004 (fallback)'
    print(f'[PGB nivel] {url}', flush=True)
    wb = download_xlsx(url)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[1] or []
    sub = rows[2] or []
    year_cols = []
    for c, h in enumerate(header):
        if h is None: continue
        m = re.search(r'(20\d{2})', str(h))
        if m: year_cols.append({'year': int(m.group(1)), 'startCol': c})
    trimestres = []
    for yc in year_cols:
        for q in range(4):
            col = yc['startCol'] + q
            if col >= len(sub): continue
            s = sub[col]
            if not s: continue
            qm = re.search(r'(\d)', str(s))
            qi = int(qm.group(1)) if qm else (q+1)
            trimestres.append({'col': col, 'year': yc['year'], 'q': qi, 'label': f"{yc['year']}-T{qi}"})
    pgb_total = None
    categorias = []
    for r in rows[3:]:
        name = r[0] if r else None
        if not name: continue
        s = str(name).strip()
        if not s or s.startswith('*') or re.match(r'^Fuente', s, re.I): break
        valores = []
        for tt in trimestres:
            v = r[tt['col']] if tt['col'] < len(r) else None
            valores.append(round(float(v), 1) if isinstance(v, (int, float)) else None)
        if all(v is None for v in valores): continue
        item = {'nombre': s, 'valores': valores}
        if re.search(r'^Producto\s+Geogr', s, re.I):
            pgb_total = item
        else:
            categorias.append(item)
    return {
        'fuente': 'IDECBA · ' + url.split('/')[-1],
        'fuente_url': url,
        'trimestres': trimestres,
        'pgb_total': pgb_total,
        'categorias': categorias,
    }

# ─── IPCBA · apertura por rubro (CORREGIDO Y ROBUSTO) ──────────────
DIVISIONES = [
    'Alimentos y bebidas no alcohólicas','Bebidas alcohólicas y tabaco','Prendas de vestir y calzado',
    'Vivienda, agua, electricidad, gas y otros combustibles','Equipamiento y mantenimiento del hogar',
    'Salud','Transporte','Información y comunicación','Recreación y cultura','Educación',
    'Restaurantes y hoteles','Seguros y servicios financieros',
    'Cuidado personal, protección social y otros productos'
]

def parse_ipcba_rubros():
    print('[IPCBA rubros] descubriendo URL...', flush=True)
    url, title = find_xlsx_for_search('IPCBA aperturas indice mensual')
    if not url:
        url = 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2026/07/IPCBA_base_2021100-Principales_aperturas_indices.xlsx'
        title = 'IPCBA por aperturas (fallback)'
    print(f'[IPCBA rubros] {url}', flush=True)
    
    wb = download_xlsx(url)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    
    # 1. Encontrar dinámicamente la fila base ("Nivel General")
    ng_idx = -1
    for i, r in enumerate(rows):
        if r and r[0] and str(r[0]).strip() == 'Nivel General':
            ng_idx = i
            break
            
    if ng_idx == -1:
        raise RuntimeError("IPCBA rubros: No se encontró la fila 'Nivel General'")

    # 2. Buscar la fila de fechas mirando hacia arriba desde Nivel General
    months = []
    for offset in [1, 2, 3]:  # Probar la fila anterior, la otra, y la otra
        date_row = rows[ng_idx - offset]
        temp_months = []
        for c, v in enumerate(date_row):
            if c == 0: continue
            ym = None
            if isinstance(v, (int, float)) and v > 30000:
                try:
                    ym = (datetime(1899, 12, 30) + timedelta(days=int(v))).strftime('%Y-%m')
                except Exception: pass
            elif hasattr(v, 'strftime'):
                ym = v.strftime('%Y-%m')
            elif isinstance(v, str):
                # Extrae fechas si vienen en texto (ej: "Ene-25" o "2025-01")
                v_str = str(v).lower().strip()
                m_mes = re.search(r'(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)[^\d]*(\d{2,4})', v_str)
                m_num = re.search(r'(20\d{2})[-/](0[1-9]|1[0-2])', v_str)
                if m_mes:
                    dict_meses = {'ene':'01','feb':'02','mar':'03','abr':'04','may':'05','jun':'06','jul':'07','ago':'08','sep':'09','oct':'10','nov':'11','dic':'12'}
                    yy = m_mes.group(2)
                    if len(yy) == 2: yy = '20' + yy
                    ym = f"{yy}-{dict_meses[m_mes.group(1)]}"
                elif m_num:
                    ym = f"{m_num.group(1)}-{m_num.group(2)}"
            
            if ym:
                temp_months.append({'col': c, 'ym': ym})
        
        # Si encontró suficientes meses, esa es la fila correcta
        if len(temp_months) > 12:
            months = temp_months
            break

    if len(months) < 13:
        raise RuntimeError(f'IPCBA rubros: solo se extrajeron {len(months)} meses. El formato del Excel cambió drásticamente.')

    last_idx = len(months) - 1
    ia_idx = last_idx - 12

    # 3. Extraer los datos de Nivel General y Divisiones
    def pick(name):
        for r in rows[ng_idx:]:
            cell = r[0] if r else None
            if cell and str(cell).strip() == name:
                last = r[months[last_idx]['col']] if months[last_idx]['col'] < len(r) else None
                prev = r[months[last_idx-1]['col']] if months[last_idx-1]['col'] < len(r) else None
                ia   = r[months[ia_idx]['col']] if months[ia_idx]['col'] < len(r) else None
                
                if not isinstance(last, (int, float)): return None
                return {
                    'nombre': name,
                    'indice': round(float(last), 2),
                    'var_mensual': round((float(last)/float(prev)-1)*100, 2) if isinstance(prev, (int, float)) and float(prev)!=0 else None,
                    'var_ia':      round((float(last)/float(ia)-1)*100, 2) if isinstance(ia, (int, float)) and float(ia)!=0 else None,
                }
        return None

    ng = pick('Nivel General')
    divisiones = [pick(n) for n in DIVISIONES]
    divisiones = [d for d in divisiones if d]
    
    return {
        'fuente': 'IDECBA · ' + url.split('/')[-1],
        'fuente_url': url,
        'titulo_dataset': title,
        'periodo': months[last_idx]['ym'],
        'nivel_general': ng,
        'divisiones': divisiones,
    }

# ─── Industria · ingresos fabriles por rama (re-process for "peso") ──
def industria_pesos(macro):
    ind = macro.get('industria_ingresos', {})
    ramas = ind.get('ramas', {})
    if not ramas: return None
    last_vals = {r: (v[-1] if v and isinstance(v[-1], (int, float)) else None) for r, v in ramas.items()}
    suma = sum(v for v in last_vals.values() if v is not None)
    if not suma: return None
    pesos = {r: round(v/suma*100, 2) if v is not None else None for r, v in last_vals.items()}
    return {
        'periodo': ind.get('periodos', [None])[-1],
        'suma_ramas': round(suma, 1),
        'pesos': pesos,
    }

# ─── Patcher del HTML ──────────────────────────────────────────────
PATTERN_MACRO = re.compile(r'(const MACRO\s*=\s*)(\{.*?\})(\s*;)', re.S)

def patch_html(html_path, macro):
    if not os.path.exists(html_path):
        print(f'  (skip {html_path} — no existe)', flush=True)
        return
    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()
    inline = json.dumps(macro, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    new, n = PATTERN_MACRO.subn(lambda m: m.group(1) + inline + m.group(3), html, count=1)
    if not n:
        print(f'  WARN: no encontré const MACRO en {html_path}', flush=True)
        return
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(new)
    print(f'  patched -> {os.path.basename(html_path)} ({len(inline):,} bytes inline)', flush=True)

def macro_publicado():
    """macro_data.json tal como está en el último commit (antes de que
    build_macro_data.py lo regenere sin los bloques de este script)."""
    try:
        out = subprocess.run(['git', 'show', 'HEAD:macro_data.json'], cwd=BASE,
                             capture_output=True, check=True).stdout
        return json.loads(out.decode('utf-8'))
    except Exception as e:
        print(f'  (sin macro_data.json previo en git: {e})', flush=True)
        return {}

# ─── Main ──────────────────────────────────────────────────────────
def main():
    if not os.path.exists(MACRO_JSON):
        raise RuntimeError(f'falta {MACRO_JSON}')
    with open(MACRO_JSON, 'r', encoding='utf-8') as f:
        macro = json.load(f)

    # Último macro_data.json publicado: respaldo si IDECBA no responde
    prev = macro_publicado()

    # PGB var i.a.
    try:
        pgb_var = parse_pgb_variacion()
    except Exception as e:
        if not prev.get('pgb'):
            raise
        print(f'[PGB var] AVISO: falló ({e}); se conserva el PGB publicado ({prev["pgb"].get("ultimo_trim")})', flush=True)
        pgb_var = None
    if pgb_var is None:
        macro['pgb'] = prev['pgb']
        for k in ('actividad', '_iae_legacy'):
            if k in prev:
                macro[k] = prev[k]
        if 'iae' in macro:
            macro['_iae_legacy'] = macro.pop('iae')
    else:
        pgb_nivel = None
        try:
            pgb_nivel = parse_pgb_nivel()
        except Exception as e:
            print(f'[PGB nivel] falló (no es crítico): {e}', flush=True)
        pesos_previos = None
        if not pgb_nivel and prev.get('pgb', {}).get('pesos_ultimo'):
            pesos_previos = {
                'pesos_ultimo': prev['pgb'].get('pesos_ultimo'),
                'nivel_total_ultimo': prev['pgb'].get('nivel_total_ultimo'),
                'nivel_trim': prev['pgb'].get('nivel_trim'),
            }
            print(f'[PGB nivel] conservando pesos previos ({len(pesos_previos["pesos_ultimo"])} sectores)', flush=True)

        pgb = {
            'fuente': pgb_var['fuente'],
            'titulo_dataset': pgb_var.get('titulo_dataset'),
            'trimestres': pgb_var['trimestres'],
            'pgb_total': pgb_var['pgb_total'],
            'categorias': pgb_var['categorias'],
            'ultimo_trim': pgb_var['ultimo_trim'],
            'ultimo_var_ia': pgb_var['ultimo_var_ia'],
            'prev_var_ia': pgb_var['prev_var_ia'],
            'sectores_ultimo': pgb_var['sectores_ultimo'],
        }
        if pesos_previos:
            pgb.update(pesos_previos)
        if pgb_nivel:
            last_idx = len(pgb_nivel['trimestres']) - 1
            while last_idx >= 0 and (pgb_nivel['pgb_total'] is None or pgb_nivel['pgb_total']['valores'][last_idx] is None):
                last_idx -= 1
            if last_idx >= 0 and pgb_nivel['pgb_total']:
                total_q = pgb_nivel['pgb_total']['valores'][last_idx]
                pesos = []
                for c in pgb_nivel['categorias']:
                    v = c['valores'][last_idx] if last_idx < len(c['valores']) else None
                    if v and total_q:
                        pesos.append({'nombre': c['nombre'], 'nivel': v, 'peso': round(v/total_q*100, 2)})
                pesos.sort(key=lambda x: -x['peso'])
                pgb['pesos_ultimo'] = pesos
                pgb['nivel_total_ultimo'] = total_q
                pgb['nivel_trim'] = pgb_nivel['trimestres'][last_idx]['label']
                # Serie completa de niveles trimestrales del total, útil para calcular
                # variación trim/trim y i.a. desde el HTML sin volver a bajar XLSX.
                pgb['nivel_total_serie'] = [
                    {'trim': pgb_nivel['trimestres'][i]['label'],
                     'nivel': pgb_nivel['pgb_total']['valores'][i]}
                    for i in range(len(pgb_nivel['trimestres']))
                    if pgb_nivel['pgb_total']['valores'][i] is not None
                ]

        macro['pgb'] = pgb
        if 'iae' in macro:
            macro['_iae_legacy'] = macro.pop('iae')
        trims_lbl = [t['label'] for t in pgb['trimestres']]
        var_ia = pgb['pgb_total']['valores']
        idx = []
        base = None
        for v in var_ia:
            if v is None:
                idx.append(None); continue
            if base is None:
                base = 100.0
                idx.append(round(base, 2))
            else:
                prev_idx = next((x for x in reversed(idx) if x is not None), None)
                if prev_idx is None:
                    idx.append(round(100.0, 2))
                else:
                    idx.append(round(prev_idx * (1 + (v - (var_ia[var_ia.index(v)-1] if var_ia.index(v) > 0 and var_ia[var_ia.index(v)-1] is not None else 0))/100), 2))
        macro['actividad'] = {
            'trimestres': trims_lbl,
            'var_ia': var_ia,
            'fuente': 'PGB · IDECBA',
        }

    # IPCBA rubros
    try:
        rubros = parse_ipcba_rubros()
    except Exception as e:
        rubros = (prev.get('ipcba') or {}).get('rubros')
        if not rubros:
            raise
        print(f'[IPCBA rubros] AVISO: falló ({e}); se conservan los rubros publicados ({rubros.get("periodo")})', flush=True)
    macro.setdefault('ipcba', {})['rubros'] = rubros

    # Industria pesos
    ipesos = industria_pesos(macro)
    if ipesos:
        macro.setdefault('industria_ingresos', {})['pesos_ultimo'] = ipesos

    macro['generado_pgb_rubros'] = time.strftime('%Y-%m-%d %H:%M:%S')

    with open(MACRO_JSON, 'w', encoding='utf-8') as f:
        json.dump(macro, f, ensure_ascii=False, separators=(',', ':'))
    print(f'OK macro_data.json actualizado (PGB {len(macro["pgb"]["categorias"])} cats; IPCBA rubros {len(rubros["divisiones"])}; industria pesos {len((ipesos or {}).get("pesos") or {})})', flush=True)

    for h in HTML_FILES:
        patch_html(h, macro)

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'ERROR: {e}', file=sys.stderr)
        sys.exit(1)
