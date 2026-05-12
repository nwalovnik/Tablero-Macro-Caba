"""Parsea los XLSX de IDECBA descargados y genera macro_data.json consolidado."""
import json, os, openpyxl, xlrd
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(BASE, 'idecba')
# CORRECCIÓN: Se quitó el '..' para que guarde en la misma carpeta en GitHub
OUT = os.path.join(BASE, 'macro_data.json') 

def month_label(dt):
    meses = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
    return f"{meses[dt.month-1]}-{str(dt.year)[2:]}"

def parse_year_cell(v):
    """Return int year from cells like 2024, '2024', '2024*', 2024.0, '2024.0'. None if not a year."""
    if v is None: return None
    if isinstance(v,(int,float)):
        y = int(v)
        return y if 1900 < y < 2100 else None
    s = str(v).strip().rstrip('*').strip()
    # remove trailing .0
    if s.endswith('.0'): s = s[:-2]
    if s.isdigit():
        y = int(s)
        return y if 1900 < y < 2100 else None
    return None

# ── IPCBA ────────────────────────────────────────────────
def parse_ipcba():
    wb = openpyxl.load_workbook(os.path.join(XLSX,'ipcba_evol.xlsx'), data_only=True)
    ws = wb['Evol_gral_estac_reg_resto']
    rows = list(ws.iter_rows(values_only=True))
    meses, niv, var_m = [], [], []
    
    for row in rows[4:]:
        mes = row[0]
        if mes is None: continue # Ignora filas vacías en lugar de frenar
        
        # Intentar convertir el mes a fecha si viene como texto o número
        dt_mes = None
        if isinstance(mes, datetime):
            dt_mes = mes
        elif isinstance(mes, str):
            # Intenta detectar "Abr-26", "2026-04", etc.
            import re
            m = re.search(r'(\d{2,4})', mes)
            if m: dt_mes = datetime.strptime(mes.strip()[:7], '%Y-%m') if '-' in mes else None
        
        if dt_mes is None: continue # Si no es fecha, sigue a la siguiente fila
        
        niv_gen = row[1]
        if not isinstance(niv_gen,(int,float)): continue
        
        meses.append(dt_mes.strftime('%Y-%m'))
        niv.append(round(float(niv_gen),2))
        v = row[5] # Variación mensual
        var_m.append(round(float(v),1) if isinstance(v,(int,float)) else None)
    
    # Cálculo de interanual
    var_ia = []
    for i,n in enumerate(niv):
        if i>=12 and niv[i-12]:
            var_ia.append(round((n/niv[i-12]-1)*100,1))
        else:
            var_ia.append(None)
            
    return {
        'meses': meses,
        'meses_label': [month_label(datetime.strptime(m,'%Y-%m')) for m in meses],
        'nivel_general': niv,
        'var_mensual': var_m,
        'var_ia': var_ia,
    }

# ── IAE ────────────────────────────────────────────────
def parse_iae():
    wb = openpyxl.load_workbook(os.path.join(XLSX,'iae.xlsx'), data_only=True)
    ws = wb['PGB_ITAE_b12']
    rows = list(ws.iter_rows(values_only=True))
    trims, idx, var_ia = [], [], []
    cur_year = None
    for row in rows[2:]:
        per = row[0]
        val = row[1]
        va = row[2]
        if per is None: continue
        y = parse_year_cell(per)
        if y is not None:
            cur_year = y; continue
        if cur_year is None: continue
        per_s = str(per).strip()
        m = None
        if '1er' in per_s: m='T1'
        elif '2do' in per_s: m='T2'
        elif '3er' in per_s: m='T3'
        elif '4to' in per_s: m='T4'
        if m is None or not isinstance(val,(int,float)): continue
        trims.append(f"{cur_year}-{m}")
        idx.append(round(float(val),1))
        var_ia.append(round(float(va),1) if isinstance(va,(int,float)) else None)
    return {'trimestres':trims, 'indice':idx, 'var_ia':var_ia}

# ── Canastas ───────────────────────────────────────────
def parse_canastas():
    wb = openpyxl.load_workbook(os.path.join(XLSX,'canastas.xlsx'), data_only=True)
    ws = wb['Canasta_cons_hogar1']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[2]
    row_ca = row_caysh = row_total = None
    for r in rows[3:]:
        name = r[0]
        if not name: continue
        n = str(name).lower()
        if 'canasta alimentaria (ca)' in n: row_ca = r
        elif 'alimentaria y de servicios' in n: row_caysh = r
        elif n.strip() == 'canasta total': row_total = r
    meses, ca, caysh, total = [], [], [], []
    def _get(r,i):
        if r is None: return None
        v = r[i] if i < len(r) else None
        return round(float(v),2) if isinstance(v,(int,float)) else None
    for i,val in enumerate(header[1:],1):
        if isinstance(val, datetime):
            meses.append(val.strftime('%Y-%m'))
            ca.append(_get(row_ca,i))
            caysh.append(_get(row_caysh,i))
            total.append(_get(row_total,i))
    return {'meses':meses, 'ca':ca, 'caysh':caysh, 'total':total}

# ── Empleo ─────────────────────────────────────────────
def parse_empleo():
    wb = openpyxl.load_workbook(os.path.join(XLSX,'empleo.xlsx'), data_only=True)
    ws = wb['ETOI_O_TG1']
    rows = list(ws.iter_rows(values_only=True))
    def _f(v):
        if isinstance(v,(int,float)): return round(float(v),1)
        if isinstance(v,str):
            s = v.replace(',','.').strip()
            # strip trailing letters like '5.8a'
            num = ''
            for ch in s:
                if ch.isdigit() or ch=='.' or ch=='-': num += ch
                elif num: break
            try: return round(float(num),1) if num else None
            except: return None
        return None
    trims, act, emp, des, sub = [], [], [], [], []
    cur_year = None
    for row in rows[2:]:
        per = row[0]
        if per is None: continue
        y = parse_year_cell(per)
        if y is not None:
            cur_year = y; continue
        if cur_year is None: continue
        per_s = str(per).strip()
        m = None
        if '1er' in per_s: m='T1'
        elif '2do' in per_s: m='T2'
        elif '3er' in per_s: m='T3'
        elif '4to' in per_s: m='T4'
        if m is None: continue
        a = _f(row[1])
        if a is None: continue
        trims.append(f"{cur_year}-{m}")
        act.append(a)
        emp.append(_f(row[2]))
        des.append(_f(row[3]))
        sub.append(_f(row[4]))
    return {'trimestres':trims, 'actividad':act, 'empleo':emp, 'desocupacion':des, 'subocupacion':sub}

# ── Locales ────────────────────────────────────────────
def _sheet_to_periodo(name):
    """'3er. cuatr. de 2025' → '2025-C3'"""
    name = name.strip()
    y_m = None
    import re
    m = re.search(r'(\d{4})', name)
    if m: y_m = m.group(1)
    if '1er' in name: q = 'C1'
    elif '2do' in name: q = 'C2'
    elif '3er' in name: q = 'C3'
    else: return None
    return f'{y_m}-{q}' if y_m else None

def parse_locales():
    # Archivo con conteos absolutos + tasas + var. i.a. (AC_EJ_2022_02)
    fname_abs = os.path.join(XLSX, 'locales_abs.xlsx')
    wb = openpyxl.load_workbook(fname_abs, data_only=True)

    # Identificar hojas de periodos (excluir la consolidada y ficha técnica)
    sheet_periods = []
    for sh in wb.sheetnames:
        p = _sheet_to_periodo(sh)
        if p: sheet_periods.append((p, sh))
    # Ordenar cronológicamente
    sheet_periods.sort(key=lambda x: x[0])

    periodos = [p for p, _ in sheet_periods]

    # Estructuras por eje
    ejes_tasa    = {}   # tasa ocupacion % (col 8) — backward compat
    ejes_relev   = {}   # locales relevados (col 1)
    ejes_ocup    = {}   # locales ocupados (col 3)
    ejes_desoc   = {}   # total desocupados = relevados - ocupados
    ejes_var_ia  = {}   # var interanual pp (col 10)

    for per, sh in sheet_periods:
        ws = wb[sh]
        rows = list(ws.iter_rows(values_only=True))
        # Fila de datos comienza en row index 3 (0-based)
        data_start = 3
        for row in rows[data_start:]:
            name = row[0]
            if not name: continue
            name = str(name).strip()
            if not name: continue

            def _f(col): return round(float(row[col]), 1) if len(row)>col and isinstance(row[col],(int,float)) else None
            def _i(col): return int(row[col]) if len(row)>col and isinstance(row[col],(int,float)) else None

            relev  = _i(1)
            ocup   = _i(3)
            tasa   = _f(8)
            var_ia = _f(10)
            desoc  = (relev - ocup) if (relev is not None and ocup is not None) else None

            for d, v in [(ejes_tasa, tasa), (ejes_relev, relev), (ejes_ocup, ocup),
                         (ejes_desoc, desoc), (ejes_var_ia, var_ia)]:
                d.setdefault(name, []).append(v)

    return {
        'periodos':  periodos,
        'ejes':      ejes_tasa,    # tasa ocupación % por eje (series)
        'relevados': ejes_relev,   # conteo total de locales relevados
        'ocupados':  ejes_ocup,    # conteo de ocupados
        'desocupados': ejes_desoc, # conteo total desocupados
        'var_ia':    ejes_var_ia,  # variación i.a. en pp de tasa de ocupación
    }

# ── Industria (ingresos y personal) ──────────────────────
def parse_industria(kind='ingresos'):
    fname = 'industria_ing.xlsx' if kind=='ingresos' else 'industria_pers.xlsx'
    sheet = 'ee_industria_ingresos_fabriles' if kind=='ingresos' else 'ee_ind_personal_asalariado'
    wb = openpyxl.load_workbook(os.path.join(XLSX,fname), data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    # row 2: 'Valores corrientes' en cols 2-11, 'Valores constantes' en cols 12-21
    # row 3 headers: col2=Total(corr), cols 3-11=9 ramas(corr), col12=Total(const), cols13-21=9 ramas(const)
    ramas = [str(v).strip() if v else '' for v in rows[2][3:12]]
    ramas = [r for r in ramas if r]
    n_ramas = len(ramas)
    # columna de inicio de valores constantes (Total constantes = col 12, ramas const = 13..21)
    col_const_total = 12
    col_const_ramas = 13
    periodos, total, total_const = [], [], []
    por_rama = {r:[] for r in ramas}
    por_rama_const = {r:[] for r in ramas}
    cur_year = None
    for row in rows[3:]:
        y = parse_year_cell(row[0])
        if y is not None: cur_year = y
        m = row[1]
        if not m or cur_year is None: continue
        if not isinstance(row[2],(int,float)): continue
        mes = str(m).strip()
        mes_num = {'Enero':'01','Febrero':'02','Marzo':'03','Abril':'04','Mayo':'05','Junio':'06','Julio':'07','Agosto':'08','Septiembre':'09','Octubre':'10','Noviembre':'11','Diciembre':'12'}.get(mes)
        if not mes_num: continue
        periodos.append(f"{cur_year}-{mes_num}")
        total.append(round(float(row[2]),1))
        vc = row[col_const_total] if len(row) > col_const_total else None
        total_const.append(round(float(vc),1) if isinstance(vc,(int,float)) else None)
        for i,r in enumerate(ramas):
            v = row[3+i] if len(row)>3+i else None
            por_rama[r].append(round(float(v),1) if isinstance(v,(int,float)) else None)
            vc2 = row[col_const_ramas+i] if len(row)>col_const_ramas+i else None
            por_rama_const[r].append(round(float(vc2),1) if isinstance(vc2,(int,float)) else None)
    return {
        'periodos': periodos,
        'total': total,
        'total_constantes': total_const,
        'ramas': por_rama,
        'ramas_constantes': por_rama_const,
    }

# ── Pobreza e indigencia (CV_AX15) ─────────────────────
def parse_pobreza_tasas():
    """
    CV_AX15.xlsx — una hoja por año (2015-2025).
    Cada hoja tiene 4 bloques de 8 columnas (T1=col1, T2=col9, T3=col17, T4=col25).
    Dentro de cada bloque: Hogares%, nota, Hogares_abs, nota, Personas%, nota, Personas_abs, nota
    Fila 6 (0-based) = pobreza, Fila 7 = indigencia.
    """
    wb = openpyxl.load_workbook(os.path.join(XLSX, 'pobreza_tasas.xlsx'), data_only=True)
    periodos = []
    pob_hog_pct, pob_hog_abs = [], []
    pob_per_pct, pob_per_abs = [], []
    ind_hog_pct, ind_hog_abs = [], []
    ind_per_pct, ind_per_abs = [], []

    def _f(v):
        if isinstance(v, (int, float)): return round(float(v), 2)
        return None

    def _i(v):
        if isinstance(v, (int, float)): return int(v)
        return None

    year_sheets = []
    for sh in wb.sheetnames:
        try:
            y = int(sh.strip())
            if 2000 < y < 2100:
                year_sheets.append((y, sh))
        except:
            pass
    year_sheets.sort()

    def _nums(row):
        """Extrae valores numéricos de la fila en orden, ignorando texto y separadores.
        Heurística: % de hogares siempre <100, abs siempre >=1000.
        Retorna lista en orden: T1[H%,H_abs,P%,P_abs], T2[...], T3[...], T4[...]"""
        return [v for v in row[1:] if isinstance(v, (int, float))]

    for year, sh in year_sheets:
        rows = list(wb[sh].iter_rows(values_only=True))
        row_pob = row_ind = None
        for row in rows:
            cell0 = str(row[0]).lower() if row[0] else ''
            if 'pobreza' in cell0 and row_pob is None:
                row_pob = row
            elif 'indigencia' in cell0 and row_ind is None:
                row_ind = row
            if row_pob and row_ind: break
        if row_pob is None or row_ind is None:
            continue

        nums_pob = _nums(row_pob)
        nums_ind = _nums(row_ind)
        # Esperamos 16 valores por fila (4 trimestres × 4 valores: H%, H_abs, P%, P_abs)
        for t in range(4):
            base = t * 4
            if base + 3 >= len(nums_pob):
                continue
            ph_pct, ph_abs, pp_pct, pp_abs = nums_pob[base:base+4]
            # Validación sanity: % debe ser <100, abs debe ser >=1000
            if ph_pct >= 100 or pp_pct >= 100:
                continue
            periodos.append(f"{year}-T{t+1}")
            pob_hog_pct.append(round(ph_pct, 2))
            pob_hog_abs.append(int(ph_abs))
            pob_per_pct.append(round(pp_pct, 2))
            pob_per_abs.append(int(pp_abs))
            if base + 3 < len(nums_ind):
                ih_pct, ih_abs, ip_pct, ip_abs = nums_ind[base:base+4]
                ind_hog_pct.append(round(ih_pct, 2) if ih_pct < 100 else None)
                ind_hog_abs.append(int(ih_abs))
                ind_per_pct.append(round(ip_pct, 2) if ip_pct < 100 else None)
                ind_per_abs.append(int(ip_abs))
            else:
                ind_hog_pct.append(None); ind_hog_abs.append(None)
                ind_per_pct.append(None); ind_per_abs.append(None)

    return {
        'periodos': periodos,
        'pob_hog_pct': pob_hog_pct,
        'pob_hog_abs': pob_hog_abs,
        'pob_per_pct': pob_per_pct,
        'pob_per_abs': pob_per_abs,
        'ind_hog_pct': ind_hog_pct,
        'ind_hog_abs': ind_hog_abs,
        'ind_per_pct': ind_per_pct,
        'ind_per_abs': ind_per_abs,
    }

# ── Autoservicios mayoristas (consumo masivo) ──────────
MESES_ES = {
    'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
    'julio':7,'agosto':8,'septiembre':9,'octubre':10,'noviembre':11,'diciembre':12,
}

def _norm(s):
    """Normalize string for comparison (lower, no accents, strip)."""
    if s is None: return ''
    s = str(s).strip().lower()
    repl = {'á':'a','é':'e','í':'i','ó':'o','ú':'u','ñ':'n'}
    for k,v in repl.items(): s = s.replace(k,v)
    return s

def parse_autoservicios():
    """AC_M_01.xlsx — variación interanual % de ventas a valores constantes.
    Col 0: alterna 'YYYY' (fila resumen anual) y nombres de mes.
    Col 1: var. i.a. %.
    """
    wb = openpyxl.load_workbook(os.path.join(XLSX, 'autoservicios.xlsx'), data_only=True)
    ws = wb['AC_M_01']
    rows = list(ws.iter_rows(values_only=True))
    periodos, var_ia = [], []
    cur_year = None
    for row in rows[2:]:   # filas 0-1 son títulos/header
        if row[0] is None:
            continue
        cell = str(row[0]).strip()
        # Año
        try:
            y = int(float(cell.replace('*','').strip()))
            if 2000 <= y <= 2100:
                cur_year = y
                continue
        except Exception:
            pass
        # Mes
        m = MESES_ES.get(_norm(cell).replace('*','').strip())
        if m is None or cur_year is None:
            continue
        v = row[1]
        try:
            val = round(float(v), 2) if v is not None else None
        except Exception:
            val = None
        periodos.append(f'{cur_year}-{m:02d}')
        var_ia.append(val)
    return {'periodos': periodos, 'var_ia': var_ia,
            'fuente': 'IDECBA · AC_M_01.xlsx'}


# ── Shoppings (centros de compras, ventas por rubro) ───
def parse_shoppings():
    """AC_CC_AX07.xlsx — variación i.a. % de ventas constantes por rubro.
    Fila 2: header ('Período','Rubro').
    Fila 3: nombres de rubros en columnas 1..N.
    Filas 4+: alterna 'YYYY' y meses, valores por rubro en cada columna.
    """
    wb = openpyxl.load_workbook(os.path.join(XLSX, 'shoppings.xlsx'), data_only=True)
    ws = wb['AC_CC_AX07']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[3]
    # Detectar columnas con nombre de rubro (no None y no nota al pie)
    rubros_idx = []
    for i, name in enumerate(header[1:], 1):
        if name is None: break
        s = str(name).strip()
        if not s or s.startswith('.') or len(s) < 3: break
        rubros_idx.append((i, s))
    rubros = {name: [] for _, name in rubros_idx}
    periodos = []
    cur_year = None
    for row in rows[4:]:
        if row[0] is None:
            continue
        cell = str(row[0]).strip()
        # Año
        try:
            y = int(float(cell.replace('*','').strip()))
            if 2000 <= y <= 2100:
                cur_year = y
                continue
        except Exception:
            pass
        m = MESES_ES.get(_norm(cell).replace('*','').strip())
        if m is None or cur_year is None:
            # Llegamos a notas al pie / texto
            break
        periodos.append(f'{cur_year}-{m:02d}')
        for col_i, name in rubros_idx:
            v = row[col_i] if col_i < len(row) else None
            try:
                rubros[name].append(round(float(v), 2) if v is not None else None)
            except Exception:
                rubros[name].append(None)
    return {'periodos': periodos, 'rubros': rubros,
            'rubros_orden': [name for _, name in rubros_idx],
            'fuente': 'IDECBA · AC_CC_AX07.xlsx'}


# ── Supermercados (retail, var i.a. de ventas constantes) ──
def parse_supermercados():
    """AC_S_01.xlsx — variación interanual % de ventas a valores constantes.
    Col 0: alterna 'YYYY' y nombres de mes.
    Col 1: IPI Encuesta (serie larga, 2001+).
    Col 2: IP elaborado IDECBA con IPCBA (más reciente, 2013+).
    Tomamos col 1 (cobertura completa) y si es '.' o None, fallback a col 2.
    """
    wb = openpyxl.load_workbook(os.path.join(XLSX, 'supermercados.xlsx'), data_only=True)
    ws = wb['AC_S_01']
    rows = list(ws.iter_rows(values_only=True))
    periodos, var_ia = [], []
    cur_year = None
    for row in rows[4:]:
        if row[0] is None:
            continue
        cell = str(row[0]).strip()
        try:
            y = int(float(cell.replace('*','').strip()))
            if 2000 <= y <= 2100:
                cur_year = y
                continue
        except Exception:
            pass
        m = MESES_ES.get(_norm(cell).replace('*','').strip())
        if m is None or cur_year is None:
            break
        # Preferir col 1 (IPI Encuesta), fallback col 2
        v = None
        for col in (1, 2):
            if col < len(row) and row[col] is not None and row[col] != '.':
                try:
                    v = round(float(row[col]), 2)
                    break
                except Exception:
                    pass
        periodos.append(f'{cur_year}-{m:02d}')
        var_ia.append(v)
    return {'periodos': periodos, 'var_ia': var_ia,
            'fuente': 'IDECBA · AC_S_01.xlsx'}


# ── Masa salarial industrial (Total nominal índice base oct-2001=100) ──
def parse_masa_salarial():
    """ee_industria_masa_salarial.xlsx — índice nominal base oct-2001=100.
    Col 0: año (sólo en filas de cambio), col 1: mes, col 2: Total.
    El HTML deflacta por IPCBA para mostrar valor real base 2016=100.
    """
    wb = openpyxl.load_workbook(os.path.join(XLSX, 'masa_salarial.xlsx'), data_only=True)
    ws = wb['ee_industria_masa_salarial']
    rows = list(ws.iter_rows(values_only=True))
    periodos, total = [], []
    cur_year = None
    for row in rows[3:]:
        # Año puede estar en col 0
        if row[0] is not None:
            cell = str(row[0]).strip()
            try:
                y = int(float(cell.replace('*','').strip()))
                if 2000 <= y <= 2100:
                    cur_year = y
            except Exception:
                # Llegamos a notas al pie
                if row[1] is None:
                    break
        # Mes en col 1
        if row[1] is None or cur_year is None:
            continue
        m = MESES_ES.get(_norm(str(row[1])).replace('*','').strip())
        if m is None:
            continue
        v = row[2]
        try:
            val = round(float(v), 3) if v is not None else None
        except Exception:
            val = None
        periodos.append(f'{cur_year}-{m:02d}')
        total.append(val)
    return {'periodos': periodos, 'total': total,
            'fuente': 'IDECBA · ee_industria_masa_salarial.xlsx · índice nominal base oct-2001=100'}


# ── Comercio exterior (exportaciones) ──────────────────
def parse_comex():
    """Parsea exportaciones CABA.
    comex_tot.xlsx: series anuales USD + %PGB (1993-2025).
    comex_zon.xlsx: por continente/zona económica en millones USD (2000-2025).
    """
    # ── TOT ────────────────────────────────────────────
    wb_t = openpyxl.load_workbook(os.path.join(XLSX, 'comex_tot.xlsx'), data_only=True)
    ws_t = wb_t['AX_CX_TOT']
    rows_t = list(ws_t.iter_rows(values_only=True))
    anios, exportaciones_usd, pct_pgb = [], [], []
    for row in rows_t[3:]:   # fila 0 título, 1 vacía, 2 header, 3 vacía, 4+ datos
        if row[0] is None:
            continue   # saltar filas vacías intermedias
        raw_anio = str(row[0]).replace('*', '').strip()
        try:
            anio = int(float(raw_anio))
        except Exception:
            break   # nota al pie u otro texto → fin de datos
        try:
            usd = float(row[1]) if row[1] is not None else None
            pct = float(row[2]) if row[2] is not None else None
        except Exception:
            break
        anios.append(anio)
        exportaciones_usd.append(round(usd) if usd is not None else None)
        pct_pgb.append(round(pct, 4) if pct is not None else None)

    # ── ZON ────────────────────────────────────────────
    wb_z = openpyxl.load_workbook(os.path.join(XLSX, 'comex_zon.xlsx'), data_only=True)
    ws_z = wb_z['AX_CX_ZON']
    rows_z = list(ws_z.iter_rows(values_only=True))
    header_z = rows_z[1]    # fila 1: nombre | año1 | año2 | ...
    anios_zona = []
    for v in header_z[1:]:
        if v is None:
            break
        raw = str(v).replace('*', '').strip()
        try:
            anios_zona.append(int(float(raw)))
        except Exception:
            break
    n_years = len(anios_zona)

    # Zonas: bloques económicos sub-continentales + Total y continentes principales
    # Esto da más detalle al chart de "Top zonas económicas"
    ZONAS_OK = {
        'Total', 'América', 'MERCOSUR', 'USMCA', 'MCCA', 'Resto de América',
        'Europa', 'Unión Europea', 'Resto de Europa',
        'Asia', 'ASEAN', 'Resto de Asia',
        'África', 'SACU', 'Resto de África',
        'Oceanía', 'Indeterminado',
    }
    STOP_PREFIXES = ('.', 'Nota', 'Fuente', 'Desde', 'La info', 'Indeterm comprende')
    por_zona = {}
    for row in rows_z[2:]:
        if not row[0]:
            break
        name = str(row[0]).strip()
        if any(name.startswith(p) for p in STOP_PREFIXES):
            break
        if name not in ZONAS_OK:
            continue
        vals = []
        for i in range(1, n_years + 1):
            v = row[i] if i < len(row) else None
            try:
                vals.append(round(float(v), 3) if v is not None else None)
            except Exception:
                vals.append(None)
        por_zona[name] = vals

    return {
        'anios': anios,
        'exportaciones_usd': exportaciones_usd,
        'pct_pgb': pct_pgb,
        'anios_zona': anios_zona,
        'por_zona': por_zona,
    }


# ── Población por comuna ────────────────────────────────
def parse_poblacion():
    wb = xlrd.open_workbook(os.path.join(XLSX,'pob_comuna.xls'))
    ws = wb.sheet_by_index(0)
    # r2: años (2022..2035), r3: Total, r4..r18: comunas 1-15
    years = [int(ws.cell_value(2,j)) for j in range(1,ws.ncols) if isinstance(ws.cell_value(2,j),(int,float)) and ws.cell_value(2,j)>2000]
    ambos = {}
    for ci in range(15):
        row_i = 4 + ci
        comuna = int(ws.cell_value(row_i,0))
        vals = [int(ws.cell_value(row_i,j+1)) for j in range(len(years))]
        ambos[comuna] = vals
    total_by_year = [int(ws.cell_value(3,j+1)) for j in range(len(years))]
    return {'años':years, 'total':total_by_year, 'comunas':ambos}

def main():
    out = {
        'ipcba': parse_ipcba(),
        'iae': parse_iae(),
        'canastas': parse_canastas(),
        'empleo': parse_empleo(),
        'locales': parse_locales(),
        'industria_ingresos': parse_industria('ingresos'),
        'industria_personal': parse_industria('personal'),
        'poblacion': parse_poblacion(),
        'pobreza_tasas': parse_pobreza_tasas(),
        'comex': parse_comex(),
        'autoservicios': parse_autoservicios(),
        'supermercados': parse_supermercados(),
        'shoppings': parse_shoppings(),
        'masa_salarial': parse_masa_salarial(),
        'fuente': 'IDECBA — Instituto de Estadística y Censos GCBA',
        'generado': datetime.now().strftime('%Y-%m-%d'),
    }
    # Derivar MACRO.pobreza desde canastas (Hogar 1a: pareja adulta).
    # lp_hogares[i][1] y li_hogares[i][1] = canasta total y CA respectivamente.
    # El HTML usa arr[1] para "hogar tipo 2 (pareja con hijos)" — coincide con Hogar 1a de IDECBA.
    c = out['canastas']
    out['pobreza'] = {
        'periodos': c['meses'],
        'lp_hogares': [[None, v, None, None, None] for v in c['total']],
        'li_hogares': [[None, v, None, None, None] for v in c['ca']],
        'hogares_labels': [None, 'Pareja c/ 2 hijos', None, None, None],
    }
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',',':'))
    print(f'Wrote {OUT} ({os.path.getsize(OUT):,} bytes)')
    # resumen
    print(f'  ipcba: {len(out["ipcba"]["meses"])} meses, último: {out["ipcba"]["meses"][-1]} niv={out["ipcba"]["nivel_general"][-1]} varM={out["ipcba"]["var_mensual"][-1]}')
    print(f'  iae: {len(out["iae"]["trimestres"])} trims, último: {out["iae"]["trimestres"][-1]} idx={out["iae"]["indice"][-1]}')
    print(f'  canastas: {len(out["canastas"]["meses"])} meses, último: {out["canastas"]["meses"][-1]} CA={out["canastas"]["ca"][-1]} total={out["canastas"]["total"][-1]}')
    print(f'  empleo: {len(out["empleo"]["trimestres"])} trims, último: {out["empleo"]["trimestres"][-1]} desoc={out["empleo"]["desocupacion"][-1]}')
    print(f'  locales: {len(out["locales"]["periodos"])} periodos, ejes: {list(out["locales"]["ejes"].keys())[:5]}')
    print(f'  industria_ing: {len(out["industria_ingresos"]["periodos"])} meses, último: {out["industria_ingresos"]["periodos"][-1]}')
    print(f'  poblacion: {len(out["poblacion"]["comunas"])} comunas, años {out["poblacion"]["años"][0]}-{out["poblacion"]["años"][-1]}')
    pob = out['pobreza_tasas']
    print(f'  pobreza_tasas: {len(pob["periodos"])} periodos, ultimo: {pob["periodos"][-1]} pob_hog={pob["pob_hog_pct"][-1]}% ind_hog={pob["ind_hog_pct"][-1]}%')
    cx = out['comex']
    print(f'  comex: {len(cx["anios"])} años ({cx["anios"][0]}-{cx["anios"][-1]}) USD={cx["exportaciones_usd"][-1]:,} zonas={list(cx["por_zona"].keys())}')
    print(f'  pobreza (derivada): {len(out["pobreza"]["periodos"])} meses, último LP={out["pobreza"]["lp_hogares"][-1][1]}')
    aut = out['autoservicios']
    print(f'  autoservicios: {len(aut["periodos"])} meses ({aut["periodos"][0]}..{aut["periodos"][-1]}) últ. var.i.a.={aut["var_ia"][-1]}')
    shp = out['shoppings']
    print(f'  shoppings: {len(shp["periodos"])} meses ({shp["periodos"][0]}..{shp["periodos"][-1]}) {len(shp["rubros"])} rubros')
    sm = out['supermercados']
    nn = next((sm["var_ia"][i] for i in range(len(sm["var_ia"])-1,-1,-1) if sm["var_ia"][i] is not None), None)
    print(f'  supermercados: {len(sm["periodos"])} meses ({sm["periodos"][0]}..{sm["periodos"][-1]}) últ. var.i.a.={nn}')
    ms = out['masa_salarial']
    print(f'  masa_salarial: {len(ms["periodos"])} meses ({ms["periodos"][0]}..{ms["periodos"][-1]}) últ. índice={ms["total"][-1]}')

if __name__=='__main__': main()
