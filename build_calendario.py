"""Scrapea el cronograma de publicaciones de IDECBA y genera calendario.json.

Server-side (corre en GitHub Actions, no necesita CORS proxy).
Reemplaza al scraping en vivo que hacia el HTML via proxies inestables.

La logica replica la del JS calParseHTML del tablero:
  - El HTML del cronograma esta armado con bloques data-elementor-type="loop-item"
  - Cada bloque contiene fecha DD/MM/YYYY y un <h6> con el titulo
  - Categorizacion por regex sobre el titulo
"""
import json
import os
import re
import sys
import html as ihtml
from datetime import datetime, date
import requests

BASE = os.path.dirname(os.path.abspath(__file__))
OUT  = os.path.join(BASE, 'calendario.json')
HTML = os.path.join(BASE, 'tablero-macro.html')
SRC  = 'https://www.estadisticaciudad.gob.ar/eyc/calendario-listado/'

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/120.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'es-AR,es;q=0.9,en;q=0.8',
}

# Categorias - replica del CAL_CATS en el HTML
CATS = [
    (re.compile(r'\bIPCBA\b', re.I),                                          'ipcba',      'IPCBA',                          'macro'),
    (re.compile(r'L[íi]neas de pobreza|Canastas', re.I),                      'canastas',   'Canastas y líneas de pobreza',   'macro'),
    (re.compile(r'Actividad Econ[óo]mica.*PGB|PGB.*Trimestral', re.I),        'iae',        'Actividad económica · PGB',      'macro'),
    (re.compile(r'\bSIPCBA\b|[ÍI]ndices de Precios de la Construcci[óo]n', re.I), 'sipcba', 'SIPCBA · construcción',          'macro'),
    (re.compile(r'Mercado laboral|empleo en la Ciudad|ETOI', re.I),           'empleo',     'Mercado laboral (ETOI)',         'macro'),
    (re.compile(r'Ejes comerciales|Din[áa]mica del comercio minorista', re.I),'locales',    'Comercio minorista · ejes',      'macro'),
    (re.compile(r'Ingresos en la Ciudad', re.I),                              'ingresos',   'Ingresos',                       'macro'),
    (re.compile(r'Condiciones de vida', re.I),                                'condiciones','Condiciones de vida',            'macro'),
    (re.compile(r'Mercado de alquiler', re.I),                                'alquiler',   'Mercado de alquiler',            'otros'),
    (re.compile(r'Mercado de venta|Din[áa]mica de departamentos', re.I),      'venta',      'Mercado de venta',               'otros'),
    (re.compile(r'Canasta de Crianza', re.I),                                 'crianza',    'Canasta de Crianza',             'otros'),
    (re.compile(r'Exportaciones', re.I),                                      'export',     'Exportaciones',                  'otros'),
    (re.compile(r'Encuesta Anual de Hogares', re.I),                          'eah',        'Encuesta Anual de Hogares',      'otros'),
    (re.compile(r'Barrios Populares|\bBaPIs\b', re.I),                        'bapis',      'Barrios Populares',              'otros'),
    (re.compile(r'Personas en Situaci[óo]n de Calle', re.I),                  'calle',      'Personas en situación de calle', 'otros'),
    (re.compile(r'Pobreza multidimensional', re.I),                           'pobrezaM',   'Pobreza multidimensional',       'otros'),
    (re.compile(r'Censos Demogr[áa]ficos|Caracter[íi]sticas demogr[áa]ficas', re.I), 'censo','Censos demográficos',            'otros'),
    (re.compile(r'Nupcialidad|Matrimonios', re.I),                            'nupci',      'Nupcialidad',                    'otros'),
    (re.compile(r'Mortalidad|defunciones', re.I),                             'mort',       'Mortalidad',                     'otros'),
    (re.compile(r'Fecundidad|nacimientos', re.I),                             'fecund',     'Fecundidad',                     'otros'),
    (re.compile(r'Caracter[íi]sticas poblacionales', re.I),                   'poblac',     'Características poblacionales',  'otros'),
    (re.compile(r'Elecciones', re.I),                                         'elec',       'Elecciones',                     'otros'),
    (re.compile(r'Comunas en la web', re.I),                                  'comunasweb', 'Portal web de comunas',          'otros'),
]


def categorize(titulo):
    for pat, cid, cl, tab in CATS:
        if pat.search(titulo):
            return cid, cl, tab
    return 'otros', 'Otros', 'otros'


def clean(s):
    if not s:
        return ''
    s = ihtml.unescape(s)
    return re.sub(r'\s+', ' ', s).strip()


def short_desc(t):
    t = re.sub(r'\s+Ciudad de Buenos Aires\.?', '', t)
    t = re.sub(r'\s+GCBA\.?', '', t)
    return re.sub(r'\.$', '', t.strip())


MESES_NUM = {
    'Enero': 1, 'Febrero': 2, 'Marzo': 3, 'Abril': 4, 'Mayo': 5, 'Junio': 6,
    'Julio': 7, 'Agosto': 8, 'Septiembre': 9, 'Octubre': 10, 'Noviembre': 11, 'Diciembre': 12,
}


def extract_periodo(t):
    m = re.search(r'(\d(?:er|do|to|ro|vo|mo)\.?\s*(?:trimestre|cuatrimestre))\s+(?:de\s+)?(20\d{2})', t, re.I)
    if m:
        return clean(m.group(1) + ' ' + m.group(2))
    m = re.search(r'(' + '|'.join(MESES_NUM.keys()) + r')\s+(?:de\s+)?(20\d{2})', t, re.I)
    if m:
        return clean(m.group(1).capitalize() + ' ' + m.group(2))
    m = re.search(r'A[ñn]o\s+(20\d{2})', t, re.I)
    if m:
        return 'Año ' + m.group(1)
    m = re.search(r'\b(20\d{2})\b', t)
    if m:
        return m.group(1)
    return ''


def parse_html(html):
    """Replica calParseHTML del JS: split por loop-item, extrae fecha + h6."""
    chunks = html.split('<div data-elementor-type="loop-item"')
    events = []
    seen = set()
    for ch in chunks[1:]:
        dm = re.search(r'(\d{2})/(\d{2})/(20\d{2})', ch)
        tm = re.search(r'<h6[^>]*elementor-heading-title[^>]*>([^<]+)</h6>', ch)
        if not dm or not tm:
            continue
        titulo = clean(tm.group(1))
        if not titulo:
            continue
        fecha = f'{dm.group(3)}-{dm.group(2)}-{dm.group(1)}'
        k = fecha + '|' + titulo
        if k in seen:
            continue
        seen.add(k)
        lm = re.search(r'<a\s+href="(https://www\.estadisticaciudad\.gob\.ar/eyc/publicaciones/[^"]+)"', ch)
        cid, cl, tab = categorize(titulo)
        events.append({
            'fecha':      fecha,
            'titulo':     titulo,
            'desc_corta': short_desc(titulo),
            'periodo':    extract_periodo(titulo),
            'cat_id':     cid,
            'cat_label':  cl,
            'tab':        tab,
            'url':        lm.group(1) if lm else '',
        })
    events.sort(key=lambda x: (x['fecha'], x['titulo']))
    return events


def main():
    print(f'[cal] descargando {SRC}', flush=True)
    r = requests.get(SRC, headers=HEADERS, timeout=60)
    r.raise_for_status()
    html = r.text
    print(f'[cal] HTML recibido ({len(html):,} bytes)', flush=True)

    events = parse_html(html)
    if not events:
        # Si la estructura cambia, fallar visiblemente para que se note en el workflow
        raise RuntimeError('No se parseo ningun evento. La estructura del HTML cambio?')

    hoy = date.today().isoformat()
    macro = [e for e in events if e['tab'] == 'macro']
    otros = [e for e in events if e['tab'] == 'otros']

    # Conservar presupuesto del JSON previo si existe (ese bloque no viene del scraping)
    presupuesto = []
    if os.path.exists(OUT):
        try:
            with open(OUT, 'r', encoding='utf-8') as f:
                prev = json.load(f)
            presupuesto = prev.get('presupuesto', []) or []
        except Exception:
            pass

    out = {
        'generado':    datetime.now().isoformat(timespec='seconds'),
        'fuente':      'IDECBA · ' + SRC,
        'hoy':         hoy,
        'macro':       macro,
        'otros':       otros,
        'presupuesto': presupuesto,
    }
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))

    print(f'[cal] OK -> {OUT}', flush=True)
    print(f'  macro: {len(macro)} eventos · otros: {len(otros)} eventos', flush=True)
    if macro:
        prox = [e for e in macro if e['fecha'] >= hoy][:3]
        for e in prox:
            print(f'    - {e["fecha"]} · {e["cat_label"]} · {e["periodo"]}', flush=True)

    # Inyectar en el HTML standalone para que la primera carga ya tenga datos
    if os.path.exists(HTML):
        with open(HTML, 'r', encoding='utf-8') as f:
            doc = f.read()
        inline = json.dumps(out, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
        pat = re.compile(r'(<script\s+id="calData"[^>]*>)([\s\S]*?)(</script>)')
        new, n = pat.subn(lambda m: m.group(1) + inline + m.group(3), doc, count=1)
        if n:
            with open(HTML, 'w', encoding='utf-8') as f:
                f.write(new)
            print(f'[cal] HTML parcheado · calData inline ({len(inline):,} bytes)', flush=True)
        else:
            print('[cal] WARN: no se encontro <script id="calData"> en el HTML', flush=True)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'ERROR build_calendario: {e}', file=sys.stderr)
        sys.exit(1)
