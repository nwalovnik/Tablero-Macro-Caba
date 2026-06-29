"""Baja los XLSX mas recientes del banco de datos IDECBA.

Estrategia:
  1. Para cada archivo necesario, hay una pagina de categoria de IDECBA hardcodeada
     (las categorias son estables, los posts dentro rotan).
  2. Bajamos la pagina de categoria y extraemos los hrefs de /eyc/banco-datos/<slug>/.
  3. Tomamos el primer slug que matchea una regex especifica del archivo.
     (la categoria lista los posts ordenados por fecha desc, asi que el primer
     match es el mas reciente).
  4. Bajamos esa pagina de dataset y extraemos el primer .xlsx.
  5. Bajamos el .xlsx forzando HTTPS y guardamos en idecba/<nombre>.

Si cualquier paso falla, mantenemos la version previa committeada del XLSX en
idecba/ — build_macro_data.py va a procesar lo que haya.

Archivos que no estan acá (empleo, industria, pobreza_tasas) se mantienen con la
version committeada porque los parsers de build_macro_data.py esperan layouts
muy especificos y no es trivial encontrar un equivalente actualizado en banco_datos.
"""
import os
import re
import sys
import time
import requests

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX_DIR = os.path.join(BASE, 'idecba')
os.makedirs(XLSX_DIR, exist_ok=True)

H = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/120.0 Safari/537.36',
    'Accept': 'text/html,application/json,*/*;q=0.8',
    'Accept-Language': 'es-AR,es;q=0.9,en;q=0.8',
}
TIMEOUT = 15

CAT_BASE = 'https://www.estadisticaciudad.gob.ar/eyc/categoria-banco-datos/'

# Cada entrada:
#   filename: nombre local (lo que espera build_macro_data.py)
#   categoria: URL slug de la pagina de categoria
#   slug_re: regex (case-insensitive) que matchea el slug del post correcto
#            dentro de esa categoria. El primer match (=mas reciente) gana.
ARCHIVOS = [
    {
        'filename':  'ipcba_evol.xlsx',
        'categoria': 'indice-mensual-base-2021',
        # build_macro_data.parse_ipcba() lee la hoja 'Evol_gral_estac_reg_resto'
        'slug_re':   r'evolucion-del-nivel-general-estacionales-regulados-y-resto',
    },
    # NOTA: canastas.xlsx no se baja — IDECBA cambio el layout (CV_01_AX20 en vez de Canasta_cons_hogar1).
    # parse_canastas() en build_macro_data.py espera el formato viejo. Hasta adaptar
    # el parser, mantenemos el XLSX previo committeado en idecba/.
    {
        'filename':  'iae.xlsx',
        'categoria': 'producto-geografico-bruto-pgb',
        # build_macro_data.parse_iae() lee la hoja 'PGB_ITAE_b12'
        'slug_re':   r'indicador-trimestral-de-actividad-economica',
    },
    {
        'filename':  'locales_abs.xlsx',
        'categoria': 'ejes-comerciales',
        # build_macro_data.parse_locales() lee multiples hojas '3er. cuatr. de YYYY'
        'slug_re':   r'locales-relevados-ocupados-y-desocupados.*53-ejes-comerciales',
    },
    {
        'filename':  'ejes48_comuna_tasas.xlsx',
        'categoria': 'ejes-comerciales',
        # 48 ejes (nueva metodología desde 2025): tasa ocupación por comuna, 4 cuatrimestres
        'slug_re':   r'locales-relevados-ocupados-densidad-comercial-tasa-de-ocupacion.*por-comuna-48-ejes',
    },
    {
        'filename':  'comex_tot.xlsx',
        'categoria': 'comercio-exterior',
        # build_macro_data.parse_comex() lee hoja 'AX_CX_TOT'
        'slug_re':   r'exportaciones-monto-fob-en-dolares-y-participacion',
    },
    {
        'filename':  'comex_zon.xlsx',
        'categoria': 'comercio-exterior',
        # build_macro_data.parse_comex() lee hoja 'AX_CX_ZON'
        'slug_re':   r'exportaciones-clasificadas-por-continente-y-zona-economica-millones',
    },
    {
        'filename':  'autoservicios.xlsx',
        'categoria': 'autoservicios-mayoristas',
        # build_macro_data.parse_autoservicios() lee hoja 'AC_M_01'
        'slug_re':   r'ventas-a-valores-constantes-en-autoservicios-mayoristas-variacion-interanual',
    },
    {
        'filename':  'shoppings.xlsx',
        'categoria': 'centros-de-compras',
        # build_macro_data.parse_shoppings() lee hoja 'AC_CC_AX07'
        'slug_re':   r'variacion-interanual-del-indice-a-valores-constantes-base-2021100-de-ventas-por-rubro-en-centros-de-compras',
    },
    {
        'filename':  'supermercados.xlsx',
        'categoria': 'supermercados',
        # build_macro_data.parse_supermercados() lee hoja 'AC_S_01'
        'slug_re':   r'ventas-a-valores-constantes-en-supermercados-variacion-interanual',
    },
    {
        'filename':  'masa_salarial.xlsx',
        'categoria': 'industria',
        # build_macro_data.parse_masa_salarial() lee hoja 'ee_industria_masa_salarial'
        'slug_re':   r'masa-salarial-por-rama-de-actividad-indice-base-octubre-2001',
    },
]

# Posts cuyo slug ya conocemos NO matchean nuestra regex (falsos positivos a evitar)
# por ahora ninguno, pero dejo la estructura por si aparecen.


HREF_POST = re.compile(
    r'href="(https://www\.estadisticaciudad\.gob\.ar/eyc/banco-datos/([^"/]+)/)"',
    re.I,
)
HREF_XLSX = re.compile(
    r'href="(https://www\.estadisticaciudad\.gob\.ar/eyc/wp-content/uploads/[^"]+\.xlsx)"',
    re.I,
)


def get(url, label):
    """GET con 1 reintento. Devuelve response o None."""
    last = None
    for attempt in range(2):
        try:
            r = requests.get(url, headers=H, timeout=TIMEOUT, allow_redirects=True)
            r.raise_for_status()
            return r
        except Exception as e:
            last = e
            if attempt == 0:
                time.sleep(2)
    print(f'   {label} fallo: {str(last)[:120]}', flush=True)
    return None


def encontrar_post(categoria, slug_re):
    """Bajar la categoria, devolver la URL del primer post cuyo slug matchea."""
    cat_url = CAT_BASE + categoria + '/'
    r = get(cat_url, f'categoria {categoria}')
    if r is None:
        return None
    pat = re.compile(slug_re, re.I)
    for m in HREF_POST.finditer(r.text):
        post_url, slug = m.group(1), m.group(2)
        if pat.search(slug):
            return post_url
    print(f'   ningun post matchea {slug_re!r} en la categoria', flush=True)
    return None


def extraer_xlsx(post_url):
    """Bajar la pagina del post, devolver la URL del primer XLSX."""
    r = get(post_url, 'post page')
    if r is None:
        return None
    m = HREF_XLSX.search(r.text)
    if not m:
        print(f'   sin .xlsx en {post_url}', flush=True)
        return None
    return m.group(1)


def descargar(xlsx_url, dest):
    """Bajar el XLSX forzando HTTPS, sin seguir redirects a HTTP."""
    if xlsx_url.startswith('http://'):
        xlsx_url = 'https://' + xlsx_url[len('http://'):]
    try:
        r = requests.get(xlsx_url, headers=H, timeout=TIMEOUT, stream=True,
                         allow_redirects=False)
        # Si hay redirect, seguir manualmente forzando HTTPS
        hops = 0
        while r.status_code in (301, 302, 303, 307, 308) and hops < 3:
            loc = r.headers.get('Location', '')
            if not loc:
                break
            if loc.startswith('http://'):
                loc = 'https://' + loc[len('http://'):]
            elif loc.startswith('/'):
                loc = 'https://www.estadisticaciudad.gob.ar' + loc
            r = requests.get(loc, headers=H, timeout=TIMEOUT, stream=True,
                             allow_redirects=False)
            hops += 1
        r.raise_for_status()
        with open(dest, 'wb') as f:
            for chunk in r.iter_content(chunk_size=65536):
                if chunk:
                    f.write(chunk)
        return True
    except Exception as e:
        print(f'   descarga fallo: {str(e)[:120]}', flush=True)
        return False


def main():
    print(f'IDECBA · {len(ARCHIVOS)} archivos via categoria-banco-datos', flush=True)
    ok = 0
    for cfg in ARCHIVOS:
        name = cfg['filename']
        dest = os.path.join(XLSX_DIR, name)
        print(f'\n[{name}] cat={cfg["categoria"]}  slug~{cfg["slug_re"][:60]}', flush=True)

        post_url = encontrar_post(cfg['categoria'], cfg['slug_re'])
        if not post_url:
            if os.path.exists(dest):
                print(f'   manteniendo version previa', flush=True)
            continue
        print(f'   post: {post_url}', flush=True)

        xlsx_url = extraer_xlsx(post_url)
        if not xlsx_url:
            if os.path.exists(dest):
                print(f'   manteniendo version previa', flush=True)
            continue
        print(f'   xlsx: {xlsx_url}', flush=True)

        if descargar(xlsx_url, dest):
            size = os.path.getsize(dest)
            print(f'   OK ({size:,} bytes)', flush=True)
            ok += 1
        elif os.path.exists(dest):
            print(f'   manteniendo version previa', flush=True)

    print(f'\nResumen: {ok}/{len(ARCHIVOS)} actualizados', flush=True)
    # El workflow no debe fallar por archivos que no se actualizaron — los previos sirven


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'ERROR descargar_datos: {e}', file=sys.stderr)
        sys.exit(1)
