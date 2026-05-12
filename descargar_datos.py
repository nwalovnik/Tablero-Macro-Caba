"""Baja los XLSX/XLS mas recientes del banco de datos IDECBA.

Para cada archivo:
  1. Busca el post mas reciente en wp-json/v2/banco_datos con el termino dado.
  2. Abre la pagina del post y extrae el primer href que matchee el patron.
  3. Si todo eso falla, usa el URL fallback hardcodeado.

Los archivos se guardan en idecba/ con el nombre que espera build_macro_data.py.
"""
import os
import re
import sys
import time
import requests

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX_DIR = os.path.join(BASE, 'idecba')
os.makedirs(XLSX_DIR, exist_ok=True)

WP_REST = 'https://www.estadisticaciudad.gob.ar/eyc/wp-json/wp/v2/banco_datos'
H = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/120.0 Safari/537.36',
    'Accept': 'text/html,application/json,*/*;q=0.8',
    'Accept-Language': 'es-AR,es;q=0.9,en;q=0.8',
}
TIMEOUT_API = 20
TIMEOUT_DL  = 60

# Cada entrada:
#   filename: nombre local (lo que espera build_macro_data.py)
#   search:   query para wp-json (terminos especificos)
#   pattern:  regex que matchea el .xlsx en el HTML del post
#   fallback: URL hardcoded del XLSX (se usa solo si descubrir dinamicamente falla)
ARCHIVOS = [
    {
        'filename': 'ipcba_evol.xlsx',
        'search':   'IPCBA indice mensual nivel general empalme',
        'pattern':  r'href="(https://www\.estadisticaciudad\.gob\.ar/eyc/wp-content/uploads/[^"]+IPCBA[^"]+(?:nivel[_\-]general|empalme|evol|serie[_\-]mensual)[^"]*\.xlsx)"',
        'fallback': 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2026/03/IPCBA_base_2021100-Nivel_general_empalme.xlsx',
    },
    {
        'filename': 'iae.xlsx',
        'search':   'producto geografico bruto trimestral PGB Ciudad Buenos Aires',
        'pattern':  r'href="(https://www\.estadisticaciudad\.gob\.ar/eyc/wp-content/uploads/[^"]+PGB[^"]+\.xlsx)"',
        'fallback': 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2025/12/PGB_K_variacion_porcentual.xlsx',
    },
    {
        'filename': 'canastas.xlsx',
        'search':   'canastas consumo hogar lineas pobreza indigencia',
        'pattern':  r'href="(https://www\.estadisticaciudad\.gob\.ar/eyc/wp-content/uploads/[^"]+[Cc]anasta[^"]*\.xlsx)"',
        'fallback': 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2024/12/Canastas_y_Lineas_de_Pobreza.xlsx',
    },
    {
        'filename': 'empleo.xlsx',
        'search':   'mercado laboral ETOI tasas actividad empleo desocupacion trimestral',
        'pattern':  r'href="(https://www\.estadisticaciudad\.gob\.ar/eyc/wp-content/uploads/[^"]+(?:ETOI|etoi|laboral|empleo)[^"]*\.xlsx)"',
        'fallback': 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2024/12/ETOI_series_historicas.xlsx',
    },
    {
        'filename': 'locales_abs.xlsx',
        'search':   'ejes comerciales locales vacancia cuatrimestre',
        'pattern':  r'href="(https://www\.estadisticaciudad\.gob\.ar/eyc/wp-content/uploads/[^"]+(?:ejes|locales|vacancia|AC[_\-]EJ)[^"]*\.xlsx)"',
        'fallback': 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2024/12/Ejes_comerciales_series.xlsx',
    },
    {
        'filename': 'industria_ing.xlsx',
        'search':   'industria manufacturera ingresos fabriles rama actividad',
        'pattern':  r'href="(https://www\.estadisticaciudad\.gob\.ar/eyc/wp-content/uploads/[^"]+(?:industria[^"]*ing|ingresos_fabril|ee_industria)[^"]*\.xlsx)"',
        'fallback': 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2024/12/ee_industria_ingresos_fabriles.xlsx',
    },
    {
        'filename': 'industria_pers.xlsx',
        'search':   'industria manufacturera personal asalariado rama actividad',
        'pattern':  r'href="(https://www\.estadisticaciudad\.gob\.ar/eyc/wp-content/uploads/[^"]+(?:industria[^"]*pers|personal_asalariado|ee_ind_personal)[^"]*\.xlsx)"',
        'fallback': 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2024/12/ee_ind_personal_asalariado.xlsx',
    },
    {
        'filename': 'pobreza_tasas.xlsx',
        'search':   'pobreza indigencia hogares personas tasas trimestral CV_AX15',
        'pattern':  r'href="(https://www\.estadisticaciudad\.gob\.ar/eyc/wp-content/uploads/[^"]+(?:CV[_\-]AX15|pobreza[^"]*tasas|tasas[^"]*pobreza)[^"]*\.xlsx)"',
        'fallback': 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2025/09/CV_AX15.xlsx',
    },
]


def _get_with_retry(url, retries=1, label='get'):
    """GET con reintentos. Devuelve None si todos los intentos fallan."""
    last = None
    for attempt in range(retries + 1):
        try:
            r = requests.get(url, headers=H, timeout=TIMEOUT_API)
            r.raise_for_status()
            return r
        except Exception as e:
            last = e
            if attempt < retries:
                time.sleep(2)
    print(f'   {label} fallo: {str(last)[:80]}', flush=True)
    return None


def find_xlsx_url(search, pattern):
    """Devuelve el .xlsx mas reciente cuya URL matchee el patron especifico.
    Si ningun post matchea, devuelve None (no usar fallback generico: peligroso
    para datasets con multiples aperturas como IPCBA).
    """
    r = _get_with_retry(
        f'{WP_REST}?search={requests.utils.quote(search)}&orderby=date&order=desc&per_page=3',
        retries=1, label='wp-json',
    )
    if r is None:
        return None
    try:
        posts = r.json()
    except Exception as e:
        print(f'   wp-json JSON parse: {e}', flush=True)
        return None
    if not posts:
        return None

    pat = re.compile(pattern, re.I)
    for post in posts:
        link = post.get('link')
        if not link:
            continue
        pr = _get_with_retry(link, retries=1, label='post page')
        if pr is None:
            continue
        m = pat.search(pr.text)
        if m:
            return m.group(1)
    return None


def download_with_retry(url, dest, retries=1):
    # Forzar HTTPS: el servidor IDECBA timea conexiones por HTTP port 80
    if url.startswith('http://'):
        url = 'https://' + url[len('http://'):]
    last = None
    for attempt in range(retries + 1):
        try:
            r = requests.get(url, headers=H, timeout=TIMEOUT_DL, stream=True)
            r.raise_for_status()
            with open(dest, 'wb') as f:
                for chunk in r.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)
            return True
        except Exception as e:
            last = e
            if attempt < retries:
                time.sleep(3 * (attempt + 1))
    print(f'   descarga fallo despues de {retries+1} intentos: {last}', flush=True)
    return False


def main():
    print(f'IDECBA · descargando {len(ARCHIVOS)} archivos a {XLSX_DIR}', flush=True)
    ok, fallback_used, fail = 0, 0, 0
    for cfg in ARCHIVOS:
        name = cfg['filename']
        dest = os.path.join(XLSX_DIR, name)
        print(f'\n[{name}] buscando: {cfg["search"]!r}', flush=True)
        url = find_xlsx_url(cfg['search'], cfg['pattern'])
        if url:
            print(f'   dinamico OK: {url}', flush=True)
        else:
            url = cfg['fallback']
            print(f'   usando fallback: {url}', flush=True)
            fallback_used += 1
        if download_with_retry(url, dest):
            size = os.path.getsize(dest)
            print(f'   guardado {name} ({size:,} bytes)', flush=True)
            ok += 1
        else:
            # Si la descarga fallo pero el archivo previo existe (del commit), no fallar
            if os.path.exists(dest):
                print(f'   manteniendo version previa de {name}', flush=True)
            else:
                print(f'   ERROR: {name} no se pudo descargar y no hay version previa', flush=True)
                fail += 1

    print(f'\nResumen: {ok}/{len(ARCHIVOS)} bajados · {fallback_used} via fallback · {fail} sin datos', flush=True)
    # No abortar el workflow si la descarga dinamica falla — los XLSX previos del repo
    # sirven como respaldo; build_macro_data.py funciona con lo que haya en idecba/.


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'ERROR descargar_datos: {e}', file=sys.stderr)
        sys.exit(1)
