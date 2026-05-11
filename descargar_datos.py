import os
import requests
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX_DIR = os.path.join(BASE, 'idecba')
os.makedirs(XLSX_DIR, exist_ok=True)

WP_REST = 'https://www.estadisticaciudad.gob.ar/eyc/wp-json/wp/v2/banco_datos'
H = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}

# Diccionario de búsqueda refinada
ARCHIVOS = {
    'ipcba_evol.xlsx': 'ipcba mensual', # Buscamos "mensual" para asegurar el último
    'iae.xlsx': 'iae',
    'canastas.xlsx': 'canastas',
    'empleo.xlsx': 'etoi',
}

def descargar_lo_mas_nuevo():
    for nombre_local, termino in ARCHIVOS.items():
        print(f"Buscando el más reciente para: {termino}...")
        try:
            # Pedimos los últimos 5 resultados ordenados por fecha de publicación
            params = {'search': termino, 'orderby': 'date', 'order': 'desc', 'per_page': 5}
            r = requests.get(WP_REST, params=params, headers=H, timeout=20)
            
            if r.status_code == 200:
                items = r.json()
                if not items:
                    print(f" [!] No se encontró nada para {termino}")
                    continue
                
                # Buscamos el primer link válido que termine en .xlsx
                encontrado = False
                for item in items:
                    link = item.get('link_archivo', '')
                    if link.lower().endswith('.xlsx'):
                        print(f" [+] Encontrado dato nuevo: {link}")
                        res = requests.get(link, headers=H, timeout=20)
                        with open(os.path.join(XLSX_DIR, nombre_local), 'wb') as f:
                            f.write(res.content)
                        print(f" [OK] {nombre_local} actualizado.")
                        encontrado = True
                        break
                if not encontrado:
                    print(f" [!] No se halló link .xlsx en los resultados de {termino}")
        except Exception as e:
            print(f" [X] Error descargando {nombre_local}: {e}")

if __name__ == '__main__':
    descargar_lo_mas_nuevo()
