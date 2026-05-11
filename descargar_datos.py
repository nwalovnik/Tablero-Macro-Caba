import os
import requests

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX_DIR = os.path.join(BASE, 'idecba')
os.makedirs(XLSX_DIR, exist_ok=True)

WP_REST = 'https://www.estadisticaciudad.gob.ar/eyc/wp-json/wp/v2/banco_datos'
H = {'User-Agent': 'Mozilla/5.0'}

# Diccionario: Nombre que espera tu código -> Qué buscar en la web de la Ciudad
ARCHIVOS_A_ACTUALIZAR = {
    'ipcba_evol.xlsx': 'ipcba',
    'iae.xlsx': 'iae',
    'canastas.xlsx': 'canastas',
    'empleo.xlsx': 'etoi',
}

def actualizar_archivos():
    for nombre_local, termino_busqueda in ARCHIVOS_A_ACTUALIZAR.items():
        print(f"Buscando en API: '{termino_busqueda}' -> para actualizar {nombre_local}")
        try:
            r = requests.get(WP_REST, params={'search': termino_busqueda}, headers=H, timeout=15)
            if r.status_code == 200:
                resultados = r.json()
                for item in resultados:
                    link = item.get('link_archivo', '')
                    if link.endswith('.xlsx') or link.endswith('.xls'):
                        print(f" [+] Descargando: {link}")
                        res = requests.get(link, headers=H, timeout=15)
                        ruta_destino = os.path.join(XLSX_DIR, nombre_local)
                        with open(ruta_destino, 'wb') as f:
                            f.write(res.content)
                        print(f" [OK] Guardado y actualizado: {nombre_local}")
                        break
        except Exception as e:
            print(f" [X] Error con {nombre_local}: {e}")

if __name__ == '__main__':
    # 1. Arreglamos el problema de Linux (Mayúsculas vs Minúsculas) en lo que ya está subido
    for filename in os.listdir(XLSX_DIR):
        ruta_vieja = os.path.join(XLSX_DIR, filename)
        ruta_nueva = os.path.join(XLSX_DIR, filename.lower())
        if ruta_vieja != ruta_nueva:
            os.rename(ruta_vieja, ruta_nueva)
            print(f"Renombrado para compatibilidad Linux: {filename} a {filename.lower()}")
            
    # 2. Descargamos las actualizaciones de la Ciudad
    actualizar_archivos()
