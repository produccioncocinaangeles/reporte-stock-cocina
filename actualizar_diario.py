# Actualización diaria del historial — corre en GitHub Actions a las 19:00.
#
# El token de Bsale entrega ventas, guías de despacho y el stock ACTUAL,
# pero no las recepciones ni los consumos. Este script lleva nuestro propio
# registro de stock y lo cuadra a diario contra el stock real de Bsale:
#
#   1. Baja las ventas y guías nuevas desde la API y las agrega al historial.
#   2. Reconstruye el stock según NUESTRO sistema (igual que el dashboard).
#   3. Lo compara con el stock real del token, producto por producto:
#        - sube sin explicación  -> recepción  (entrada deducida)
#        - baja sin explicación  -> consumo    (salida deducida) => ALERTA
#   4. Guarda historial.json, la foto del día (stock_snapshot.json) y las
#      alertas (alertas_stock.json) que el dashboard muestra en el Resumen.
#
# Tras cada corrida, nuestro stock reconstruido == stock real de Bsale.
# Si la API no responde o no hay token, NO modifica nada.
import json, os, time, sys
from datetime import datetime, timedelta
import requests

from generar_dashboard import NOMBRES, bsale_stock, CARPETA, BSALE_TOKEN

ARCHIVO_HISTORIAL = os.path.join(CARPETA, 'historial.json')
ARCHIVO_SNAPSHOT  = os.path.join(CARPETA, 'stock_snapshot.json')
ARCHIVO_ALERTAS   = os.path.join(CARPETA, 'alertas_stock.json')

HEADERS   = {'access_token': BSALE_TOKEN}
OFICINAS  = {1: 'VIT', 3: 'PAT'}
DOC_VENTA = {35, 6}
DOC_DESP  = {8}

def get_all(url, params=None):
    items, offset = [], 0
    while True:
        p = {**(params or {}), 'limit': 50, 'offset': offset}
        r = requests.get(url, headers=HEADERS, params=p, timeout=30)
        r.raise_for_status()
        data  = r.json()
        batch = data.get('items', [])
        items.extend(batch)
        if not data.get('next') or len(batch) < 50:
            break
        offset += 50
        time.sleep(0.1)
    return items

def descargar_movimientos(desde, hasta):
    ts = lambda dt: int(dt.timestamp())
    movs = []
    for office_id, nombre in OFICINAS.items():
        for dt_id in DOC_VENTA | DOC_DESP:
            tipo = 'venta' if dt_id in DOC_VENTA else 'despacho'
            docs = get_all('https://api.bsale.cl/v1/documents.json', {
                'documenttypeid': dt_id,
                'officeid': office_id,
                'emissiondaterange': f'[{ts(desde)},{ts(hasta)}]',
            })
            for doc in docs:
                fecha = datetime.fromtimestamp(doc['emissionDate']).strftime('%Y-%m-%d')
                dets  = get_all(f"https://api.bsale.cl/v1/documents/{doc['id']}/details.json")
                for det in dets:
                    sku = str(det.get('variant', {}).get('code', '')).strip().upper()
                    if sku:
                        movs.append({'fecha': fecha, 'sku': sku, 'oficina': nombre,
                                     'tipo': tipo, 'cantidad': float(det.get('quantity', 0)),
                                     'doc_id': doc['id']})
            time.sleep(0.2)
    return movs

def reconstruir_stock(movs):
    # Réplica exacta de leer_json() del dashboard: por oficina, movimientos
    # ordenados por fecha, stock nunca negativo.
    stock = {}
    for oficina in ('VIT', 'PAT'):
        acum = {}
        for m in sorted([x for x in movs if x['oficina'] == oficina], key=lambda x: x['fecha']):
            ent = m['cantidad'] if m['tipo'] == 'produccion' else 0
            sal = m['cantidad'] if m['tipo'] in ('venta', 'despacho', 'consumo') else 0
            acum[m['sku']] = max(0, acum.get(m['sku'], 0) + ent - sal)
        for sku, q in acum.items():
            stock[(sku, oficina)] = q
    return stock

def main():
    hoy = datetime.now()
    print(f"Actualización diaria — {hoy:%Y-%m-%d %H:%M}")

    # ── Stock real (si falla, abortamos sin tocar nada) ──
    stock_real = bsale_stock()
    if not stock_real:
        print("ERROR: sin datos de stock desde Bsale (token/API). No se modifica nada.")
        sys.exit(1)
    print(f"  Stock real: {len(stock_real)} SKUs")

    cache = json.load(open(ARCHIVO_HISTORIAL, encoding='utf-8'))
    movs  = cache['movimientos']
    conocidos = {(m['doc_id'], m['sku'], m['oficina'], m['tipo'])
                 for m in movs if m.get('doc_id')}

    # ── 1. Ventas y guías nuevas desde la API ──
    ultima = max(m['fecha'] for m in movs)
    desde  = datetime.fromisoformat(ultima) - timedelta(days=1)
    nuevos_api = descargar_movimientos(desde, hoy)
    agregados = 0
    for m in nuevos_api:
        clave = (m['doc_id'], m['sku'], m['oficina'], m['tipo'])
        if clave not in conocidos:
            movs.append(m)
            conocidos.add(clave)
            agregados += 1
    print(f"  API: {len(nuevos_api)} movimientos, {agregados} nuevos")

    # ── 2. Cuadratura: nuestro stock vs stock real de Bsale ──
    recon = reconstruir_stock(movs)
    fecha_hoy = hoy.strftime('%Y-%m-%d')
    inferidos = []
    for sku in NOMBRES:
        s_real = stock_real.get(sku, stock_real.get(sku.replace('Ñ', 'N'), {'vit': 0, 'pat': 0}))
        for of_key, of_nombre in (('vit', 'VIT'), ('pat', 'PAT')):
            residuo = s_real[of_key] - recon.get((sku, of_nombre), 0)
            if residuo >= 1:
                inferidos.append({'fecha': fecha_hoy, 'sku': sku, 'oficina': of_nombre,
                                  'tipo': 'produccion', 'cantidad': float(round(residuo)),
                                  'doc_id': 0, 'inferido': True})
            elif residuo <= -1:
                inferidos.append({'fecha': fecha_hoy, 'sku': sku, 'oficina': of_nombre,
                                  'tipo': 'consumo', 'cantidad': float(round(-residuo)),
                                  'doc_id': 0, 'inferido': True})
    movs.extend(inferidos)

    entradas = [i for i in inferidos if i['tipo'] == 'produccion']
    salidas  = [i for i in inferidos if i['tipo'] == 'consumo']
    for i in entradas:
        print(f"    entrada deducida: {i['sku']} {i['oficina']} +{i['cantidad']:.0f}")
    for i in salidas:
        print(f"    ALERTA salida sin explicación: {i['sku']} {i['oficina']} -{i['cantidad']:.0f}")
    print(f"  Cuadratura: {len(entradas)} entradas, {len(salidas)} salidas sin explicación")

    # ── 3. Guardar todo ──
    cache['ultimo_update'] = hoy.strftime('%Y-%m-%dT%H:%M:%S')
    cache['movimientos']   = movs
    with open(ARCHIVO_HISTORIAL, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

    foto = {sku: stock_real.get(sku, stock_real.get(sku.replace('Ñ', 'N'), {'vit': 0, 'pat': 0}))
            for sku in NOMBRES}
    with open(ARCHIVO_SNAPSHOT, 'w', encoding='utf-8') as f:
        json.dump({'fecha': fecha_hoy, 'stock': foto}, f, ensure_ascii=False, indent=2)

    with open(ARCHIVO_ALERTAS, 'w', encoding='utf-8') as f:
        json.dump({'fecha': fecha_hoy,
                   'entradas': [{'sku': i['sku'], 'oficina': i['oficina'], 'cantidad': i['cantidad']} for i in entradas],
                   'salidas':  [{'sku': i['sku'], 'oficina': i['oficina'], 'cantidad': i['cantidad']} for i in salidas]},
                  f, ensure_ascii=False, indent=2)

    print(f"Historial: {len(movs)} movimientos | Foto y alertas guardadas")

if __name__ == '__main__':
    main()
