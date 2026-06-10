# Actualización diaria del historial — corre en GitHub Actions a las 19:00.
#
# El token de Bsale entrega ventas, guías de despacho y el stock ACTUAL,
# pero no las recepciones ni los consumos. Este script los deduce:
#
#   1. Baja las ventas y guías nuevas desde la API y las agrega al historial.
#   2. Compara el stock real de hoy contra la foto de ayer + movimientos conocidos.
#   3. La diferencia que no explican las ventas/guías se registra como
#      recepción (sube) o consumo (baja), marcada con 'inferido': true.
#   4. Guarda historial.json y la nueva foto en stock_snapshot.json.
#
# Si la API no responde o no hay token, NO modifica nada.
import json, os, time, sys
from datetime import datetime, timedelta
import requests

from generar_dashboard import NOMBRES, bsale_stock, CARPETA, BSALE_TOKEN

ARCHIVO_HISTORIAL = os.path.join(CARPETA, 'historial.json')
ARCHIVO_SNAPSHOT  = os.path.join(CARPETA, 'stock_snapshot.json')

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

def main():
    hoy = datetime.now()
    print(f"Actualización diaria — {hoy:%Y-%m-%d %H:%M}")

    # ── Stock real de hoy (si falla, abortamos sin tocar nada) ──
    stock_real = bsale_stock()
    if not stock_real:
        print("ERROR: sin datos de stock desde Bsale (token/API). No se modifica nada.")
        sys.exit(1)
    print(f"  Stock real: {len(stock_real)} SKUs")

    cache = json.load(open(ARCHIVO_HISTORIAL, encoding='utf-8'))
    movs  = cache['movimientos']
    conocidos = {(m['doc_id'], m['sku'], m['oficina'], m['tipo'])
                 for m in movs if m.get('doc_id')}

    snap = None
    if os.path.exists(ARCHIVO_SNAPSHOT):
        snap = json.load(open(ARCHIVO_SNAPSHOT, encoding='utf-8'))

    # ── 1. Ventas y guías nuevas desde la API ──
    if snap:
        desde = datetime.fromisoformat(snap['fecha']) - timedelta(days=1)
    else:
        ultima = max(m['fecha'] for m in movs)
        desde  = datetime.fromisoformat(ultima)
    nuevos_api = descargar_movimientos(desde, hoy)
    agregados = []
    for m in nuevos_api:
        clave = (m['doc_id'], m['sku'], m['oficina'], m['tipo'])
        if clave not in conocidos:
            movs.append(m)
            conocidos.add(clave)
            agregados.append(m)
    print(f"  API: {len(nuevos_api)} movimientos, {len(agregados)} nuevos agregados")

    # ── 2. Inferir recepciones/consumos comparando con la foto anterior ──
    inferidos = []
    if snap:
        fecha_hoy = hoy.strftime('%Y-%m-%d')
        for sku in NOMBRES:
            s_ant  = snap['stock'].get(sku, {'vit': 0, 'pat': 0})
            s_real = stock_real.get(sku, stock_real.get(sku.replace('Ñ', 'N'), {'vit': 0, 'pat': 0}))
            for of_key, of_nombre in (('vit', 'VIT'), ('pat', 'PAT')):
                salidas = sum(m['cantidad'] for m in agregados
                              if m['sku'] == sku and m['oficina'] == of_nombre)
                esperado = s_ant[of_key] - salidas
                residuo  = s_real[of_key] - esperado
                if residuo >= 1:
                    inferidos.append({'fecha': fecha_hoy, 'sku': sku, 'oficina': of_nombre,
                                      'tipo': 'produccion', 'cantidad': float(round(residuo)),
                                      'doc_id': 0, 'inferido': True})
                elif residuo <= -1:
                    inferidos.append({'fecha': fecha_hoy, 'sku': sku, 'oficina': of_nombre,
                                      'tipo': 'consumo', 'cantidad': float(round(-residuo)),
                                      'doc_id': 0, 'inferido': True})
        movs.extend(inferidos)
        for i in inferidos:
            signo = '+' if i['tipo'] == 'produccion' else '-'
            print(f"    inferido: {i['sku']} {i['oficina']} {signo}{i['cantidad']:.0f} ({i['tipo']})")
        print(f"  Inferidos: {len(inferidos)} movimientos")
    else:
        print("  Primera corrida: se guarda la foto base, la deducción parte mañana.")

    # ── 3. Guardar historial y nueva foto ──
    cache['ultimo_update'] = hoy.strftime('%Y-%m-%dT%H:%M:%S')
    cache['movimientos']   = movs
    with open(ARCHIVO_HISTORIAL, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

    foto = {sku: stock_real.get(sku, stock_real.get(sku.replace('Ñ', 'N'), {'vit': 0, 'pat': 0}))
            for sku in NOMBRES}
    with open(ARCHIVO_SNAPSHOT, 'w', encoding='utf-8') as f:
        json.dump({'fecha': hoy.strftime('%Y-%m-%d'), 'stock': foto}, f, ensure_ascii=False, indent=2)

    print(f"Historial: {len(movs)} movimientos | Foto guardada ({len(foto)} SKUs)")

if __name__ == '__main__':
    main()
