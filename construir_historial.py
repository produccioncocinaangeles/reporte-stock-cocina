import requests, json, os, time
from datetime import datetime, timedelta

TOKEN     = os.environ.get('BSALE_TOKEN', '')
HEADERS   = {'access_token': TOKEN}
CACHE     = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'historial.json')

OFICINAS  = {1: 'VIT', 3: 'PAT'}
DOC_VENTA = {35, 6}   # boleta + factura
DOC_DESP  = {8}       # guía de despacho

# ── Helpers ──────────────────────────────────────────────────

def get_all(url, params=None):
    items, offset = [], 0
    while True:
        p = {**(params or {}), 'limit': 50, 'offset': offset}
        r = requests.get(url, headers=HEADERS, params=p)
        if r.status_code != 200:
            print(f"  Error {r.status_code}: {url}")
            break
        data = r.json()
        batch = data.get('items', [])
        items.extend(batch)
        if not data.get('next') or len(batch) < 50:
            break
        offset += 50
        time.sleep(0.1)
    return items

def ts(dt): return int(dt.timestamp())

# ── Descarga documentos de una sucursal ──────────────────────

def descargar_documentos(office_id, desde, hasta, doc_types):
    nombre = OFICINAS[office_id]
    movimientos = []

    for dt_id in doc_types:
        tipo = 'venta' if dt_id in DOC_VENTA else 'despacho'
        print(f"  {nombre} tipo {dt_id} ({tipo})...")
        docs = get_all('https://api.bsale.cl/v1/documents.json', {
            'documenttypeid': dt_id,
            'officeid': office_id,
            'emissiondatefieldinitial': ts(desde),
            'emissiondatefieldend': ts(hasta),
        })
        print(f"    {len(docs)} documentos")

        for doc in docs:
            doc_id   = doc['id']
            fecha    = datetime.fromtimestamp(doc['emissionDate']).strftime('%Y-%m-%d')
            detalles = get_all(f"https://api.bsale.cl/v1/documents/{doc_id}/details.json")
            for det in detalles:
                sku = det.get('variant', {}).get('code', '').strip().upper()
                if not sku:
                    continue
                movimientos.append({
                    'fecha':   fecha,
                    'sku':     sku,
                    'oficina': nombre,
                    'tipo':    tipo,
                    'cantidad': float(det.get('quantity', 0)),
                    'doc_id':  doc_id,
                })
        time.sleep(0.2)

    return movimientos

# ── Descarga recepciones (producción) ────────────────────────

def descargar_recepciones(office_id, desde, hasta):
    nombre = OFICINAS[office_id]
    print(f"  {nombre} recepciones (producción)...")
    recs = get_all('https://api.bsale.cl/v1/stocks/receptions.json', {
        'officeid': office_id,
        'admissiondateinitial': ts(desde),
        'admissiondateend': ts(hasta),
    })
    print(f"    {len(recs)} recepciones")

    movimientos = []
    for rec in recs:
        rec_id = rec['id']
        fecha  = rec.get('rawAdmissionDate', '')
        dets   = get_all(f"https://api.bsale.cl/v1/stocks/receptions/{rec_id}/details.json")
        for det in dets:
            sku = det.get('variant', {}).get('code', '').strip().upper()
            if not sku:
                continue
            movimientos.append({
                'fecha':    fecha,
                'sku':      sku,
                'oficina':  nombre,
                'tipo':     'produccion',
                'cantidad': float(det.get('quantity', 0)),
                'doc_id':   rec_id,
            })
        time.sleep(0.1)

    return movimientos

# ── Main ─────────────────────────────────────────────────────

def main():
    hasta = datetime.now()
    desde = hasta - timedelta(days=180)

    # Cargar caché existente
    if os.path.exists(CACHE):
        with open(CACHE) as f:
            cache = json.load(f)
        ultimo = cache.get('ultimo_update', '')
        if ultimo:
            desde = datetime.fromisoformat(ultimo) - timedelta(days=1)
            print(f"Actualizando desde {desde.date()} (incremental)")
        else:
            print("Reconstruyendo historial completo (6 meses)...")
        movs = cache.get('movimientos', [])
        # Eliminar movimientos del período a recargar para evitar duplicados
        movs = [m for m in movs if m['fecha'] < desde.strftime('%Y-%m-%d')]
    else:
        print("Construyendo historial por primera vez (6 meses)...")
        movs = []

    nuevos = []
    for office_id in [1, 3]:
        nuevos += descargar_documentos(office_id, desde, hasta, DOC_VENTA | DOC_DESP)
        nuevos += descargar_recepciones(office_id, desde, hasta)

    movs.extend(nuevos)
    print(f"\nTotal movimientos: {len(movs)} ({len(nuevos)} nuevos)")

    cache = {
        'ultimo_update': hasta.strftime('%Y-%m-%dT%H:%M:%S'),
        'movimientos':   movs,
    }
    with open(CACHE, 'w') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    print(f"Guardado en {CACHE}")

if __name__ == '__main__':
    main()
