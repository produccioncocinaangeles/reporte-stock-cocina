import pandas as pd, json, os, requests, time
from datetime import datetime, timedelta

TOKEN   = os.environ.get('BSALE_TOKEN', '')
HEADERS = {'access_token': TOKEN}
CACHE   = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'historial.json')
VIT_XLS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'consolidado_productos_vitacura.xlsx')
PAT_XLS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'consolidado_productos_pataguas.xlsx')

OFICINAS  = {1: 'VIT', 3: 'PAT'}
DOC_VENTA = {35, 6}
DOC_DESP  = {8}

MAPA = {
    'OSTION A LA PARMESANA':'OP','OSTIÓN A LA PARMESANA':'OP','OSTIONES A LA PARMESANA':'OP',
    'ROLLO POLLO PIMENTON':'RPP','ROLLO POLLO PIMENTÓN':'RPP','MINI CHUPE LOCO':'MCL',
    'CHUPE JAIBA GRANDE':'CJ','CHUPE LOCO CAMARON GRANDE':'CLC','CHUPE LOCO CAMARÓN GRANDE':'CLC',
    'MINI CHUPE JAIBA':'MCJ','ROLLO CAMARON':'RC','ROLLO CAMARÓN':'RC',
    'LOMITO ACARAMELADO':'LA','ROLLO MECHADA':'RM','LASAÑA CARNE MECHADA':'LCM',
    'TEQUEÑOS LOMO SALTADO':'TLS','CARPACCIO DE LOCOS':'CL','LOMO SALTADO':'LS',
    'LASAÑA JAIBA CAMARON':'LJC','LASAÑA JAIBA CAMARÓN':'LJC',
    'PURE DE CAMOTE':'PC','PURÉ DE CAMOTE':'PC','MIX MASITAS RELLENAS':'MMR',
    'TARTARO SALMON':'TS','TÁRTARO SALMÓN':'TS',
    'ÑOQUIS CON JAMON SERRANO':'ÑJ','ÑOQUIS CON JAMÓN SERRANO':'ÑJ',
    'TALLARIN SALMON AHUMADO':'TSA','TALLARÍN SALMÓN AHUMADO':'TSA','TALLARINES SALMON AHUMADO':'TSA',
    'ROAST BEEF':'RB','CARPACCIO DE FILETE':'CF','ROLLO SALMON':'RS','ROLLO SALMÓN':'RS',
    'FILETE CHAMPIÑON TOCINO':'FCT','FILETE CHAMPIÑÓN TOCINO':'FCT',
    'PASTEL DE CHOCLO':'PCM','ÑOQUIS PESTO TOMATE CHERRY':'ÑP',
    'TARTARO FILETE':'TF','TÁRTARO FILETE':'TF',
    'MINI CHUPE CAMARON':'MCC','MINI CHUPE CAMARÓN':'MCC','CAMARONES APANADOS':'CAMA',
    'TALLARINES POLLO PIMENTÓN':'TPP','TALLARINES POLLO PIMENTON':'TPP',
    'ROLLO ALCACHOFA':'RA','ARROZ SALVAJE':'AS','MILHOJAS DE PAPAS':'PAC',
    'EMPANADITAS MECHADA':'EM','ÑOQUIS DE CAMARON':'ÑC','ÑOQUIS DE CAMARÓN':'ÑC',
    'EMPANADITAS JAMON SERRANO':'EJ','EMPANADITAS JAMÓN SERRANO':'EJ',
    'BERENJENAS A LA PARMESANA':'BAP','MIGNON DE POLLO':'MIGNON',
    'CAJA DE POSTRES EN VASITO':'PV','CAJA POSTRES EN VASITOS':'PV',
    'TARTARO ATUN':'TA','TÁRTARO ATÚN':'TA','MIX MASITAS RELLENAS DEL MAR':'MMRM',
    'LOMO RELLENO':'LR','CARPACCIO PULPO CON SALSA AL OLIVO':'CP',
    'LASAÑA SALMON':'LSA','LASAÑA SALMÓN':'LSA',
    'ROLLO JAMON SERRANO':'RJ','ROLLO JAMÓN SERRANO':'RJ','CARNE MECHADA':'CM',
    'TALLARIN CARNE MECHADA':'TCM','TALLARÍN CARNE MECHADA':'TCM','TALLARINES CARNE MECHADA':'TCM',
    'PATE CON PERAS':'PCP','PATÉ CON PERAS':'PCP',
    'EMPANADITAS CAMARON':'EC','EMPANADITAS CAMARÓN':'EC','CHUPE CENTOLLA GRANDE':'CC',
    'SALMON CON SALSA DE ALCAPARRAS':'SSA','SALMÓN CON SALSA DE ALCAPARRAS':'SSA',
    'BARQUILLOS':'BAR','LASAÑA MEDITERRANEA':'LM','LASAÑA MEDITERRÁNEA':'LM',
    'CHOCLO A LA CREMA':'CAC','TEQUEÑOS DE PULPO CON SALSA':'TPU',
    'SALSA EN FRASCO':'SA','CAMARONES COCIDOS CON SALSA':'CAMAC',
}

# ── Paso 1: Convertir Excel a movimientos ────────────────────

def excel_a_movimientos(path, oficina):
    df = pd.read_excel(path)
    df['Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True)
    movs = []

    for _, row in df.iterrows():
        nombre = str(row.get('Producto', '')).strip().upper()
        sku = MAPA.get(nombre, nombre)  # convierte nombre a SKU, o usa el nombre si ya es SKU
        if not sku:
            continue
        fecha = row['Fecha'].strftime('%Y-%m-%d')
        mov_sal = str(row.get('Movimiento de salida', '') or '')
        mov_ent = str(row.get('Movimiento de entrada', '') or '')

        if row.get('Salida', 0) > 0:
            if 'DESPACHO' in mov_sal.upper():
                tipo = 'despacho'
            elif 'CONSUMO' in mov_sal.upper():
                tipo = 'consumo'
            else:
                tipo = 'venta'
            movs.append({'fecha': fecha, 'sku': sku, 'oficina': oficina,
                         'tipo': tipo, 'cantidad': float(row['Salida']), 'doc_id': 0})

        if row.get('Entrada', 0) > 0:
            movs.append({'fecha': fecha, 'sku': sku, 'oficina': oficina,
                         'tipo': 'produccion', 'cantidad': float(row['Entrada']), 'doc_id': 0})

    return movs

# ── Paso 2: Descargar días faltantes desde API ───────────────

def get_all(url, params=None):
    items, offset = [], 0
    while True:
        p = {**(params or {}), 'limit': 50, 'offset': offset}
        r = requests.get(url, headers=HEADERS, params=p)
        if r.status_code != 200:
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

def descargar_desde(office_id, desde, hasta):
    nombre = OFICINAS[office_id]
    movs = []

    for dt_id in DOC_VENTA | DOC_DESP:
        tipo = 'venta' if dt_id in DOC_VENTA else 'despacho'
        print(f"  {nombre} tipo {dt_id} ({tipo}) desde {desde.date()}...")
        docs = get_all('https://api.bsale.cl/v1/documents.json', {
            'documenttypeid': dt_id,
            'officeid': office_id,
            'emissiondaterange': f'[{ts(desde)},{ts(hasta)}]',
        })
        print(f"    {len(docs)} documentos")
        for doc in docs:
            fecha = datetime.fromtimestamp(doc['emissionDate']).strftime('%Y-%m-%d')
            dets  = get_all(f"https://api.bsale.cl/v1/documents/{doc['id']}/details.json")
            for det in dets:
                sku = det.get('variant', {}).get('code', '').strip().upper()
                if sku:
                    movs.append({'fecha': fecha, 'sku': sku, 'oficina': nombre,
                                 'tipo': tipo, 'cantidad': float(det.get('quantity', 0)),
                                 'doc_id': doc['id']})
        time.sleep(0.2)

    return movs

# ── Main ─────────────────────────────────────────────────────

def main():
    print("Paso 1: Convirtiendo Excel a historial...")
    movs = []
    movs += excel_a_movimientos(VIT_XLS, 'VIT')
    movs += excel_a_movimientos(PAT_XLS, 'PAT')
    print(f"  {len(movs)} movimientos desde Excel")

    # Fecha más reciente del Excel
    fechas = sorted(set(m['fecha'] for m in movs))
    ultimo_excel = datetime.fromisoformat(fechas[-1])
    desde_api = ultimo_excel + timedelta(days=1)
    hasta = datetime.now()

    print(f"\nPaso 2: Descargando API desde {desde_api.date()} hasta {hasta.date()}...")
    for office_id in [1, 3]:
        nuevos = descargar_desde(office_id, desde_api, hasta)
        movs.extend(nuevos)
        print(f"  +{len(nuevos)} movimientos nuevos")

    print(f"\nTotal: {len(movs)} movimientos")
    cache = {
        'ultimo_update': hasta.strftime('%Y-%m-%dT%H:%M:%S'),
        'movimientos': movs,
    }
    with open(CACHE, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    print(f"Guardado en {CACHE}")

if __name__ == '__main__':
    main()
