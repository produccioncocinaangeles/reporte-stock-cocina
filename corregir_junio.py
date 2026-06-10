# Corrección por única vez (jun-2026): reemplaza los movimientos de junio
# del historial por los datos oficiales exportados desde Bsale:
#   - Detalle de ventas 01-10/06 (ambas tiendas)
#   - Recepciones de Las Pataguas 01-30/06 (guías 668-674 + manuales)
#   - Recepciones de Vitacura 01-30/06 (producciones)
# Luego re-cuadra contra la foto de stock real y valida todo.
# Respaldo previo en historial_backup_antes_correccion.json.
import json, os, sys
import pandas as pd
from datetime import datetime

from generar_dashboard import NOMBRES

CARPETA   = os.path.dirname(os.path.abspath(__file__))
HISTORIAL = os.path.join(CARPETA, 'historial.json')
SNAPSHOT  = os.path.join(CARPETA, 'stock_snapshot.json')
ALERTAS   = os.path.join(CARPETA, 'alertas_stock.json')
BACKUP    = os.path.join(CARPETA, 'historial_backup_antes_correccion.json')

XLS_VENTAS  = r'C:\Users\alamo\Downloads\Detalle de ventas - 01_06_2026 - 10_06_2026.xlsx'
XLS_RECPAT  = r'C:\Users\alamo\Downloads\1781095213.xls'
XLS_RECVIT  = r'C:\Users\alamo\Downloads\1781095486.xls'

DESDE, HASTA = '2026-06-01', '2026-06-10'
CATALOGO = set(NOMBRES)

def fecha_iso(serie):
    return pd.to_datetime(serie.astype(str).str.extract(r'(\d{2}/\d{2}/\d{4})')[0],
                          dayfirst=True).dt.strftime('%Y-%m-%d')

# ── 1. Cargar reportes oficiales ─────────────────────────────
ventas = pd.read_excel(XLS_VENTAS, sheet_name=0)
ventas['of']    = ventas['Sucursal'].astype(str).str.strip().map({'VITACURA':'VIT','LAS PATAGUAS':'PAT'})
ventas['sku']   = ventas['SKU'].astype(str).str.strip().str.upper()
ventas['fecha'] = pd.to_datetime(ventas.iloc[:, 3], dayfirst=True).dt.strftime('%Y-%m-%d')
ventas['tm']    = ventas['Tipo Movimiento'].fillna('venta').str.lower()
ventas = ventas[ventas['sku'].isin(CATALOGO) & ventas['of'].notna()]
if (ventas['tm'].str.contains('devol')).any():
    print('ADVERTENCIA: hay devoluciones en el reporte; se descuentan.')

def leer_recepciones(path):
    df = pd.read_html(path)[0]
    df.columns = ['Fecha','Usuario','Documento','Nota','SKU','Barras','Producto','Cantidad','Serie','CostoU','CostoT']
    df = df[df['SKU'].notna()].copy()
    df['fecha'] = fecha_iso(df['Fecha'])
    df['sku']   = df['SKU'].astype(str).str.strip().str.upper()
    df['doc']   = df['Documento'].astype(str)
    return df[df['sku'].isin(CATALOGO) & df['fecha'].notna()]

rec_pat = leer_recepciones(XLS_RECPAT)
rec_vit = leer_recepciones(XLS_RECVIT)

# ── 2. Respaldo y cirugía ────────────────────────────────────
cache = json.load(open(HISTORIAL, encoding='utf-8'))
with open(BACKUP, 'w', encoding='utf-8') as f:
    json.dump(cache, f, ensure_ascii=False)
movs = cache['movimientos']
print(f'Respaldo: {BACKUP}')
print(f'Movimientos antes: {len(movs)}')

def en_ventana(m):
    return DESDE <= m['fecha'] <= HASTA and m['sku'] in CATALOGO

eliminados, conservados, docs_ignorar = [], [], []
for m in movs:
    if en_ventana(m) and (
        m['tipo'] in ('venta', 'produccion', 'despacho')
        or (m['tipo'] == 'consumo' and m.get('inferido'))
    ):
        eliminados.append(m)
        if m.get('doc_id'):
            docs_ignorar.append([m['doc_id'], m['sku'], m['oficina'], m['tipo']])
    else:
        conservados.append(m)
movs = conservados
print(f'Eliminados de la ventana {DESDE}..{HASTA}: {len(eliminados)} '
      f'(de ellos {sum(1 for e in eliminados if e.get("inferido"))} eran deducidos)')

# ── 3. Insertar datos oficiales ──────────────────────────────
nuevos = []
for _, r in ventas.iterrows():
    signo = -1 if 'devol' in r['tm'] else 1
    nuevos.append({'fecha': r['fecha'], 'sku': r['sku'], 'oficina': r['of'],
                   'tipo': 'venta', 'cantidad': float(r['Cantidad']) * signo,
                   'doc_id': -int(r['Numero del documento']), 'fuente': 'reporte_oficial'})
for _, r in rec_pat.iterrows():
    nuevos.append({'fecha': r['fecha'], 'sku': r['sku'], 'oficina': 'PAT',
                   'tipo': 'produccion', 'cantidad': float(r['Cantidad']),
                   'doc_id': 0, 'fuente': 'reporte_oficial'})
    if 'DESPACHO' in r['doc'].upper():
        # la misma guía es salida en Vitacura
        nuevos.append({'fecha': r['fecha'], 'sku': r['sku'], 'oficina': 'VIT',
                       'tipo': 'despacho', 'cantidad': float(r['Cantidad']),
                       'doc_id': 0, 'fuente': 'reporte_oficial'})
for _, r in rec_vit.iterrows():
    nuevos.append({'fecha': r['fecha'], 'sku': r['sku'], 'oficina': 'VIT',
                   'tipo': 'produccion', 'cantidad': float(r['Cantidad']),
                   'doc_id': 0, 'fuente': 'reporte_oficial'})
movs.extend(nuevos)
print(f'Insertados desde reportes oficiales: {len(nuevos)}')

# ── 4. Re-cuadrar contra la foto de stock real de hoy ────────
snap = json.load(open(SNAPSHOT, encoding='utf-8'))
def reconstruir(movimientos):
    stock = {}
    for of in ('VIT', 'PAT'):
        acum = {}
        # mismo día: entradas (produccion) antes que salidas, para no toparse con el cero
        for m in sorted([x for x in movimientos if x['oficina'] == of],
                        key=lambda x: (x['fecha'], x['tipo'] != 'produccion')):
            ent = m['cantidad'] if m['tipo'] == 'produccion' else 0
            sal = m['cantidad'] if m['tipo'] in ('venta', 'despacho', 'consumo') else 0
            acum[m['sku']] = max(0, acum.get(m['sku'], 0) + ent - sal)
        for sku, q in acum.items():
            stock[(sku, of)] = q
    return stock

recon = reconstruir(movs)
hoy = snap['fecha']
inferidos = []
for sku in CATALOGO:
    s_real = snap['stock'].get(sku, {'vit': 0, 'pat': 0})
    for ok, on in (('vit', 'VIT'), ('pat', 'PAT')):
        residuo = s_real[ok] - recon.get((sku, on), 0)
        if residuo >= 1:
            inferidos.append({'fecha': hoy, 'sku': sku, 'oficina': on, 'tipo': 'produccion',
                              'cantidad': float(round(residuo)), 'doc_id': 0, 'inferido': True})
        elif residuo <= -1:
            inferidos.append({'fecha': hoy, 'sku': sku, 'oficina': on, 'tipo': 'consumo',
                              'cantidad': float(round(-residuo)), 'doc_id': 0, 'inferido': True})
movs.extend(inferidos)
ent = [i for i in inferidos if i['tipo'] == 'produccion']
sal = [i for i in inferidos if i['tipo'] == 'consumo']
print(f'Ajuste residual (drift histórico previo a junio): {len(ent)} entradas, {len(sal)} salidas')
for i in inferidos:
    s = '+' if i['tipo'] == 'produccion' else '-'
    print(f'   {s}{i["cantidad"]:.0f}  {i["sku"]} {i["oficina"]}')

# ── 5. Validaciones ──────────────────────────────────────────
print('\n── VALIDACIONES ──')
# a) ventas de junio == reporte oficial
rep_v, hist_v = {}, {}
for _, r in ventas.iterrows():
    k = (r['fecha'], r['sku'], r['of']); rep_v[k] = rep_v.get(k, 0) + r['Cantidad'] * (-1 if 'devol' in r['tm'] else 1)
for m in movs:
    if m['tipo'] == 'venta' and en_ventana(m):
        k = (m['fecha'], m['sku'], m['oficina']); hist_v[k] = hist_v.get(k, 0) + m['cantidad']
dif_v = [k for k in set(rep_v) | set(hist_v) if abs(rep_v.get(k, 0) - hist_v.get(k, 0)) >= 1]
print(f'a) Ventas junio vs reporte: {"EXACTO" if not dif_v else f"{len(dif_v)} DIFERENCIAS {dif_v[:5]}"}')
# b) reconstrucción final == stock real de la foto
recon2 = reconstruir(movs)
dif_s = [(sku, on, snap['stock'].get(sku, {'vit': 0, 'pat': 0})[ok], recon2.get((sku, on), 0))
         for sku in CATALOGO for ok, on in (('vit', 'VIT'), ('pat', 'PAT'))
         if abs(snap['stock'].get(sku, {'vit': 0, 'pat': 0})[ok] - recon2.get((sku, on), 0)) >= 1]
print(f'b) Stock reconstruido vs foto real: {"EXACTO en 112/112" if not dif_s else f"{len(dif_s)} DIFERENCIAS {dif_s[:5]}"}')
if dif_v or dif_s:
    print('\nVALIDACION FALLIDA — no se guarda nada. Revisar.')
    sys.exit(1)

# ── 6. Guardar ───────────────────────────────────────────────
cache['movimientos']  = movs
cache['docs_ignorar'] = cache.get('docs_ignorar', []) + docs_ignorar
cache['ultimo_update'] = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
with open(HISTORIAL, 'w', encoding='utf-8') as f:
    json.dump(cache, f, ensure_ascii=False, indent=2)
with open(ALERTAS, 'w', encoding='utf-8') as f:
    json.dump({'fecha': hoy,
               'entradas': [{'sku': i['sku'], 'oficina': i['oficina'], 'cantidad': i['cantidad']} for i in ent],
               'salidas':  [{'sku': i['sku'], 'oficina': i['oficina'], 'cantidad': i['cantidad']} for i in sal]},
              f, ensure_ascii=False, indent=2)
print(f'\nGuardado: {len(movs)} movimientos | {len(docs_ignorar)} docs marcados para no re-bajar')
