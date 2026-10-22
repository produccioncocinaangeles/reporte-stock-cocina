# Actualización diaria del historial — corre en GitHub Actions a las 19:00.
#
# El token de Bsale entrega ventas, guías de despacho, RECEPCIONES, CONSUMOS
# (/v1/stocks/receptions y /consumptions) y el stock ACTUAL. No existe endpoint
# de ajustes/traspasos, así que lo que no tenga documento se sigue deduciendo.
# Este script lleva nuestro propio registro de stock y lo cuadra a diario
# contra el stock real de Bsale:
#
#   1. Baja las ventas y guías nuevas desde la API y las agrega al historial.
#   2. Baja las recepciones y consumos NUEVOS (por id) y los registra como
#      movimientos reales (doc_id 'R<id>' / 'C<id>', con su nota/motivo).
#   3. Reconstruye el stock según NUESTRO sistema (igual que el dashboard).
#   4. Lo compara con el stock real del token, producto por producto; lo que
#      aún no cuadra (sin documento) se deduce como respaldo:
#        - sube sin explicación  -> recepción  (entrada deducida)
#        - baja sin explicación  -> consumo    (salida deducida) => ALERTA
#   5. Guarda historial.json, la foto del día (stock_snapshot.json) y las
#      alertas (alertas_stock.json) que el dashboard muestra en el Resumen.
#
# Tras cada corrida, nuestro stock reconstruido == stock real de Bsale.
# Si la API no responde o no hay token, NO modifica nada. Si solo fallan las
# recepciones/consumos, avisa y se sigue deduciendo como antes.
#
# Modo prueba: SIMULAR=1 hace todo pero NO escribe ningún archivo. Con
# DESDE_RECEPCION_ID / DESDE_CONSUMO_ID (solo en simulación) se sobrescribe
# el último id procesado, para ver qué entraría desde ese punto.
#
# Protecciones (para que un error de Bsale no ensucie el historial):
#   - Tope de movimientos inferidos y cobertura mínima de SKUs (ver abajo).
#   - Guardado seguro: se escribe a un archivo temporal y recién después se
#     reemplaza el original (os.replace); el historial nunca puede perder datos.
#   - Para saltarse los topes (cuando el dueño confirmó que el cambio de stock
#     fue real) correr con la variable de entorno FORZAR_INFERENCIA=1.
import json, os, time, sys
from datetime import datetime, timedelta, timezone
import requests

# GitHub Actions corre en hora UTC: a las 19:00 de Chile ya son las 23:00 UTC,
# y con el retraso habitual del plan gratuito la corrida cae en el día siguiente
# UTC, marcando despachos/recepciones con fecha de mañana. Fijamos hora chilena.
try:
    from zoneinfo import ZoneInfo
    TZ_CHILE = ZoneInfo('America/Santiago')
except Exception:
    TZ_CHILE = None  # sin base de datos de zonas horarias: usa hora local

from generar_dashboard import NOMBRES, bsale_stock, CARPETA, BSALE_TOKEN, expandir_componentes

ARCHIVO_HISTORIAL = os.path.join(CARPETA, 'historial.json')
ARCHIVO_SNAPSHOT  = os.path.join(CARPETA, 'stock_snapshot.json')
ARCHIVO_ALERTAS   = os.path.join(CARPETA, 'alertas_stock.json')

# Tope de movimientos inferidos por corrida. Sale del historial real (113 días
# con inferidos): máximo diario 29, percentil 95 = 22. Se fija en 60, más del
# doble del máximo normal, pero por debajo de lo que pasaría si Bsale devolviera
# el stock vacío o parcial (55 SKUs x 2 oficinas = hasta ~110 movimientos).
MAX_INFERIDOS_POR_CORRIDA = 60
# Si el stock descargado cubre menos de esta fracción de los SKUs de la foto
# anterior, Bsale devolvió datos incompletos y se aborta.
COBERTURA_MINIMA_SKUS = 0.80

HEADERS   = {'access_token': BSALE_TOKEN}
OFICINAS  = {1: 'VIT', 3: 'PAT'}
DOC_VENTA = {35, 6}
DOC_DESP  = {8}

def get_all(url, params=None):
    items, offset = [], 0
    while True:
        p = {**(params or {}), 'limit': 50, 'offset': offset}
        data = None
        for intento in range(3):
            try:
                r = requests.get(url, headers=HEADERS, params=p, timeout=30)
                r.raise_for_status()
                data = r.json()
                break
            except Exception as e:
                print(f"  get_all intento {intento+1}/3: {e}")
                if intento < 2:
                    print("  Reintentando en 30 s...")
                    time.sleep(30)
        if data is None:
            raise Exception(f"get_all: no se pudo obtener {url} después de 3 intentos")
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
                # Bsale entrega emissionDate como medianoche UTC del día de emisión.
                # Convertir en hora local (Chile, UTC-4) retrocede la fecha un día,
                # por eso se interpreta SIEMPRE en UTC.
                fecha = datetime.fromtimestamp(doc['emissionDate'], tz=timezone.utc).strftime('%Y-%m-%d')
                dets  = get_all(f"https://api.bsale.cl/v1/documents/{doc['id']}/details.json")
                for det in dets:
                    sku = str(det.get('variant', {}).get('code', '')).strip().upper()
                    if sku:
                        movs.append({'fecha': fecha, 'sku': sku, 'oficina': nombre,
                                     'tipo': tipo, 'cantidad': float(det.get('quantity', 0)),
                                     'doc_id': doc['id']})
            time.sleep(0.2)
    return movs

SIMULAR = os.environ.get('SIMULAR') == '1'
BASE = 'https://api.bsale.cl/v1'

def get_pagina(url, params):
    # Una sola página, con 3 reintentos.
    for intento in range(3):
        try:
            r = requests.get(url, headers=HEADERS, params=params, timeout=60)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            print(f"  pagina intento {intento+1}/3: {e}")
            if intento < 2:
                time.sleep(5)
    raise Exception(f"no se pudo obtener {url} (3 intentos)")

def mapa_variantes():
    # id de variante -> SKU. El detalle de recepciones/consumos trae la variante
    # sin "code" (solo id), así que hay que resolverlo.
    mapa = {}
    for v in get_all(f'{BASE}/variants.json'):
        code = str(v.get('code', '')).strip().upper()
        if code:
            mapa[str(v['id'])] = code
    return mapa

def docs_nuevos(endpoint, ultimo_id):
    # Estrategia: SIN filtro de oficina (un solo recorrido sirve a ambas; la
    # oficina viene en cada documento y se descartan las que no están en
    # OFICINAS). La API devuelve los ids en orden creciente y las recepciones
    # ignoran filtros de fecha, así que: count con limit=1 y se leen páginas de
    # 50 desde el final hacia atrás hasta topar un id <= ultimo_id. Después se
    # lee la cola por si llegó algo entre el count y la lectura.
    url = f'{BASE}/stocks/{endpoint}.json'
    count = get_pagina(url, {'limit': 1})['count']
    vistos = {}
    offset = max(0, count - 50)
    while True:
        items = get_pagina(url, {'limit': 50, 'offset': offset, 'expand': '[details]'}).get('items', [])
        for it in items:
            vistos[it['id']] = it
        if offset == 0 or (items and min(i['id'] for i in items) <= ultimo_id):
            break
        offset = max(0, offset - 50)
    for it in get_pagina(url, {'limit': 50, 'offset': count, 'expand': '[details]'}).get('items', []):
        vistos[it['id']] = it
    max_id = max(vistos) if vistos else ultimo_id
    return [vistos[i] for i in sorted(vistos) if i > ultimo_id], max_id

def id_maximo(endpoint):
    # Id del último documento existente ahora (count con limit=1 y última página).
    url = f'{BASE}/stocks/{endpoint}.json'
    total = get_pagina(url, {'limit': 1})['count']
    ultimos = get_pagina(url, {'limit': 50, 'offset': max(0, total - 50)}).get('items', [])
    return max(i['id'] for i in ultimos)

def detalles_de(doc, endpoint):
    # expand=[details] trunca en 25 líneas; si hay más se piden completas.
    bloque = doc.get('details', {})
    items = bloque.get('items', [])
    if bloque.get('count', len(items)) > len(items):
        items = get_all(f"{BASE}/stocks/{endpoint}/{doc['id']}/details.json")
    return items

def recepciones_y_consumos(cache, topes):
    # Devuelve (movimientos nuevos, nuevos ids {clave: id}). Lanza excepción si
    # algo falla: main() la captura y NO actualiza los ids.
    # topes = {endpoint: id máximo leído ANTES de pedir el stock real}. Solo se
    # importan documentos hasta ese id: uno creado mientras corre el proceso
    # (p.ej. Pataguas recibe una guía a mediodía) ya está en la lista de
    # recepciones pero no en el stock que se leyó antes, y la cuadratura lo
    # anulaba con una salida "sin documento". Queda para la corrida siguiente.
    movs, nuevos_ids, variantes = [], {}, None
    for clave, endpoint, prefijo, env in (
            ('ultima_recepcion_id', 'receptions', 'R', 'DESDE_RECEPCION_ID'),
            ('ultimo_consumo_id', 'consumptions', 'C', 'DESDE_CONSUMO_ID')):
        ultimo = cache.get(clave)
        if SIMULAR and os.environ.get(env):
            ultimo = int(os.environ[env])
        if ultimo is None:
            # Primera corrida: corte. Lo anterior ya está deducido en el
            # historial; importarlo lo duplicaría. Solo se fija el id máximo.
            max_id = topes[endpoint]
            nuevos_ids[clave] = max_id
            print(f"  {endpoint}: primera corrida, corte en id {max_id} (no se importa nada)")
            continue
        docs, max_id = docs_nuevos(endpoint, ultimo)
        tope = topes[endpoint]
        docs = [d for d in docs if d['id'] <= tope]
        max_id = max(ultimo, min(max_id, tope))
        for doc in docs:
            oficina = OFICINAS.get(int(doc.get('office', {}).get('id', 0)))
            if not oficina:
                continue  # Providencia/Zapallar (cerradas) u otra
            # updateStock NO sirve de filtro (verificado con las últimas 400
            # recepciones y 400 consumos): casi todo viene en 0 y aun así
            # variantStock (stock posterior) refleja el movimiento; =1 solo marca
            # los "Ajuste de Stock", que también movieron stock. Se registran todos.
            nota = (doc.get('note') or '').strip()
            if prefijo == 'R':
                fecha = doc.get('rawAdmissionDate', '')
            else:
                fecha = datetime.fromtimestamp(doc['consumptionDate'], tz=timezone.utc).strftime('%Y-%m-%d')
            for det in detalles_de(doc, endpoint):
                variant = det.get('variant', {}) or {}
                sku = str(variant.get('code', '')).strip().upper()
                if not sku:
                    if variantes is None:
                        variantes = mapa_variantes()
                    sku = variantes.get(str(variant.get('id', '')), '')
                if not sku:
                    continue
                m = {'fecha': fecha, 'sku': sku, 'oficina': oficina,
                     'tipo': 'produccion' if prefijo == 'R' else 'consumo',
                     'cantidad': float(det.get('quantity', 0)),
                     'doc_id': f"{prefijo}{doc['id']}"}
                if prefijo == 'R':
                    m['costo'] = float(det.get('cost', 0) or 0)
                    if nota:
                        m['nota'] = nota
                else:
                    m['nota'] = nota
                movs.append(m)
        nuevos_ids[clave] = max_id
    return movs, nuevos_ids

def reconstruir_stock(movs):
    # Réplica exacta de leer_json() del dashboard: por oficina, movimientos
    # ordenados por fecha (mismo día: entradas antes que salidas), stock nunca negativo.
    stock = {}
    for oficina in ('VIT', 'PAT'):
        acum = {}
        for m in sorted([x for x in expandir_componentes(movs) if x['oficina'] == oficina],
                        key=lambda x: (x['fecha'], x['tipo'] != 'produccion')):
            ent = m['cantidad'] if m['tipo'] == 'produccion' else 0
            sal = m['cantidad'] if m['tipo'] in ('venta', 'despacho', 'consumo') else 0
            acum[m['sku']] = max(0, acum.get(m['sku'], 0) + ent - sal)
        for sku, q in acum.items():
            stock[(sku, oficina)] = q
    return stock

def guardar_json_seguro(ruta, datos, min_movimientos=None):
    # Escribe a un temporal en la misma carpeta y reemplaza el original con
    # os.replace. Si min_movimientos viene dado (historial), valida antes que
    # el temporal se pueda releer y que no tenga menos movimientos que el original.
    if SIMULAR:
        return
    tmp = ruta + '.tmp'
    try:
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
        if min_movimientos is not None:
            with open(tmp, encoding='utf-8') as f:
                releido = json.load(f)
            if len(releido['movimientos']) < min_movimientos:
                raise ValueError(f"el archivo nuevo tiene {len(releido['movimientos'])} "
                                 f"movimientos y el original {min_movimientos}")
        os.replace(tmp, ruta)
    except Exception as e:
        try:
            os.remove(tmp)
        except OSError:
            pass
        print(f"ERROR: no se pudo guardar {os.path.basename(ruta)} de forma segura ({e}). "
              f"El archivo original quedó intacto.")
        sys.exit(1)

def verificar_topes(inferidos, stock_real, skus_foto_anterior):
    # Aborta (sys.exit(1)) si el stock de Bsale parece sospechoso. Se salta con FORZAR_INFERENCIA=1.
    if SIMULAR:
        print(f"  (simulación: topes solo informativos; {len(inferidos)} inferidos, máx. normal {MAX_INFERIDOS_POR_CORRIDA})")
        return
    if os.environ.get('FORZAR_INFERENCIA') == '1':
        print("  FORZAR_INFERENCIA=1: se omiten los topes de seguridad.")
        return
    if skus_foto_anterior:
        cubiertos = sum(1 for sku in skus_foto_anterior
                        if sku in stock_real or sku.replace('Ñ', 'N') in stock_real)
        cobertura = cubiertos / len(skus_foto_anterior)
        if cobertura < COBERTURA_MINIMA_SKUS:
            print(f"ERROR: Bsale devolvió stock incompleto: solo {cubiertos} de "
                  f"{len(skus_foto_anterior)} SKUs de la foto anterior ({cobertura:.0%}; mínimo "
                  f"{COBERTURA_MINIMA_SKUS:.0%}). No se guardó nada. "
                  f"Si el cambio es real, correr con FORZAR_INFERENCIA=1.")
            sys.exit(1)
    if len(inferidos) > MAX_INFERIDOS_POR_CORRIDA:
        print(f"ERROR: Bsale devolvió un stock sospechoso: {len(inferidos)} movimientos "
              f"inferidos (normal: hasta {MAX_INFERIDOS_POR_CORRIDA}). No se guardó nada. "
              f"Si el cambio es real, correr con FORZAR_INFERENCIA=1.")
        sys.exit(1)

def main():
    hoy = datetime.now(TZ_CHILE)
    print(f"Actualización diaria — {hoy:%Y-%m-%d %H:%M} (hora Chile)")

    # ── Topes de recepciones/consumos, ANTES del stock real (ver recepciones_y_consumos) ──
    topes = None
    try:
        topes = {ep: id_maximo(ep) for ep in ('receptions', 'consumptions')}
    except Exception as e:
        print(f"AVISO: no se pudieron leer los topes de recepciones/consumos ({e}).")

    # ── Stock real (si falla, abortamos sin tocar nada) ──
    stock_real = bsale_stock()
    if not stock_real:
        print("ERROR: sin datos de stock desde Bsale (token/API). No se modifica nada.")
        sys.exit(1)
    print(f"  Stock real: {len(stock_real)} SKUs")

    cache = json.load(open(ARCHIVO_HISTORIAL, encoding='utf-8'))
    movs  = cache['movimientos']
    n_original = len(movs)  # para validar que al guardar no se pierdan movimientos
    conocidos = {(m['doc_id'], m['sku'], m['oficina'], m['tipo'])
                 for m in movs if m.get('doc_id')}
    # Documentos reemplazados por datos manuales (correcciones): no re-agregar
    conocidos |= {tuple(k) for k in cache.get('docs_ignorar', [])}

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

    # ── 1b. Recepciones y consumos nuevos (documentos reales) ──
    # Si falla, NO se aborta: la cuadratura los deduce como siempre.
    reales, nuevos_ids = [], {}
    try:
        if topes is None:
            raise Exception('sin topes')
        reales, nuevos_ids = recepciones_y_consumos(cache, topes)
    except Exception as e:
        print(f"AVISO: no se pudieron leer recepciones/consumos ({e}); se sigue deduciendo como antes.")
    recepciones, consumos = [], []
    for m in reales:
        clave = (m['doc_id'], m['sku'], m['oficina'], m['tipo'])
        if clave in conocidos:
            continue
        conocidos.add(clave)
        movs.append(m)
        (recepciones if m['tipo'] == 'produccion' else consumos).append(m)
    print(f"  Recepciones nuevas: {len(recepciones)} | consumos nuevos: {len(consumos)}")
    for m in recepciones:
        print(f"    recepción: {m['fecha']} {m['sku']} {m['oficina']} +{m['cantidad']:.0f} {m.get('nota', '')}")
    for m in consumos:
        print(f"    consumo:   {m['fecha']} {m['sku']} {m['oficina']} -{m['cantidad']:.0f} [{m['nota']}]")

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
    # Tope de seguridad ANTES de tocar nada (la foto anterior aún no se sobrescribe)
    try:
        with open(ARCHIVO_SNAPSHOT, encoding='utf-8') as f:
            skus_foto_anterior = list(json.load(f).get('stock', {}).keys())
    except Exception:
        skus_foto_anterior = []
    verificar_topes(inferidos, stock_real, skus_foto_anterior)
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
    cache.update(nuevos_ids)  # ids ya procesados (vacío si la API falló)
    guardar_json_seguro(ARCHIVO_HISTORIAL, cache, min_movimientos=n_original)

    foto = {sku: stock_real.get(sku, stock_real.get(sku.replace('Ñ', 'N'), {'vit': 0, 'pat': 0}))
            for sku in NOMBRES}
    guardar_json_seguro(ARCHIVO_SNAPSHOT, {'fecha': fecha_hoy, 'stock': foto})

    guardar_json_seguro(ARCHIVO_ALERTAS, {
        'fecha': fecha_hoy,
        'entradas': [{'sku': i['sku'], 'oficina': i['oficina'], 'cantidad': i['cantidad']} for i in entradas],
        'salidas':  [{'sku': i['sku'], 'oficina': i['oficina'], 'cantidad': i['cantidad']} for i in salidas],
        'recepciones': [{'sku': m['sku'], 'oficina': m['oficina'], 'cantidad': m['cantidad'],
                         'nota': m.get('nota', '')} for m in recepciones],
        'consumos':    [{'sku': m['sku'], 'oficina': m['oficina'], 'cantidad': m['cantidad'],
                         'nota': m.get('nota', '')} for m in consumos]})

    print(f"Historial: {len(movs)} movimientos | " +
          ("SIMULAR=1: no se escribió ningún archivo" if SIMULAR else "Foto y alertas guardadas"))

if __name__ == '__main__':
    main()
