#!/usr/bin/env python3
"""
La Cocina — Generador de Dashboard v2
Lee consolidados Vitacura/Pataguas + stock Bsale API
Genera dashboard.html con datos embebidos
"""
import os, json, requests, math, time
import pandas as pd
import numpy as np
from collections import Counter

# ─── CONFIGURACIÓN ───────────────────────────────────────────
CARPETA       = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_VIT   = os.path.join(CARPETA, 'consolidado_productos_vitacura.xlsx')
ARCHIVO_PAT   = os.path.join(CARPETA, 'consolidado_productos_pataguas.xlsx')
ARCHIVO_JSON  = os.path.join(CARPETA, 'historial.json')
ARCHIVO_HTML  = os.path.join(CARPETA, 'dashboard.html')
BSALE_TOKEN   = os.environ.get('BSALE_TOKEN', '')
# Hora de Chile siempre: en GitHub Actions (UTC) la corrida nocturna cae en el
# día siguiente UTC y desplazaba la fecha del dashboard y los cálculos de venta.
_HOY          = pd.Timestamp.now(tz='America/Santiago').normalize().tz_localize(None)
FECHA_HOY     = _HOY
FECHA_STR     = _HOY.strftime('%d/%m/%Y')

# ─── CATÁLOGO ────────────────────────────────────────────────
NOMBRES = {
    'OP':'OSTIÓN A LA PARMESANA','RPP':'ROLLO POLLO PIMENTÓN','MCL':'MINI CHUPE LOCO',
    'CJ':'CHUPE JAIBA GRANDE','CLC':'CHUPE LOCO CAMARÓN GRANDE','MCJ':'MINI CHUPE JAIBA',
    'RC':'ROLLO CAMARÓN','LA':'LOMITO ACARAMELADO','RM':'ROLLO MECHADA',
    'LCM':'LASAÑA CARNE MECHADA','TLS':'TEQUEÑOS LOMO SALTADO','CL':'CARPACCIO DE LOCOS',
    'LS':'LOMO SALTADO','LJC':'LASAÑA JAIBA CAMARÓN','PC':'PURÉ DE CAMOTE',
    'MMR':'MIX MASITAS RELLENAS','TS':'TÁRTARO SALMÓN','ÑJ':'ÑOQUIS CON JAMÓN SERRANO',
    'TSA':'TALLARÍN SALMÓN AHUMADO','RB':'ROAST BEEF','CF':'CARPACCIO DE FILETE',
    'RS':'ROLLO SALMÓN','FCT':'FILETE CHAMPIÑÓN TOCINO','PCM':'PASTEL DE CHOCLO',
    'ÑP':'ÑOQUIS PESTO TOMATE CHERRY','TF':'TÁRTARO FILETE','MCC':'MINI CHUPE CAMARÓN',
    'CAMA':'CAMARONES APANADOS','TPP':'TALLARINES POLLO PIMENTÓN','RA':'ROLLO ALCACHOFA',
    'AS':'ARROZ SALVAJE','PAC':'MILHOJAS DE PAPAS','EM':'EMPANADITAS MECHADA',
    'ÑC':'ÑOQUIS DE CAMARÓN','EJ':'EMPANADITAS JAMÓN SERRANO','BAP':'BERENJENAS A LA PARMESANA',
    'MIGNON':'MIGNON DE POLLO','PV':'CAJA DE POSTRES EN VASITO','TA':'TÁRTARO ATÚN',
    'MMRM':'MIX MASITAS RELLENAS DEL MAR','LR':'LOMO RELLENO',
    'CP':'CARPACCIO PULPO CON SALSA AL OLIVO','LSA':'LASAÑA SALMÓN','RJ':'ROLLO JAMÓN SERRANO',
    'CM':'CARNE MECHADA','TCM':'TALLARÍN CARNE MECHADA','PCP':'PATÉ CON PERAS',
    'EC':'EMPANADITAS CAMARÓN','CC':'CHUPE CENTOLLA GRANDE',
    'SSA':'SALMÓN CON SALSA DE ALCAPARRAS','BAR':'BARQUILLOS','LM':'LASAÑA MEDITERRÁNEA',
    'CAC':'CHOCLO A LA CREMA','TPU':'TEQUEÑOS DE PULPO CON SALSA','SA':'SALSA EN FRASCO',
    'CAMAC':'CAMARONES COCIDOS CON SALSA',
}
COCINEROS = {
    'OP':'CAROLINA','RPP':'CÉSAR','MCL':'JESÚS','CJ':'JESÚS','CLC':'JESÚS',
    'MCJ':'JESÚS','RC':'CÉSAR','LA':'CÉSAR','RM':'CÉSAR','LCM':'ADRIANA',
    'TLS':'ADRIANA','CL':'JESÚS','LS':'CÉSAR','LJC':'ADRIANA','PC':'CAROLINA',
    'MMR':'ADRIANA','TS':'CÉSAR','ÑJ':'CAROLINA','TSA':'ADRIANA','RB':'JESÚS',
    'CF':'JESÚS','RS':'CÉSAR','FCT':'JESÚS','PCM':'CAROLINA','ÑP':'CAROLINA',
    'TF':'CÉSAR','MCC':'JESÚS','CAMA':'ADRIANA','TPP':'ADRIANA','RA':'CÉSAR',
    'AS':'JESÚS','PAC':'CÉSAR','EM':'CAROLINA','ÑC':'CAROLINA','EJ':'CAROLINA',
    'BAP':'ADRIANA','MIGNON':'CÉSAR','PV':'CAROLINA','TA':'CÉSAR','MMRM':'ADRIANA',
    'LR':'JESÚS','CP':'JESÚS','LSA':'ADRIANA','RJ':'CÉSAR','CM':'CÉSAR',
    'TCM':'ADRIANA','PCP':'JESÚS','EC':'CAROLINA','CC':'JESÚS','SSA':'JESÚS',
    'BAR':'CAROLINA','LM':'ADRIANA','CAC':'JESÚS','TPU':'ADRIANA','SA':'ADRIANA',
    'CAMAC':'CÉSAR',
}
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

# ─── LECTURA ─────────────────────────────────────────────────
def leer_json(oficina):
    with open(ARCHIVO_JSON, encoding='utf-8') as f:
        cache = json.load(f)
    movs = sorted(
        [m for m in cache['movimientos'] if m['oficina'] == oficina],
        # mismo día: entradas (produccion) antes que salidas, para no toparse con el cero
        key=lambda m: (m['fecha'], m['tipo'] != 'produccion')
    )
    # Calcular stock acumulado por SKU (entradas - salidas, mínimo 0)
    stock_acum = {}
    rows = []
    for m in movs:
        sku     = m['sku']
        entrada = m['cantidad'] if m['tipo'] == 'produccion' else 0
        salida  = m['cantidad'] if m['tipo'] in ('venta', 'despacho', 'consumo') else 0
        stock_acum[sku] = max(0, stock_acum.get(sku, 0) + entrada - salida)
        mov_sal = 'GUÍA DE DESPACHO' if m['tipo'] == 'despacho' else ('Consumo' if m['tipo'] == 'consumo' else 'BOLETA')
        mov_ent = 'Recepción' if m['tipo'] == 'produccion' else ''
        rows.append({
            'Fecha':                  pd.Timestamp(m['fecha']),
            'SKU':                    m['sku'],
            'Entrada':                entrada,
            'Salida':                 salida,
            'Stock':                  stock_acum[sku],
            'Movimiento de salida':   mov_sal if salida > 0 else '',
            'Movimiento de entrada':  mov_ent if entrada > 0 else '',
        })
    return pd.DataFrame(rows)

def leer(archivo):
    df = pd.read_excel(archivo)
    df['Fecha']   = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce')
    df['Entrada'] = pd.to_numeric(df['Entrada'], errors='coerce').fillna(0)
    df['Salida']  = pd.to_numeric(df['Salida'],  errors='coerce').fillna(0)
    df['Stock']   = pd.to_numeric(df['Stock'],   errors='coerce').fillna(0)
    df['SKU']     = df['Producto'].str.strip().map(MAPA)
    return df[df['SKU'].notna()].copy()

def bsale_stock():
    headers = {'access_token': BSALE_TOKEN}
    for intento in range(3):
        result, offset = {}, 0
        try:
            while True:
                r = requests.get('https://api.bsale.cl/v1/stocks.json', headers=headers,
                                 params={'limit':50,'offset':offset,'expand':'[variant,office]'}, timeout=30)
                if r.status_code != 200:
                    break
                data = r.json()
                for item in data.get('items', []):
                    try:
                        sku      = ' '.join(str(item['variant']['code']).split()).upper()
                        sucursal = ' '.join(str(item['office']['name']).split()).upper()
                        qty      = int(item.get('quantityAvailable', 0) or 0)
                        if sku not in result:
                            result[sku] = {'vit':0,'pat':0}
                        if   sucursal == 'VITACURA':     result[sku]['vit'] += qty
                        elif sucursal == 'LAS PATAGUAS': result[sku]['pat'] += qty
                    except: pass
                if offset + 50 >= data.get('count', 0): break
                offset += 50
            if result:
                return result
            print(f"  bsale_stock: respuesta vacía (intento {intento+1}/3)")
        except Exception as e:
            print(f"  bsale_stock intento {intento+1}/3 falló: {e}")
        if intento < 2:
            print("  Reintentando en 30 s...")
            time.sleep(30)
    return {}

# ─── CÁLCULOS ────────────────────────────────────────────────
def stock_diario(df_sku):
    if len(df_sku) == 0: return {}
    dias = pd.date_range(df_sku['Fecha'].min(), FECHA_HOY)
    serie = df_sku.groupby('Fecha')['Stock'].last().reindex(dias, method='ffill').fillna(0)
    return serie.to_dict()

def periodos_sin_stock(stock_d):
    if not stock_d: return []
    resultado, inicio = [], None
    for d in sorted(stock_d):
        if stock_d[d] <= 0:
            if inicio is None: inicio = d
        else:
            if inicio is not None:
                dur = (d - inicio).days
                if dur > 0:
                    resultado.append({'inicio': inicio.strftime('%d/%m/%Y'),
                                      'fin': d.strftime('%d/%m/%Y'), 'dias': dur})
                inicio = None
    return resultado

def moda(lst):
    return Counter(lst).most_common(1)[0][0] if lst else 0

def mediana(lst):
    if not lst: return 0
    s = sorted(lst); n = len(s)
    m = s[n//2] if n % 2 else (s[n//2 - 1] + s[n//2]) / 2
    return max(1, round(m))

def velocidad(df_sku, stock_d):
    # Ventana: 3 meses completos anteriores al mes en curso (~90 días).
    # Más reactiva a productos que cambian de ritmo (crecen o caen) que un
    # promedio de 6 meses, que los arrastra con datos viejos.
    # Excluye meses con quiebre de stock (< 7 días con stock disponible)
    # y el mes en curso (incompleto, sesga la velocidad).
    UMBRAL_DIAS = 7
    mes_actual  = pd.Period(FECHA_HOY, 'M')
    fecha_hist  = (mes_actual - 3).start_time

    mask_v = (df_sku['Salida'] > 0) & (~df_sku['Movimiento de salida'].str.contains(
        'GUÍA DE DESPACHO|Guía de Despacho|Consumo', na=False))

    # Días con stock por mes (ventana de 3 meses, sin el mes en curso)
    dias_mes = {}
    for d, v in stock_d.items():
        if d >= fecha_hist and v > 0:
            mes = pd.Period(d, 'M')
            if mes == mes_actual:
                continue
            dias_mes[mes] = dias_mes.get(mes, 0) + 1

    # Ventas por mes (últimos 6 meses)
    df_v = df_sku[mask_v & (df_sku['Fecha'] >= fecha_hist)].copy()
    ventas_mes = {}
    if len(df_v) > 0:
        df_v['mes'] = df_v['Fecha'].dt.to_period('M')
        ventas_mes = df_v.groupby('mes')['Salida'].sum().to_dict()

    # Meses válidos: al menos UMBRAL_DIAS días con stock
    meses_validos = [(mes, dias) for mes, dias in dias_mes.items() if dias >= UMBRAL_DIAS]

    if meses_validos:
        total_v = sum(ventas_mes.get(mes, 0) for mes, _ in meses_validos)
        total_d = sum(dias for _, dias in meses_validos)
        vel = round(total_v / total_d, 4) if total_d > 0 else 0.0
    else:
        total_v, total_d, vel = 0, 0, 0.0

    # Para display: los MISMOS números que producen la velocidad
    # (ventas y días con stock de los meses válidos), así el análisis cuadra.
    return vel, int(total_v), int(total_d)

def lotes(df_sku):
    ent = df_sku[df_sku['Entrada'] > 0].sort_values('Fecha')
    if len(ent) == 0: return [], 0, 0
    result = []
    for _, row in ent.iterrows():
        sig  = ent[ent['Fecha'] > row['Fecha']]['Fecha'].min()
        if pd.isna(sig): sig = FECHA_HOY
        mask = (df_sku['Fecha'] >= row['Fecha']) & (df_sku['Fecha'] < sig)
        ventas = int(df_sku[mask & (df_sku['Salida'] > 0) &
            (~df_sku['Movimiento de salida'].str.contains(
                'GUÍA DE DESPACHO|Guía de Despacho|Consumo', na=False))]['Salida'].sum())
        doc = str(row.get('Movimiento de entrada', ''))
        num = doc.split('Nº')[-1].strip() if 'Nº' in doc else doc[:20]
        result.append({'fecha': row['Fecha'].strftime('%d/%m/%Y'),
                       'cantidad': int(row['Entrada']), 'documento': num,
                       'ventas_posteriores': ventas,
                       'dias_hasta_siguiente': (sig - row['Fecha']).days})
    tamaños = [l['cantidad'] for l in result]
    return result, round(float(np.mean(tamaños)), 1), len(result)

def ventas_mes(df_sku):
    mask = (df_sku['Salida'] > 0) & (~df_sku['Movimiento de salida'].str.contains(
        'GUÍA DE DESPACHO|Guía de Despacho|Consumo', na=False))
    v = df_sku[mask].copy()
    if len(v) == 0: return {}
    v['mes'] = v['Fecha'].dt.to_period('M')
    return {str(k): int(val) for k, val in v.groupby('mes')['Salida'].sum().items()}

def movimientos(df_vit, df_pat):
    MESES = {'01':'Ene','02':'Feb','03':'Mar','04':'Abr','05':'May','06':'Jun',
             '07':'Jul','08':'Ago','09':'Sep','10':'Oct','11':'Nov','12':'Dic'}
    def fmt(ts):
        d = ts.strftime('%d/%m/%Y').split('/')
        return f"{int(d[0])} {MESES[d[1]]}"

    movs = []
    for _, row in df_vit.sort_values('Fecha').iterrows():
        de = str(row['Movimiento de entrada']) if pd.notna(row['Movimiento de entrada']) else ''
        ds = str(row['Movimiento de salida'])  if pd.notna(row['Movimiento de salida'])  else ''
        if row['Entrada'] > 0:
            t = 'Producción' if ('Recepción' in de or 'Sin Documento' in de) else 'Entrada'
            n = de.split('Nº')[-1].strip() if 'Nº' in de else de[:20]
            movs.append({'fecha':fmt(row['Fecha']),'fecha_ord':row['Fecha'].isoformat(),
                         'tipo':t,'documento':n,'tienda':'Vitacura',
                         'cantidad':int(row['Entrada']),'signo':'+','stock':int(row['Stock'])})
        if row['Salida'] > 0:
            if   'GUÍA DE DESPACHO' in ds.upper(): t,n = 'Despacho', ds.split('Nº')[-1].strip() if 'Nº' in ds else ds[:20]
            elif 'Consumo' in ds:                  t,n = 'Consumo',  ds.split('Nº')[-1].strip() if 'Nº' in ds else ds[:20]
            else:                                  t,n = 'Venta',    ds.replace('BOLETA ELECTRÓNICA T Nº','').replace('BOLETA ELECTRÓNICA','').strip()
            movs.append({'fecha':fmt(row['Fecha']),'fecha_ord':row['Fecha'].isoformat(),
                         'tipo':t,'documento':n,'tienda':'Vitacura',
                         'cantidad':int(row['Salida']),'signo':'-','stock':int(row['Stock'])})

    for _, row in df_pat.sort_values('Fecha').iterrows():
        de = str(row['Movimiento de entrada']) if pd.notna(row['Movimiento de entrada']) else ''
        ds = str(row['Movimiento de salida'])  if pd.notna(row['Movimiento de salida'])  else ''
        if row['Entrada'] > 0:
            n = de.split('Nº')[-1].strip() if 'Nº' in de else de[:20]
            movs.append({'fecha':fmt(row['Fecha']),'fecha_ord':row['Fecha'].isoformat(),
                         'tipo':'Despacho recibido','documento':n,'tienda':'Pataguas',
                         'cantidad':int(row['Entrada']),'signo':'+','stock':int(row['Stock'])})
        if row['Salida'] > 0:
            t = 'Consumo' if 'Consumo' in ds else 'Venta'
            n = ds.replace('BOLETA ELECTRÓNICA T Nº','').replace('BOLETA ELECTRÓNICA','').strip()
            movs.append({'fecha':fmt(row['Fecha']),'fecha_ord':row['Fecha'].isoformat(),
                         'tipo':t,'documento':n,'tienda':'Pataguas',
                         'cantidad':int(row['Salida']),'signo':'-','stock':int(row['Stock'])})

    movs.sort(key=lambda x: x['fecha_ord'])
    return movs

# ─── PIPELINE ────────────────────────────────────────────────
def procesar():
    if os.path.exists(ARCHIVO_JSON):
        print("Leyendo historial.json...")
        vit = leer_json('VIT')
        pat = leer_json('PAT')
    else:
        print("Leyendo consolidados Excel...")
        vit = leer(ARCHIVO_VIT)
        pat = leer(ARCHIVO_PAT)

    print("Obteniendo stock desde Bsale...")
    sb = bsale_stock()
    if not sb:
        print("  Warning: sin datos Bsale")
    else:
        print(f"  {len(sb)} productos con stock desde Bsale")

    resultados = []
    for sku in NOMBRES:
        dv = vit[vit['SKU']==sku].copy()
        dp = pat[pat['SKU']==sku].copy()

        # Stock actual
        s = sb.get(sku, {'vit':0,'pat':0})
        if s['vit']==0 and s['pat']==0:
            s = sb.get(sku.replace('Ñ','N'), {'vit':0,'pat':0})
        sv_hoy, sp_hoy = s['vit'], s['pat']

        # Stock diario
        sd_v = stock_diario(dv)
        sd_p = stock_diario(dp)

        # Períodos sin stock — SOLO VITACURA para tiempo de reposición
        per_v = periodos_sin_stock(sd_v)
        per_p = periodos_sin_stock(sd_p)
        durs  = [p['dias'] for p in per_v if p['dias'] <= 30]
        base  = durs if durs else [p['dias'] for p in per_v]
        # Mediana de las duraciones de quiebre (más representativa que la moda,
        # que con duraciones únicas devolvía un valor casi al azar).
        trepo = mediana(base) if base else 7
        trepo = min(trepo, 7)  # máximo 7 días para evitar que casos atípicos bloqueen despachos

        # Velocidades
        vel_v, tot_v, dsc_v = velocidad(dv, sd_v)
        vel_p, tot_p, dsc_p = velocidad(dp, sd_p)
        vel_t = vel_v + vel_p

        # Días restantes
        total_hoy = sv_hoy + sp_hoy
        dias_v    = round(sv_hoy/vel_v,1) if vel_v>0 and sv_hoy>0 else (0 if sv_hoy==0 else None)
        dias_p    = round(sp_hoy/vel_p,1) if vel_p>0 and sp_hoy>0 else (0 if sp_hoy==0 else None)
        dias_t    = round(total_hoy/vel_t,1) if vel_t>0 else None
        # días que le queda a Vitacura considerando demanda total (ventas VIT + PAT)
        dias_prod = round(sv_hoy/vel_t,1) if vel_t>0 and sv_hoy>0 else (0 if sv_hoy==0 else None)

        alerta = (sv_hoy==0 and sp_hoy>0 and vel_v>0) or (sp_hoy==0 and sv_hoy>0 and vel_p>0)
        pto_reorden   = math.ceil(vel_t * trepo)              if vel_t > 0 else 0
        lote_sugerido = math.ceil(vel_t * 30)                 if vel_t > 0 else 0
        despacho      = max(0, math.ceil(vel_p * max(trepo,7)) - sp_hoy) if vel_p > 0 else 0

        # Prioridad real: considera el despacho pendiente a Pataguas
        vit_tras_despacho = max(0, sv_hoy - despacho)
        dias_prod_real    = round(vit_tras_despacho / vel_t, 1) if vel_t > 0 else None

        if   sv_hoy == 0:                                                estado = 'sin_stock'
        elif dias_prod_real is not None and dias_prod_real <= 3:         estado = 'critico'
        elif dias_prod_real is not None and dias_prod_real <= 7:         estado = 'bajo'
        else:                                                            estado = 'ok'

        lts, prom_lote, n_lotes = lotes(dv)

        # Ventas por mes combinadas
        vm_v  = ventas_mes(dv)
        vm_p  = ventas_mes(dp)
        # Últimos 4 meses del conjunto combinado
        meses = sorted(set(list(vm_v)+list(vm_p)))[-4:]
        vm = [{'mes':m,'vit':vm_v.get(m,0),'pat':vm_p.get(m,0),'total':vm_v.get(m,0)+vm_p.get(m,0)} for m in meses]

        resultados.append({
            'sku': sku, 'nombre': NOMBRES[sku], 'cocinero': COCINEROS.get(sku,''),
            'vit': sv_hoy, 'pat': sp_hoy, 'total': total_hoy,
            'vel_vit': vel_v, 'vel_pat': vel_p, 'vel_total': round(vel_t,4),
            'dias_vit': dias_v, 'dias_pat': dias_p, 'dias_total': dias_t, 'dias_prod': dias_prod,
            'estado': estado, 'alerta_dist': alerta,
            'tiempo_repo': trepo, 'pto_reorden': pto_reorden,
            'lote_sugerido': lote_sugerido, 'despacho_sug': despacho,
            'total_vit': tot_v, 'total_pat': tot_p,
            'dias_stock_vit': dsc_v, 'dias_stock_pat': dsc_p,
            'n_lotes': n_lotes, 'prom_lote': prom_lote,
            'periodos_sin_stock_vit': per_v,
            'lotes': lts, 'movs': movimientos(dv, dp), 'ventas_mes': vm,
        })
        print(f"  OK {sku} - {NOMBRES[sku][:30]}")

    orden = {'sin_stock':0,'critico':1,'bajo':2,'ok':3}
    resultados.sort(key=lambda x: (orden[x['estado']], x['dias_prod'] if x['dias_prod'] is not None else 9999))
    return resultados

ARCHIVO_HIST_PROY = os.path.join(CARPETA, 'historial_proyecciones.json')

def guardar_historial_proyecciones(datos):
    # Snapshot mensual de lote_sugerido por SKU (un registro por mes, no por dia)
    # para poder comparar como cambia la proyeccion historica con el tiempo.
    try:
        with open(ARCHIVO_HIST_PROY, encoding='utf-8') as f:
            hist = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        hist = {}
    mes_str = FECHA_HOY.strftime('%Y-%m')
    if mes_str not in hist:
        hist[mes_str] = {d['sku']: d['lote_sugerido'] for d in datos}
        with open(ARCHIVO_HIST_PROY, 'w', encoding='utf-8') as f:
            json.dump(hist, f, ensure_ascii=False, indent=2)
    return hist

# ─── ANÁLISIS DE VENTAS ───────────────────────────────────────
FERIADOS_CHILE = {
    '2024-01-01':'Año Nuevo','2024-03-29':'Viernes Santo','2024-03-30':'Sábado Santo',
    '2024-05-01':'Día del Trabajo','2024-05-21':'Glorias Navales',
    '2024-06-20':'Pueblo Mapuche','2024-06-29':'San Pedro y San Pablo',
    '2024-07-16':'Virgen del Carmen','2024-08-15':'Asunción de la Virgen',
    '2024-09-18':'Independencia','2024-09-19':'Glorias del Ejército',
    '2024-10-12':'Encuentro dos Mundos','2024-10-31':'Iglesias Evangélicas',
    '2024-11-01':'Todos los Santos','2024-12-08':'Inmaculada Concepción',
    '2024-12-25':'Navidad',
    '2025-01-01':'Año Nuevo','2025-04-18':'Viernes Santo','2025-04-19':'Sábado Santo',
    '2025-05-01':'Día del Trabajo','2025-05-21':'Glorias Navales',
    '2025-06-20':'Pueblo Mapuche','2025-06-29':'San Pedro y San Pablo',
    '2025-07-16':'Virgen del Carmen','2025-08-15':'Asunción de la Virgen',
    '2025-09-18':'Independencia','2025-09-19':'Glorias del Ejército',
    '2025-10-12':'Encuentro dos Mundos','2025-10-31':'Iglesias Evangélicas',
    '2025-11-01':'Todos los Santos','2025-11-17':'Elecciones Presidenciales',
    '2025-12-08':'Inmaculada Concepción','2025-12-25':'Navidad',
    '2026-01-01':'Año Nuevo','2026-04-03':'Viernes Santo','2026-04-04':'Sábado Santo',
    '2026-05-01':'Día del Trabajo','2026-05-21':'Glorias Navales',
    '2026-06-20':'Pueblo Mapuche','2026-06-29':'San Pedro y San Pablo',
    '2026-07-16':'Virgen del Carmen','2026-08-15':'Asunción de la Virgen',
    '2026-09-18':'Independencia','2026-09-19':'Glorias del Ejército',
    '2026-10-12':'Encuentro dos Mundos','2026-10-31':'Iglesias Evangélicas',
    '2026-11-01':'Todos los Santos','2026-12-08':'Inmaculada Concepción',
    '2026-12-25':'Navidad',
}
VACACIONES_CHILE = [
    ('2024-01-01','2024-03-03','Vacaciones de verano'),
    ('2024-04-29','2024-05-05','Vacaciones de otoño'),
    ('2024-07-08','2024-07-21','Vacaciones de invierno'),
    ('2025-01-01','2025-03-02','Vacaciones de verano'),
    ('2025-04-25','2025-05-04','Vacaciones de otoño'),
    ('2025-07-07','2025-07-20','Vacaciones de invierno'),
    ('2026-01-01','2026-03-03','Vacaciones de verano'),
    ('2026-06-22','2026-07-03','Vacaciones de invierno'),
]
# Fechas comerciales relevantes (no feriados oficiales)
FESTIVIDADES_CHILE = {
    '2024-02-14': 'San Valentín',
    '2024-05-12': 'Día de las Madres',
    '2024-06-16': 'Día del Padre',
    '2024-12-24': 'Nochebuena',
    '2025-02-14': 'San Valentín',
    '2025-05-11': 'Día de las Madres',
    '2025-06-15': 'Día del Padre',
    '2025-12-24': 'Nochebuena',
    '2026-02-14': 'San Valentín',
    '2026-05-10': 'Día de las Madres',
    '2026-06-21': 'Día del Padre',
    '2026-12-24': 'Nochebuena',
}

def calcular_analisis():
    if not os.path.exists(ARCHIVO_JSON):
        return {'meses': [], 'por_mes': {}, 'promedio_mensual': 0,
                'por_dia_historico': [0.0]*7, 'totales_mensuales': {}}
    try:
        with open(ARCHIVO_JSON, encoding='utf-8') as _f:
            _cache = json.load(_f)
        _ultimo_str = _cache.get('ultimo_update', FECHA_HOY.isoformat())
        _ultimo_ts  = pd.Timestamp(_ultimo_str).normalize().tz_localize(None)
        vit = leer_json('VIT')
        pat = leer_json('PAT')
    except Exception as e:
        print(f'  calcular_analisis error: {e}')
        return {'meses': [], 'por_mes': {}, 'promedio_mensual': 0,
                'por_dia_historico': [0.0]*7, 'totales_mensuales': {}}

    def es_venta(df):
        return (df['Salida'] > 0) & (df['Movimiento de salida'] == 'BOLETA')

    ventas = pd.concat([
        vit[es_venta(vit)][['Fecha', 'SKU', 'Salida']],
        pat[es_venta(pat)][['Fecha', 'SKU', 'Salida']],
    ])
    ventas = ventas[ventas['SKU'].isin(NOMBRES.keys())].copy()

    if len(ventas) == 0:
        return {'meses': [], 'por_mes': {}, 'promedio_mensual': 0,
                'por_dia_historico': [0.0]*7, 'totales_mensuales': {}}

    ventas['mes']     = ventas['Fecha'].dt.to_period('M').astype(str)
    ventas['dia_sem'] = ventas['Fecha'].dt.dayofweek          # 0=Lun, 6=Dom
    ventas['sem_mes'] = ((ventas['Fecha'].dt.day - 1) // 7).clip(upper=4)
    ventas['dia_num'] = ventas['Fecha'].dt.day                # 1-31

    # Stock diario VIT para detectar quiebres
    stock_vit = vit[['Fecha', 'SKU', 'Stock']].copy()
    stock_vit['mes'] = stock_vit['Fecha'].dt.to_period('M').astype(str)

    # Totales mensuales por SKU (para sparklines)
    monthly_sku = ventas.groupby(['mes', 'SKU'])['Salida'].sum().to_dict()

    todos_meses = sorted(ventas['mes'].unique())

    # Incluir mes actual solo si ya transcurrió ≥40% del mes
    mes_actual_str  = FECHA_HOY.to_period('M').strftime('%Y-%m')
    dias_mes_actual = pd.Period(mes_actual_str, 'M').days_in_month
    if FECHA_HOY.day < dias_mes_actual * 0.4:
        todos_meses = [m for m in todos_meses if m != mes_actual_str]

    meses_disp = todos_meses[-12:]

    totales_mes_dict = {m: int(ventas[ventas['mes'] == m]['Salida'].sum()) for m in todos_meses}
    vals = list(totales_mes_dict.values())
    promedio_mensual = int(sum(vals) / len(vals)) if vals else 0

    # Patrón histórico por día de semana
    por_dia_h = ventas.groupby('dia_sem')['Salida'].sum().to_dict()
    total_h   = float(sum(por_dia_h.values()))
    por_dia_historico = [round(float(por_dia_h.get(i, 0)) / total_h * 100, 1) if total_h > 0 else 0.0 for i in range(7)]

    vac_ts = [(pd.Timestamp(vi), pd.Timestamp(vf), vn) for vi, vf, vn in VACACIONES_CHILE]

    por_mes_res = {}
    for mes in meses_disp:
        df_m     = ventas[ventas['mes'] == mes]
        total_m  = int(df_m['Salida'].sum())
        if total_m == 0:
            continue

        pd_dict  = df_m.groupby('dia_sem')['Salida'].sum().to_dict()
        por_dia  = [round(float(pd_dict.get(i, 0)) / total_m * 100, 1) for i in range(7)]

        ps_dict  = df_m.groupby('sem_mes')['Salida'].sum().to_dict()
        por_semana = [int(ps_dict.get(i, 0)) for i in range(5)]

        dn_dict  = df_m.groupby('dia_num')['Salida'].sum().to_dict()
        por_dia_num = {int(k): int(v) for k, v in dn_dict.items()}

        periodo     = pd.Period(mes, 'M')
        dias_n      = periodo.days_in_month
        inicio_mes  = pd.Timestamp(mes + '-01')
        feriados_mes, vacaciones_mes, festividades_mes = [], [], []
        for d_n in range(1, dias_n + 1):
            fecha_d   = inicio_mes + pd.Timedelta(days=d_n - 1)
            fecha_str = fecha_d.strftime('%Y-%m-%d')
            if fecha_str in FERIADOS_CHILE:
                feriados_mes.append({'dia': d_n, 'nombre': FERIADOS_CHILE[fecha_str]})
            else:
                for vi_t, vf_t, vn in vac_ts:
                    if vi_t <= fecha_d <= vf_t:
                        vacaciones_mes.append({'dia': d_n, 'nombre': vn})
                        break
            if fecha_str in FESTIVIDADES_CHILE:
                festividades_mes.append({'dia': d_n, 'nombre': FESTIVIDADES_CHILE[fecha_str]})

        # Quiebres: días con stock mínimo = 0 en VIT para este mes
        sv_mes   = stock_vit[stock_vit['mes'] == mes]
        quiebres = {}
        if len(sv_mes) > 0:
            for sku_q, grp in sv_mes.groupby('SKU'):
                dias_q = int((grp.groupby('Fecha')['Stock'].min() == 0).sum())
                if dias_q > 0:
                    quiebres[sku_q] = dias_q

        por_sku = df_m.groupby('SKU')['Salida'].sum().sort_values(ascending=False)
        productos = []
        idx_m = todos_meses.index(mes) if mes in todos_meses else len(todos_meses) - 1
        spark_meses_list = todos_meses[max(0, idx_m - 5):idx_m + 1]

        es_incompleto   = (mes == mes_actual_str)
        dias_n_mes      = pd.Period(mes, 'M').days_in_month
        dias_datos_est  = int(df_m['dia_num'].max()) if es_incompleto and len(df_m) > 0 else dias_n_mes

        for sku_p, tot_sku in por_sku.items():
            if sku_p not in NOMBRES:
                continue
            spark_vals = [int(monthly_sku.get((sm, sku_p), 0)) for sm in spark_meses_list]
            mx = max(spark_vals) if spark_vals else 1
            spark_norm = [round(v / mx * 10) if mx > 0 else 0 for v in spark_vals]
            tend = 'estable'
            MIN_VOL_TEND = 10  # bajo este volumen mensual el % es ruido, no tendencia real
            if len(spark_vals) >= 2 and spark_vals[-2] > 0:
                val_tend = int(round(spark_vals[-1] * dias_n_mes / dias_datos_est)) if es_incompleto and dias_datos_est > 0 else spark_vals[-1]
                if val_tend >= MIN_VOL_TEND or spark_vals[-2] >= MIN_VOL_TEND:
                    cambio = (val_tend - spark_vals[-2]) / spark_vals[-2]
                    if cambio > 0.15:   tend = 'sube'
                    elif cambio < -0.15: tend = 'baja'
            productos.append({
                'sku': sku_p, 'nombre': NOMBRES[sku_p],
                'total': int(tot_sku),
                'pct': round(float(tot_sku) / total_m * 100, 1),
                'dias_quiebre': quiebres.get(sku_p, 0),
                'spark': spark_norm, 'tendencia': tend,
            })

        diff_pct = round((total_m - promedio_mensual) / promedio_mensual * 100, 1) if promedio_mensual > 0 else 0.0
        incompleto = (mes == mes_actual_str)
        if incompleto:
            last_day = int(df_m['dia_num'].max()) if len(df_m) > 0 else int(FECHA_HOY.day)
            dias_datos = last_day
            df_m_cut   = df_m[df_m['dia_num'] <= dias_datos]
            total_m    = int(df_m_cut['Salida'].sum())
        else:
            dias_datos = int(dias_n)
        if incompleto and dias_datos > 0 and total_m > 0 and sum(por_dia_historico) > 0:
            # Proyección ponderada: usa el patrón histórico de cada día de semana
            # para saber qué fracción del mes ya representan los días transcurridos
            w_elapsed = sum(
                por_dia_historico[(inicio_mes + pd.Timedelta(days=d - 1)).dayofweek]
                for d in range(1, dias_datos + 1)
            )
            w_total = sum(
                por_dia_historico[(inicio_mes + pd.Timedelta(days=d - 1)).dayofweek]
                for d in range(1, dias_n + 1)
            )
            proyeccion = int(round(total_m * w_total / w_elapsed)) if w_elapsed > 0 else int(round(total_m / dias_datos * dias_n))
            diff_proy  = round((proyeccion - promedio_mensual) / promedio_mensual * 100, 1) if promedio_mensual > 0 else 0.0
        elif incompleto and dias_datos > 0:
            proyeccion = int(round(total_m / dias_datos * dias_n))
            diff_proy  = round((proyeccion - promedio_mensual) / promedio_mensual * 100, 1) if promedio_mensual > 0 else 0.0
        else:
            proyeccion = None
            diff_proy  = None
        por_mes_res[mes] = {
            'total': total_m, 'diff_pct': diff_pct,
            'incompleto': incompleto, 'dias_datos': dias_datos, 'dias_mes': int(dias_n),
            'proyeccion': proyeccion, 'diff_proy': diff_proy,
            'por_dia': por_dia, 'por_semana': por_semana, 'por_dia_num': por_dia_num,
            'feriados': feriados_mes, 'vacaciones': vacaciones_mes,
            'festividades': festividades_mes,
            'productos': productos,
        }

    # Promedio mensual por SKU — últimos 3 meses completos con ventas
    meses_completos = [m for m in meses_disp if m != mes_actual_str]
    meses_prom = meses_completos[-3:] if len(meses_completos) >= 3 else meses_completos
    promedios_sku = {}
    for sku in NOMBRES.keys():
        vals = [int(monthly_sku.get((m, sku), 0)) for m in meses_prom if monthly_sku.get((m, sku), 0) > 0]
        if vals:
            promedios_sku[sku] = round(sum(vals) / len(vals))

    return {
        'meses': meses_disp,
        'por_mes': por_mes_res,
        'promedio_mensual': promedio_mensual,
        'por_dia_historico': por_dia_historico,
        'totales_mensuales': {m: totales_mes_dict.get(m, 0) for m in meses_disp},
        'promedios_sku': promedios_sku,
    }

# ─── HTML (sin f-string para evitar conflictos con JS) ───────
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
:root{
  --body-bg:#f8fafc;
  --bg-app:var(--body-bg);
  --card-bg:#ffffff;
  --text-color:#1e293b;
  --border-color:#e2e8f0;
  --system-green:#10b981;
  --shadow-sm:0 1px 3px rgba(0,0,0,0.1);
  --shadow-lg:0 1px 3px rgba(0,0,0,0.1);
  --card-border:1px solid var(--border-color);
  --danger-bg:#fef2f2;--danger-text:#b91c1c;--danger-border:#fee2e2;
  --warn-bg:#fff7ed;--warn-text:#c2410c;--warn-border:#ffedd5;
  --ok-bg:#f0fdf4;--ok-text:#166534;--ok-border:#dcfce7;
  --info-bg:#eff6ff;--info-text:#1d4ed8;--info-border:#dbeafe;
  --neutral-bg:#f8fafc;--neutral-text:#475569;--neutral-border:#e2e8f0;
  --zero-color:#cbd5e1;
}
body.dark-mode{
  --body-bg:#0f172a;
  --card-bg:#1e293b;
  --text-color:#f8fafc;
  --border-color:#334155;
  --neutral-bg:#1e293b;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter',-apple-system,BlinkMacSystemFont,sans-serif;background:var(--body-bg);color:var(--text-color);font-size:14px;max-width:1280px;margin:0 auto;transition:background-color 0.2s,color 0.2s}
.num-zero{color:var(--zero-color)!important;opacity:0.7}

/* ── Header ─────────────────────────────────────────────── */
.header{background:transparent;padding:16px 20px 8px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px}
.logo{display:flex;align-items:center;gap:10px}
.logo-icon{width:30px;height:30px;background:#275300;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;color:#fff}
.logo-nombre{font-size:15px;font-weight:800;letter-spacing:-0.03em;color:var(--text-color)}
.logo-sub{font-size:11px;font-weight:400;color:#64748b;letter-spacing:0.01em;margin-left:4px}
.header-right{display:flex;gap:6px;align-items:center}
.btn-theme-toggle:hover{background-color:rgba(16,185,129,0.1);transform:scale(1.05)}
.header-nav-btns{display:none}
.btn{font-size:12px;font-weight:500;padding:7px 13px;border-radius:8px;border:1px solid var(--neutral-border);cursor:pointer;background:var(--card-bg);color:#475569;font-family:inherit;transition:all 0.15s;box-shadow:var(--shadow-sm)}
.btn:hover{background:var(--border-color);border-color:#cbd5e1}
.btn-primary{background:#166534;color:#fff;border-color:#166534}.btn-primary:hover{background:#14532d}
.nav-active{background:#dcfce7!important;color:#166534!important;border-color:#dcfce7!important;font-weight:600!important}
.fecha{font-size:11px;color:#64748b;font-weight:400;white-space:nowrap}

/* ── Navbar isla flotante ──────────────────────────────────── */
.navbar-cocina{display:flex;gap:4px;background:var(--card-bg);border-radius:12px;box-shadow:var(--shadow-sm);padding:6px;margin:4px 20px 14px;overflow-x:auto;-ms-overflow-style:none;scrollbar-width:none;width:fit-content;max-width:calc(100% - 40px)}
.navbar-cocina::-webkit-scrollbar{display:none}
.navtab{font-size:13px;font-weight:500;padding:9px 18px;border-radius:8px;border:none;background:none;color:#64748b;cursor:pointer;font-family:inherit;white-space:nowrap;transition:background-color 0.15s,color 0.15s}
.navtab:hover{background:var(--border-color)}
.navtab.nav-active{background:#f0fdf4;color:#166534;font-weight:600}
@media (max-width:640px){
  .navbar-cocina{margin:4px 12px 12px;max-width:calc(100% - 24px)}
  .navtab{padding:8px 14px;font-size:12px}
}

/* ── Transición de vistas ──────────────────────────────────── */
@keyframes vistaIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
.vista-enter{animation:vistaIn 0.28s ease-out}
@media (prefers-reduced-motion:reduce){.vista-enter{animation:none}}

/* ── Métricas ────────────────────────────────────────────── */
.metricas{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;padding:0 20px 16px}
.metrica{padding:16px 18px;background:var(--card-bg);border-radius:12px;box-shadow:var(--shadow-sm);border:var(--card-border);transition:transform 0.15s,box-shadow 0.15s}
.metrica:hover{transform:translateY(-2px);box-shadow:var(--shadow-lg)}
.metrica-label{font-size:10px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:0.07em;margin-bottom:5px}
.metrica-valor{font-size:30px;font-weight:800;line-height:1;letter-spacing:-0.02em}
.metrica-valor:not(.activo){color:#cbd5e1!important}
.metrica-valor.activo{color:var(--text-color)}
.metrica-sub{font-size:11px;color:#a0a8a0;margin-top:4px}
.val-rojo{color:var(--danger-text)}.val-amarillo{color:var(--warn-text)}.val-verde{color:var(--ok-text)}
@media (max-width:640px){
  .metricas{grid-template-columns:repeat(2,1fr);padding:0 12px 14px}
  .metrica{padding:12px 14px}
  .metrica-valor{font-size:24px}
}

/* ── Toolbar: buscador + chips ───────────────────────────── */
.toolbar{background:transparent;padding:4px 20px 12px;display:flex;flex-direction:column;gap:10px}
.toolbar-label{font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:0.06em}
select,input[type=text],input[type=number]{font-size:13px;font-weight:500;padding:6px 10px;border:1px solid var(--neutral-border);border-radius:8px;background:var(--card-bg);color:var(--text-color);font-family:inherit}
.search-wrap{position:relative}
.search-wrap .search-ico{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#94a3b8;font-size:15px;pointer-events:none}
.search-wrap input{width:100%;height:44px;padding:0 14px 0 40px;font-size:14px;border:1px solid var(--neutral-border);border-radius:12px;background:var(--card-bg);box-shadow:var(--shadow-sm)}
.search-wrap input:focus{outline:none;border-color:#94a3b8}
.search-clear{position:absolute;right:10px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;color:#94a3b8;font-size:16px;width:28px;height:28px;display:flex;align-items:center;justify-content:center;border-radius:50%}
.search-clear:hover{background:var(--border-color);color:#333}
.chips{display:flex;gap:8px;overflow-x:auto;padding-bottom:8px;-ms-overflow-style:none;scrollbar-width:none}
.chips::-webkit-scrollbar{display:none}
.chip{display:inline-flex;align-items:center;gap:6px;padding:8px 16px;border-radius:20px;border:none;background:var(--card-bg);color:#475569;font-size:12px;font-weight:700;font-family:inherit;cursor:pointer;white-space:nowrap;letter-spacing:0.03em;transition:all 0.15s;box-shadow:var(--shadow-sm)}
.chip:active{transform:scale(0.95)}
.chip-active{background:var(--ok-bg);color:var(--ok-text)}
.chip select{border:none;background:transparent;color:inherit;font-weight:700;font-size:12px;padding:0;cursor:pointer}
.chip select:focus{outline:none}

/* ── Leyenda ─────────────────────────────────────────────── */
.container{padding:4px 20px 16px}

/* ── Cards ───────────────────────────────────────────────── */
.card{background:var(--card-bg);border:var(--card-border);border-radius:12px;margin-bottom:20px;overflow:hidden;box-shadow:var(--shadow-sm);transition:box-shadow 0.15s,transform 0.15s;box-sizing:border-box}
.card.sin_stock{border-left:6px solid #ef4444}
.card.critico{border-left:6px solid #f97316}
.card.bajo{border-left:6px solid #facc15}
.card.ok{border-left:6px solid #e2e8f0}
.card:hover{box-shadow:var(--shadow-lg)}
.dato-cero{color:#cbd5e1!important}
.card-row{padding:20px 32px;display:flex;justify-content:flex-start;align-items:center;gap:24px;cursor:pointer;user-select:none}
.nicho{flex:0 0 30%;min-width:0;background:var(--body-bg);border:1px solid var(--border-color);border-radius:8px;padding:10px 12px}
.nicho-label{display:flex;align-items:center;font-size:14px;font-weight:700;line-height:1.2;color:#475569;margin-bottom:4px}
.nicho-dot{width:7px;height:7px;border-radius:50%;display:inline-block;margin-right:5px;flex-shrink:0}
.stock-data{color:var(--text-color)!important;font-weight:800}
.stock-data.dato-cero{color:#cbd5e1!important}
.nicho-valor{font-size:24px;font-weight:800;letter-spacing:-0.02em}
.nicho-unidad{font-size:11px;font-weight:500;color:#94a3b8}
.nicho-dias{font-size:12px;font-weight:600;line-height:1.2;margin-top:2px;color:#64748b}
.card-info{flex:0 0 35%;min-width:0}
.card-nombre-row{display:flex;align-items:center}
.card-nombre{font-size:15px;font-weight:700;letter-spacing:-0.01em;line-height:1.3;text-transform:lowercase!important}
.card-nombre::first-letter{text-transform:uppercase!important}
.estado-dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px;flex-shrink:0}
.card-meta{font-size:12px;color:#64748b;margin-top:3px;font-weight:400}
.card-meta .cap{text-transform:capitalize}
.card-badges{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:8px}
/* ── Indicadores de estado: misma forma y tipografía en todo el sitio ── */
.badge,.badge-cobertura,.badge-despacho{display:inline-block;background:var(--card-bg);font-size:12px;font-weight:600;padding:3px 10px;border-radius:6px;line-height:1.2}
.badge.danger{color:var(--danger-text);text-transform:lowercase}
.badge.warning{color:var(--warn-text);text-transform:lowercase}
.badge.ok{color:var(--ok-text);text-transform:lowercase}
.badge-cobertura{color:#475569}
.badge-cobertura.alerta{color:var(--danger-text)}
.badge-despacho.completo{color:#16a34a}
.badge-despacho.parcial{color:#ea580c}
.badge-despacho.producir{color:#94a3b8}
.badge-dist{background:var(--warn-bg);color:var(--warn-text);font-size:10px}
.chevron{font-size:12px;color:#cbd5e1;margin-left:6px;transition:transform 0.2s}
.chevron.open{transform:rotate(180deg)}
.detalle{display:none;border-top:1px solid #f1f5f9}
.detalle.open{display:block}

/* ── Tabs ────────────────────────────────────────────────── */
.tabs{display:flex;border-bottom:1px solid #f1f5f9;background:transparent}
.tab{font-size:11px;font-weight:600;padding:9px 16px;border:none;background:none;color:#94a3b8;cursor:pointer;font-family:inherit;border-bottom:2px solid transparent;text-transform:uppercase;letter-spacing:0.04em}
.tab.active{color:#166534;border-bottom-color:#166534}
.tab-body{display:none;padding:14px;overflow-x:auto}
.tab-body.active{display:block}

/* ── Movimientos ─────────────────────────────────────────── */
.movs-filtros{display:flex;gap:6px;margin-bottom:8px}
.movs-filter-btn{font-size:11px;font-weight:700;padding:4px 12px;border-radius:20px;border:1px solid var(--neutral-border);background:var(--card-bg);color:#555;cursor:pointer;font-family:inherit}
.movs-filter-btn.activo{background:#166534;color:#fff;border-color:#166534}
.movs-table{width:100%;border-collapse:collapse;font-size:12px}
.movs-table th{text-align:left;color:#94a3b8;padding:14px 12px;border-bottom:1px solid #f1f5f9;font-size:10px;text-transform:uppercase;letter-spacing:0.05em;font-weight:700}
.movs-table td{padding:14px 12px;border-bottom:1px solid #f1f5f9;vertical-align:middle}
.movs-table tr:last-child td{border-bottom:none}
.movs-table tr:hover td{background:var(--neutral-bg)}
.tipo-badge{display:inline-block;font-size:10px;font-weight:700;padding:2px 8px;border-radius:4px;text-transform:uppercase;letter-spacing:0.03em}
.tipo-prod{background:var(--ok-bg);color:var(--ok-text)}
.tipo-venta{background:var(--danger-bg);color:var(--danger-text)}
.tipo-despacho{background:var(--info-bg);color:var(--info-text)}
.tipo-consumo{background:var(--warn-bg);color:var(--warn-text)}
.tipo-desp-rec{background:var(--ok-bg);color:var(--ok-text)}
.tienda-vit{font-size:10px;padding:1px 6px;border-radius:3px;background:var(--ok-bg);color:var(--ok-text);font-weight:600}
.tienda-pat{font-size:10px;padding:1px 6px;border-radius:3px;background:var(--info-bg);color:var(--info-text);font-weight:600}
.cant-pos{color:var(--ok-text);font-weight:700}
.cant-neg{color:var(--danger-text);font-weight:700}

/* ── Insights ────────────────────────────────────────────── */
.insight{background:var(--neutral-bg);border-radius:12px;padding:14px 16px;font-size:12px;color:#475569;line-height:1.8;margin-bottom:10px;border:1px solid var(--neutral-border)}
.insight b{color:var(--text-color);font-weight:600}
.insight-warn{background:var(--warn-bg);border:1px solid var(--warn-border);border-radius:12px;padding:14px 16px;font-size:12px;color:var(--warn-text);line-height:1.8;margin-bottom:10px}
.insight-ok{background:var(--ok-bg);border:1px solid var(--ok-border);border-radius:12px;padding:14px 16px;font-size:12px;color:var(--ok-text);line-height:1.8;margin-bottom:10px}
.insight-peligro{background:var(--danger-bg);border:1px solid var(--danger-border);border-radius:12px;padding:14px 16px;font-size:12px;color:var(--danger-text);line-height:1.8;margin-bottom:10px}
.periodo-chip{display:inline-block;font-size:10px;padding:2px 8px;border-radius:4px;background:var(--danger-bg);color:var(--danger-text);margin:2px}
.lote-card{background:var(--card-bg);border:1px solid var(--neutral-border);border-radius:10px;padding:10px 12px;font-size:11px;margin-bottom:6px;box-shadow:var(--shadow-sm)}
.mes-table{width:100%;border-collapse:collapse;font-size:12px;margin-top:8px}
.mes-table th{text-align:left;color:#94a3b8;padding:14px 12px;border-bottom:1px solid #f1f5f9;font-size:10px;text-transform:uppercase;font-weight:700}
.mes-table td{padding:14px 12px;border-bottom:1px solid #f1f5f9}
.mes-bar{display:inline-block;height:7px;background:#166534;border-radius:3px;margin-left:6px;vertical-align:middle}
.no-res{text-align:center;color:#94a3b8;font-size:13px;padding:40px}

/* ── Días buttons ────────────────────────────────────────── */
.dias-btn{font-size:11px;font-weight:500;padding:5px 10px;border:none;background:none;border-radius:6px;cursor:pointer;color:#475569;font-family:inherit}
.dias-btn:hover{background:var(--border-color)}
.dias-btn-active{background:var(--card-bg);color:#166534;font-weight:700;box-shadow:0 1px 3px rgba(0,0,0,0.12)}

/* ── Guías ───────────────────────────────────────────────── */
.guia-section{background:var(--card-bg);border-radius:12px;box-shadow:var(--shadow-sm);border:var(--card-border);padding:24px;margin:0 20px 24px}
.guia-header{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:8px}
.guia-title{font-size:15px;font-weight:700;margin-bottom:2px;letter-spacing:-0.01em;color:var(--text-color)}
.guia-sub{font-size:11px;color:#64748b;font-weight:400}
.guia-controls{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.dias-group{display:flex;align-items:center;gap:4px;background:var(--neutral-bg);border-radius:8px;padding:4px;border:1px solid var(--neutral-border)}
.table-container-responsive{width:100%;overflow-x:auto!important;display:block!important;-webkit-overflow-scrolling:touch}
.guia-table{width:100%;border-collapse:collapse;font-size:12px;background:transparent}
.guia-table th{text-align:left;color:#94a3b8;padding:14px 12px;border-bottom:1px solid #f1f5f9;font-size:10px;text-transform:uppercase;font-weight:700;letter-spacing:0.05em}
.th-sort{cursor:pointer;user-select:none}
.th-sort:hover{color:#475569}
.sort-arrow{margin-left:4px;color:#cbd5e1;font-size:9px}
.th-sort[data-dir]>.sort-arrow{color:#166534}
.guia-table td{padding:14px 12px;border-bottom:1px solid #f1f5f9;vertical-align:middle}
.guia-table tr:hover td{background:var(--neutral-bg)}
.guia-table td:first-child{font-size:15px;font-weight:600;color:var(--text-color)}
.guia-table td:first-child,.guia-table td:nth-child(3){text-transform:lowercase!important}
.guia-table td:first-child::first-letter,.guia-table td:nth-child(3)::first-letter{text-transform:uppercase!important}
@media (max-width:640px){
  .guia-section{margin:0 12px 18px;padding:16px}
  .guia-table th,.guia-table td{padding:10px 8px}
}
@media (max-width: 768px){
  .table-container-responsive{padding-left:0!important;padding-right:0!important}
  .table-container-responsive table{border-collapse:separate!important;border-spacing:0!important;width:100%!important}
  .table-container-responsive table th,
  .table-container-responsive table td{background-color:var(--card-bg,#ffffff)!important;position:relative;white-space:nowrap!important;vertical-align:middle!important}
  .table-container-responsive table th:first-child,
  .table-container-responsive table td:first-child{
    position:sticky!important;
    left:0!important;
    z-index:12!important;
    background-color:var(--card-bg,#ffffff)!important;
    box-shadow:4px 0 8px -4px rgba(0,0,0,0.15);
    white-space:normal!important;
    width:105px!important;
    min-width:105px!important;
    max-width:105px!important;
    font-size:12px!important;
  }
  .table-container-responsive table td svg,
  .table-container-responsive table td span{
    display:inline-flex;
    align-items:center;
    line-height:1!important;
  }
}

/* ── Resumen ─────────────────────────────────────────────── */
.res-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;padding:0 20px 16px}
.res-card{background:var(--card-bg);border-radius:12px;padding:18px 20px;border:var(--card-border);box-shadow:var(--shadow-sm)}
.res-card-title{font-size:13px;font-weight:700;letter-spacing:-0.01em;color:#475569;margin-bottom:12px}
.res-item{display:flex;align-items:center;justify-content:space-between;padding:8px 0;border-bottom:1px solid #f1f5f9;font-size:12px;font-weight:500}
.res-item:last-child{border-bottom:none}
.res-item-nombre{color:var(--text-color);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:1;margin-right:8px;text-transform:lowercase!important}
.res-item-nombre::first-letter{text-transform:uppercase!important}
.card-master{background:var(--card-bg);border-radius:12px;padding:18px 20px;box-shadow:var(--shadow-sm);border:var(--card-border);margin:0 20px 16px}
.grid-resumen-operaciones{display:grid;grid-template-columns:repeat(2,1fr);gap:20px;margin:24px 20px 16px;align-items:stretch}
.grid-resumen-operaciones .card-master{margin:0;height:400px;display:flex;flex-direction:column;box-sizing:border-box;padding:24px;box-shadow:var(--shadow-sm)}
.grid-resumen-operaciones .card-master .lista-scroll{flex:1;overflow-y:auto}
.card-master-desc{font-size:12px;color:#64748b;margin-bottom:6px;line-height:1.5}
.lista-scroll{max-height:380px;overflow-y:auto}
.grid-etiquetas{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:12px}
.etiqueta-item{display:flex;align-items:center;justify-content:space-between;background:var(--neutral-bg);border:1px solid var(--neutral-border);border-radius:8px;padding:8px 12px;font-size:12px;font-weight:500}
/* ── Grid de KPIs del Resumen ──────────────────────────────── */
.grid-resumen-kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:16px;margin:0 20px 24px}
.card-kpi-individual{background:var(--card-bg);border-radius:12px;padding:16px;box-shadow:var(--shadow-sm);border:var(--card-border);text-align:center;transition:transform 0.15s,box-shadow 0.15s}
.card-kpi-individual:hover{transform:translateY(-2px);box-shadow:var(--shadow-lg)}
.card-kpi-individual .res-stat-val{font-size:28px;font-weight:800;letter-spacing:-0.02em;color:var(--text-color);display:block}
.card-kpi-individual .res-stat-val.dato-cero{color:#cbd5e1}
.card-kpi-individual .res-stat-label{font-size:11px;font-weight:600;color:#94a3b8;letter-spacing:0.02em;margin-top:4px;display:block}
.card-kpi-individual.kpi-rojo .res-stat-val{color:#b91c1c}
.card-kpi-individual.kpi-amarillo .res-stat-val{color:#c2410c}
.card-kpi-individual.kpi-verde .res-stat-val{color:#166534}
@media (max-width:640px){
  .res-grid{grid-template-columns:1fr;padding:0 12px 14px}
  .card-master{margin:0 12px 14px;padding:14px 16px}
  .grid-resumen-operaciones{grid-template-columns:1fr;margin:16px 12px 14px}
  .grid-resumen-kpis{grid-template-columns:repeat(3,1fr);gap:10px;margin:0 12px 18px}
  .card-kpi-individual{padding:12px}
  .card-kpi-individual .res-stat-val{font-size:22px}
}

/* ── Ranking ─────────────────────────────────────────────── */
.rank-table{width:100%;border-collapse:collapse;font-size:13px}
.rank-table th{text-align:left;color:#94a3b8;padding:14px 12px;border-bottom:1px solid #f1f5f9;font-size:10px;text-transform:uppercase;font-weight:700;letter-spacing:0.05em}
.rank-table td{padding:14px 12px;border-bottom:1px solid #f1f5f9;vertical-align:middle}
.rank-table td:first-child{font-size:15px;font-weight:600;color:var(--text-color);text-transform:lowercase!important}
.rank-table td:first-child::first-letter{text-transform:uppercase!important}
.sku-tag{text-transform:none!important}
.rank-table tr:hover td{background:var(--neutral-bg)}
.rank-table td:nth-child(2){text-transform:capitalize}
.rank-num{color:#cbd5e1;font-weight:700;font-size:12px;width:28px}
.rank-bar{display:inline-block;height:6px;background:#166534;border-radius:3px;vertical-align:middle;margin-left:8px}
.rank-zero{color:var(--zero-color)}

/* ── Movimientos scroll ──────────────────────────────────── */
.movs-scroll{max-height:380px;overflow-y:auto;border:1px solid #f1f5f9;border-radius:10px}
.movs-scroll thead th{position:sticky;top:0;background:var(--card-bg);z-index:1}
.btn-vermas{display:block;width:100%;margin-top:8px;padding:9px;font-size:12px;font-weight:700;font-family:inherit;color:var(--ok-text);background:var(--ok-bg);border:1px solid var(--ok-border);border-radius:8px;cursor:pointer}
.btn-vermas:hover{background:#dcfce7}

/* ── Móvil ───────────────────────────────────────────────── */
@media (max-width: 640px){
  .movs-table th:nth-child(3),
  .movs-table td:nth-child(3){display:none}
  .movs-table th,.movs-table td{padding:10px 8px;font-size:11px}
  .tab-body{padding:10px 8px}
  .header{flex-wrap:wrap;gap:8px}
  .card-row{flex-direction:column;align-items:stretch;padding:14px 16px;gap:10px}
  .card-info,.nicho{flex:1 1 auto}
}

/* ── Pestaña Análisis ─────────────────────────────────────── */
.ana-chips{display:flex;flex-wrap:wrap;gap:8px;padding:8px 20px 14px;background:transparent}
.ana-metricas{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;padding:0 20px 16px}
.ana-metricas .metrica{border:none}
.ana-section{padding:24px;border-radius:12px;background:var(--card-bg);box-shadow:var(--shadow-sm);border:var(--card-border);margin:0 20px 24px}
.ana-section-title{font-size:11px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:0.07em;margin-bottom:10px}
.analisis-header{display:flex;align-items:center;gap:10px;margin-bottom:6px;flex-wrap:wrap}
.analisis-titulo{font-size:17px;font-weight:800;letter-spacing:-0.02em}
.badge-mes{font-size:12px;font-weight:600;padding:3px 10px;border-radius:20px;background:var(--danger-bg);color:var(--danger-text)}
.badge-mes.ok{background:var(--ok-bg);color:var(--ok-text)}
.analisis-texto{font-size:13px;line-height:1.65;color:#475569;margin-bottom:10px}
.alerta-banner{margin:0 0 14px;border-radius:12px;padding:14px 16px;border:1px solid;box-shadow:var(--shadow-sm)}
.alerta-roja{background:var(--danger-bg);border-color:var(--danger-border)}
.alerta-amarilla{background:var(--warn-bg);border-color:var(--warn-border)}
.alerta-verde{background:var(--ok-bg);border-color:var(--ok-border)}
.alerta-titulo{font-size:13px;font-weight:700;margin-bottom:5px}
.alerta-roja .alerta-titulo{color:var(--danger-text)}
.alerta-amarilla .alerta-titulo{color:var(--warn-text)}
.alerta-verde .alerta-titulo{color:var(--ok-text)}
.alerta-detalle{font-size:12px;color:#475569;margin-bottom:3px;line-height:1.5}
.alerta-comp{font-size:11px;color:#64748b;margin-bottom:3px}
.alerta-pct{font-size:12px;font-weight:700;color:#475569}
.ana-contexto{padding:7px 0 9px;font-size:11px;color:#64748b;line-height:1.6;border-bottom:1px solid #f1f5f9;margin:16px 0 8px}
.ctx-label{font-weight:700;color:#475569;margin-right:4px}
.factores{display:flex;flex-direction:column;gap:5px}
.factor{display:flex;align-items:flex-start;gap:8px;padding:8px 10px;border-radius:10px;background:var(--neutral-bg);font-size:12px;line-height:1.5;border:1px solid var(--neutral-border)}
.factor-ico{width:22px;height:22px;min-width:22px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;margin-top:1px}
.factor-text{flex:1;color:#475569}
.ico-red{background:var(--danger-bg);color:var(--danger-text)}
.ico-amber{background:var(--warn-bg);color:var(--warn-text)}
.ico-green{background:var(--ok-bg);color:var(--ok-text)}
.ico-gray{background:var(--neutral-bg);color:var(--neutral-text)}
.ana-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin:0 20px 24px}
.ana-card{padding:24px;background:var(--card-bg);border-radius:12px;box-shadow:var(--shadow-sm);border:var(--card-border)}
.ana-card:nth-child(even){border-right:none}
.ana-card-title{font-size:11px;font-weight:700;color:#727969;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:10px}
.bar-row{display:flex;align-items:center;gap:8px;margin-bottom:5px}
.bar-lbl{font-size:11px;color:#727969;width:30px;flex-shrink:0;text-align:right}
.bar-track{flex:1;height:8px;background:#f0f0ee;border-radius:4px;overflow:hidden}
.bar-fill{height:100%;border-radius:4px}
.bar-val{font-size:11px;color:#42493b;width:42px;text-align:right;flex-shrink:0}
.month-wrap{display:flex;align-items:flex-end;gap:6px;margin-bottom:6px}
.month-col{display:flex;flex-direction:column;align-items:center;gap:4px;flex:1;cursor:pointer}
.month-val{font-size:11px;font-weight:600;color:#727969;white-space:nowrap}
.month-track{height:90px;width:100%;display:flex;align-items:flex-end}
.month-bar{width:100%;border-radius:3px 3px 0 0;min-height:4px}
.month-lab{font-size:9px;color:#727969;white-space:nowrap}
.cal-header{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;margin-bottom:2px}
.cal-hdr{font-size:9px;font-weight:700;color:#727969;text-align:center;text-transform:uppercase;padding:1px 0}
.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;margin-bottom:8px;overflow:visible}
.tday{height:52px;display:flex;align-items:center;justify-content:center;border-radius:6px;position:relative;cursor:default;padding:2px}
.tday[data-tip]:not([data-tip=""]):hover::after{content:attr(data-tip);position:absolute;bottom:calc(100% + 6px);left:50%;transform:translateX(-50%);background:rgba(25,25,25,0.92);color:#fff;padding:4px 9px;border-radius:5px;font-size:10px;white-space:nowrap;z-index:300;pointer-events:none;font-weight:500;letter-spacing:0.02em;box-shadow:0 2px 6px rgba(0,0,0,0.25)}
.tday.empty{background:transparent}
.tday-num{position:absolute;top:3px;right:5px;font-size:11px;font-weight:600;line-height:1}
.tday-val{font-size:18px;font-weight:700;line-height:1.1}
.tday-dot{width:5px;height:5px;border-radius:50%;position:absolute;top:3px;left:3px}
.tday-dot.feriado-dot{background:#1960a6}
.tday-dot.festividad-dot{background:#275300}
.tday-dot.vacacion-dot{background:#e6a800}
.leyenda-tl{display:flex;gap:10px;flex-wrap:wrap}
.ana-card-wide{grid-column:1/-1}
.leg-tl-item{display:flex;align-items:center;gap:4px;font-size:10px;color:#727969}
.leg-tl-dot{width:10px;height:10px;border-radius:3px;flex-shrink:0}
.tabla-wrap{padding:10px 14px;overflow-x:auto;background:var(--card-bg)}
@media(max-width:640px){
  .ana-grid{grid-template-columns:1fr;margin:0 12px 14px}
  .ana-metricas{grid-template-columns:repeat(2,1fr);font-size:12px;padding:0 12px 14px}
  .ana-section{margin:0 12px 14px;padding:16px}
  .ana-chips{padding:8px 12px 14px}
}

.ana-section-full{width:100%;margin-top:32px}
"""

JS = """
const MESES_L = {
  '2025-01':'Enero 25','2025-02':'Feb 25','2025-03':'Mar 25','2025-04':'Abr 25',
  '2025-05':'May 25','2025-06':'Jun 25','2025-07':'Jul 25','2025-08':'Ago 25',
  '2025-09':'Sep 25','2025-10':'Oct 25','2025-11':'Nov 25','2025-12':'Dic 25',
  '2026-01':'Enero','2026-02':'Febrero','2026-03':'Marzo','2026-04':'Abril',
  '2026-05':'Mayo','2026-06':'Junio','2026-07':'Julio','2026-08':'Agosto',
  '2026-09':'Septiembre','2026-10':'Octubre','2026-11':'Noviembre','2026-12':'Diciembre'
};

// ── Helpers ────────────────────────────────────────────────
// Convierte "OSTIÓN A LA PARMESANA" en "Ostión a la parmesana": minúsculas
// primero, solo la primera letra del string en mayúscula (no cada palabra).
function tituloCase(str){
  if(!str) return str;
  return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}
// Día principal de la tarjeta = cobertura del stock TOTAL (ambas tiendas)
function diasStr(p){
  if(p.total===0) return 'SIN STOCK';
  if(p.dias_total===null||p.dias_total===undefined) return '+1 año';
  if(p.dias_total>365) return '+1 año';
  return Math.round(p.dias_total)+'d';
}
// Color del badge según la cobertura total (coincide con el número y la barra)
function badgeCls(p){
  if(p.total===0) return 'badge-rojo';
  if(p.dias_total===null||p.dias_total===undefined) return 'badge-verde';
  if(p.dias_total<=3) return 'badge-rojo';
  if(p.dias_total<=14) return 'badge-amarillo';
  return 'badge-verde';
}

// ── Píldora compacta de cobertura (esquina superior derecha) ──
function badgeCobertura(p){
  const dias = (p.dias_total !== null && p.dias_total !== undefined)
    ? Math.round(p.dias_total) : (p.total>0 ? 999 : -1);
  const label = dias < 0 ? 'Sin stock' : dias >= 30 ? '+30 días' : dias+' días';
  const cls   = dias < 0 || dias <= 3 ? 'badge-cobertura alerta' : 'badge-cobertura';
  return '<span class="'+cls+'">'+label+'</span>';
}

// ── Tabla movimientos ──────────────────────────────────────
function tipoBadge(tipo){
  const m={'Producción':'tipo-prod','Venta':'tipo-venta','Despacho':'tipo-despacho',
    'Consumo':'tipo-consumo','Despacho recibido':'tipo-desp-rec','Entrada':'tipo-prod'};
  return '<span class="tipo-badge '+(m[tipo]||'tipo-prod')+'">'+tipo+'</span>';
}
function buildMovs(movs, idx, completo){
  if(!movs||movs.length===0) return '<p style="color:#bbb;font-size:12px;padding:8px 0">Sin movimientos.</p>';
  // Por defecto: últimos 2 meses (mínimo 10 movimientos). El registro completo
  // se conserva siempre; esto solo acota lo que se muestra inicialmente.
  var visibles = movs, ocultos = 0;
  if(!completo){
    var corte = new Date(); corte.setDate(corte.getDate()-60);
    var iso = corte.toISOString().slice(0,10);
    visibles = movs.filter(function(m){return m.fecha_ord >= iso;});
    if(visibles.length < 10) visibles = movs.slice(-10);
    ocultos = movs.length - visibles.length;
  }
  const rows = visibles.map(function(m){
    const tienda = m.tienda==='Vitacura'
      ? '<span class="tienda-vit">Vitacura</span>'
      : '<span class="tienda-pat">Pataguas</span>';
    const cantCls = m.signo==='+' ? 'cant-pos' : 'cant-neg';
    const rowBg = m.stock===0 ? 'background:rgba(186,26,26,0.07);' : '';
    return '<tr data-tienda="'+m.tienda+'" style="'+rowBg+'"><td style="color:#888">'+m.fecha+'</td><td>'+tipoBadge(m.tipo)+'</td>'
      +'<td style="color:#666;font-size:11px">'+m.documento+'</td><td>'+tienda+'</td>'
      +'<td class="'+cantCls+'" style="text-align:right">'+m.signo+m.cantidad+'</td>'
      +'<td style="text-align:right;font-weight:600">'+m.stock+'</td></tr>';
  }).join('');
  var boton = '';
  if(ocultos > 0){
    boton = '<button class="btn-vermas" onclick="verMovsCompletos('+idx+')">Ver historial completo ('+ocultos+' movimientos anteriores)</button>';
  } else if(completo && movs.length > 0){
    boton = '<div style="font-size:11px;color:#aaa;padding:8px 0;text-align:center">Historial completo · '+movs.length+' movimientos</div>';
  }
  var filtros = '<div class="movs-filtros">'
    +'<button class="movs-filter-btn activo" onclick="filtrarMovs('+idx+',&apos;Todas&apos;,this)">Todas</button>'
    +'<button class="movs-filter-btn" onclick="filtrarMovs('+idx+',&apos;Vitacura&apos;,this)">Vitacura</button>'
    +'<button class="movs-filter-btn" onclick="filtrarMovs('+idx+',&apos;Pataguas&apos;,this)">Pataguas</button>'
    +'</div>';
  return filtros+'<div class="movs-scroll"><table class="movs-table"><thead><tr>'
    +'<th>Fecha</th><th>Tipo</th><th>Documento</th><th>Tienda</th>'
    +'<th style="text-align:right">Cant.</th><th style="text-align:right">Stock</th>'
    +'</tr></thead><tbody>'+rows+'</tbody></table></div>'+boton;
}
function verMovsCompletos(i){
  document.getElementById('tab-'+i+'-mov').innerHTML = buildMovs(DATA[i].movs, i, true);
}
function filtrarMovs(idx, tienda, btn){
  document.querySelectorAll('#tab-'+idx+'-mov .movs-filter-btn').forEach(function(b){ b.classList.remove('activo'); });
  btn.classList.add('activo');
  document.querySelectorAll('#tab-'+idx+'-mov tr[data-tienda]').forEach(function(r){
    r.style.display = (tienda==='Todas' || r.dataset.tienda===tienda) ? '' : 'none';
  });
}

// ── Análisis ───────────────────────────────────────────────
function buildAnalisis(p){
  const vt = p.vel_total;
  const repo = p.tiempo_repo||7;
  let h = '';

  // Demanda real
  h += '<div class="insight"><b>📊 Demanda real del período</b><br>'
    + '<span style="color:#777;font-size:11px;line-height:1.5">Se mide con los días que el producto <b>sí tuvo stock</b>. '
    + 'Los días agotados no se cuentan: ahí no se vendió por falta de producto, no de demanda. '
    + 'Así el ritmo refleja lo que de verdad se vende cuando hay disponible.</span><br><br>'
    + 'Vitacura: <b>'+p.total_vit+' und</b> vendidas en '+p.dias_stock_vit+' días con stock → <b>'+p.vel_vit.toFixed(3)+' und/día</b><br>'
    + 'Pataguas: <b>'+p.total_pat+' und</b> vendidas en '+p.dias_stock_pat+' días con stock → <b>'+p.vel_pat.toFixed(3)+' und/día</b><br>'
    + 'Ritmo total: <b>'+vt.toFixed(3)+' und/día</b> → ≈<b>'+(vt*30).toFixed(0)+' und/mes</b> si nunca falta stock</div>';

  // Períodos sin stock Vitacura
  const per = p.periodos_sin_stock_vit||[];
  if(per.length>0){
    const chips = per.map(function(pp){
      return '<span class="periodo-chip">'+pp.inicio+' → '+pp.fin+' ('+pp.dias+'d)</span>';
    }).join('');
    h += '<div class="insight-warn"><b>⚠️ Vitacura sin stock ('+per.length+' veces)</b><br>'
      + chips+'<br><b>Tiempo de reposición estimado: '+repo+' días</b></div>';
  }

  // Stock actual
  const dvt = p.vit===0?'sin stock':(p.dias_vit!==null?Math.round(p.dias_vit)+'d':'—');
  const dpt = p.pat===0?'sin stock':(p.dias_pat!==null?Math.round(p.dias_pat)+'d':'—');
  const uc  = p.estado==='sin_stock'||p.estado==='critico'?'insight-peligro':p.estado==='bajo'?'insight-warn':'insight-ok';
  h += '<div class="'+uc+'"><b>📦 Stock actual</b><br>'
    + 'Vitacura: <b>'+p.vit+' und</b> → '+dvt+' &nbsp;|&nbsp; Pataguas: <b>'+p.pat+' und</b> → '+dpt+'<br>'
    + 'Total: <b>'+p.total+' und</b> → <b>'+diasStr(p)+'</b> de cobertura</div>';

  // Recomendaciones
  if(vt>0){
    h += '<div class="insight-ok"><b>✅ Recomendaciones</b><br>'
      + 'Punto de reorden: cuando queden <b>'+p.pto_reorden+' und</b> → producir de inmediato<br>'
      + 'Consumo estimado 30 días: <b>'+p.lote_sugerido+' und</b> <span style="color:#777;font-size:11px">(a tu ritmo real; ya descuenta los días sin stock)</span><br>'
      + 'Despacho sugerido Pataguas: <b>'+p.despacho_sug+' und</b></div>';
  }

  // Ventas por mes
  if(p.ventas_mes&&p.ventas_mes.length>0){
    const mx = Math.max.apply(null, p.ventas_mes.map(function(m){return m.total;}));
    const filas = p.ventas_mes.map(function(m){
      const lbl = MESES_L[m.mes]||m.mes;
      const bw  = Math.round(m.total/Math.max(mx,1)*80);
      return '<tr><td style="color:#888;width:80px">'+lbl+'</td>'
        +'<td style="text-align:right;width:50px"><b>'+m.vit+'</b></td>'
        +'<td style="text-align:right;width:50px"><b>'+m.pat+'</b></td>'
        +'<td style="text-align:right;width:50px;font-weight:700">'+m.total+'</td>'
        +'<td><div class="mes-bar" style="width:'+bw+'px"></div></td></tr>';
    }).join('');
    h += '<div class="insight"><b>📅 Ventas por mes</b>'
      +'<table class="mes-table"><thead><tr><th>Mes</th><th style="text-align:right">VIT</th>'
      +'<th style="text-align:right">PAT</th><th style="text-align:right">Total</th><th></th>'
      +'</tr></thead><tbody>'+filas+'</tbody></table></div>';
  }

  // Lotes
  if(p.lotes&&p.lotes.length>0){
    const lh = p.lotes.map(function(l){
      return '<div class="lote-card"><b>📦 '+l.fecha+'</b> — '+l.cantidad+' und ('+l.documento+')<br>'
        +'<span style="color:#888">Vendidas luego: '+l.ventas_posteriores+' und en '+l.dias_hasta_siguiente+' días</span></div>';
    }).join('');
    h += '<div class="insight"><b>🏭 Lotes ('+p.n_lotes+', promedio '+p.prom_lote+' und)</b><div style="margin-top:8px">'+lh+'</div></div>';
  }

  return h;
}

// ── Cards ──────────────────────────────────────────────────
function colorDias(dias, und){
  if(und===0) return "#E74C3C";
  if(dias===null||dias===undefined) return "#27AE60";
  if(dias<=3) return "#E74C3C";
  if(dias<=14) return "#E67E22";
  return "#27AE60";
}
function textDias(dias, und){
  if(und===0) return "Sin stock";
  if(dias===null||dias===undefined) return "—";
  return Math.round(dias)+"d restantes";
}
function bloqueSucursal(label, dotColor, valor, dias){
  var cero = valor===0;
  var urg  = colorDias(dias, valor);
  // El color queda solo en el punto de estado; el dato en sí es siempre neutro.
  var alerta = !cero && urg !== '#27AE60';
  var statusDot = alerta ? '<span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:'+urg+';margin-right:5px;vertical-align:middle"></span>' : '';
  return '<div class="nicho">'
    +   '<div class="nicho-label"><span class="nicho-dot" style="background:'+dotColor+'"></span>'+label+'</div>'
    +   '<div class="nicho-valor stock-data'+(cero?' dato-cero':'')+'">'+valor+' <span class="nicho-unidad'+(cero?' dato-cero':'')+'">un</span></div>'
    +   '<div class="nicho-dias'+(cero?' dato-cero':'')+'">'+statusDot+textDias(dias,valor)+'</div>'
    + '</div>';
}
function renderCards(data){
  const cont  = document.getElementById('productos');
  const noRes = document.getElementById('no-res');
  if(!data.length){cont.innerHTML='';noRes.style.display='block';return;}
  noRes.style.display='none';
  const EST_LBL = {sin_stock:'Sin stock', critico:'Crítico', bajo:'Bajo', ok:'OK'};
  cont.innerHTML = data.map(function(p,i){
    const alHtml = p.alerta_dist ? '<span class="badge badge-dist">⚠ Distribución</span>' : '';
    const eCls = (p.estado==='sin_stock'||p.estado==='critico') ? 'danger' : p.estado==='bajo' ? 'warning' : 'ok';

    var estadoDotColor = (p.estado==='sin_stock'||p.estado==='critico') ? 'var(--danger-text)' : p.estado==='bajo' ? 'var(--warn-text)' : 'var(--ok-text)';
    return '<div class="card '+p.estado+'">'
      +'<div class="card-row" onclick="toggleCard('+i+')">'
      +  '<div class="card-info">'
      +    '<div class="card-nombre-row"><span class="estado-dot" style="background:'+estadoDotColor+'"></span><span class="card-nombre">'+tituloCase(p.nombre)+'</span> <span class="chevron" id="chev-'+i+'">▼</span></div>'
      +    '<div class="card-meta">'+p.sku+' · <span class="cap">'+p.cocinero+'</span> · Repo: '+p.tiempo_repo+'d · Reordenar en '+p.pto_reorden+' und</div>'
      +    '<div class="card-badges">'+alHtml
      +      badgeCobertura(p)
      +      '<span class="badge '+eCls+'">'+EST_LBL[p.estado]+'</span>'
      +    '</div>'
      +  '</div>'
      +  bloqueSucursal('Vitacura','#3B6D11',p.vit,p.dias_vit)
      +  bloqueSucursal('Pataguas','#185FA5',p.pat,p.dias_pat)
      +'</div>'
      +'<div class="detalle" id="det-'+i+'">'
      +  '<div class="tabs">'
      +    '<button class="tab active" data-tab="mov" data-idx="'+i+'" onclick="switchTabD(this)">Movimientos</button>'
      +    '<button class="tab" data-tab="analisis" data-idx="'+i+'" onclick="switchTabD(this)">Análisis</button>'
      +  '</div>'
      +  '<div class="tab-body active" id="tab-'+i+'-mov">'+buildMovs(p.movs, i, false)+'</div>'
      +  '<div class="tab-body" id="tab-'+i+'-analisis">'+buildAnalisis(p)+'</div>'
      +'</div>'
      +'</div>';
  }).join('');
}
function toggleCard(i){
  document.getElementById('det-'+i).classList.toggle('open');
  document.getElementById('chev-'+i).classList.toggle('open');
}
function switchTabD(btn){
  var i   = btn.getAttribute('data-idx');
  var tab = btn.getAttribute('data-tab');
  document.querySelectorAll('#det-'+i+' .tab').forEach(function(t){t.classList.remove('active');});
  document.querySelectorAll('#det-'+i+' .tab-body').forEach(function(t){t.classList.remove('active');});
  btn.classList.add('active');
  document.getElementById('tab-'+i+'-'+tab).classList.add('active');
}
function switchTab(i,tab,btn){
  document.querySelectorAll('#det-'+i+' .tab').forEach(function(t){t.classList.remove('active');});
  document.querySelectorAll('#det-'+i+' .tab-body').forEach(function(t){t.classList.remove('active');});
  btn.classList.add('active');
  document.getElementById('tab-'+i+'-'+tab).classList.add('active');
}
var FILTRO_ESTADO = '';
function setChipEstado(btn){
  FILTRO_ESTADO = btn.getAttribute('data-estado');
  document.querySelectorAll('#chips-estado .chip[data-estado]').forEach(function(c){c.classList.remove('chip-active');});
  btn.classList.add('chip-active');
  filtrar();
}
function norm(s){return s.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'');}
function filtrar(){
  const coc = document.getElementById('f-cocinero').value;
  const est = FILTRO_ESTADO;
  const bus = norm(document.getElementById('f-buscar').value);
  const fil = DATA.filter(function(p){
    if(coc && p.cocinero!==coc) return false;
    if(est && p.estado!==est)   return false;
    if(bus && !norm(p.nombre).includes(bus) && !norm(p.sku).includes(bus)) return false;
    return true;
  });
  renderCards(fil);
  updateMetricas(fil);
}
function setMetricaVal(id, val){
  var el = document.getElementById(id);
  el.textContent = val;
  el.classList.toggle('activo', Number(val) > 0);
}
function updateMetricas(data){
  data = data||DATA;
  setMetricaVal('m1', data.filter(function(p){return p.estado==='sin_stock';}).length);
  setMetricaVal('m2', data.filter(function(p){return p.estado==='critico';}).length);
  setMetricaVal('m3', data.filter(function(p){return p.estado==='bajo';}).length);
  setMetricaVal('m4', data.filter(function(p){return p.estado==='ok';}).length);
}

// ── Navegación ─────────────────────────────────────────────
// ── Modo oscuro ──────────────────────────────────────────────
function toggleDarkMode(){
  var activo = document.body.classList.toggle('dark-mode');
  localStorage.setItem('theme', activo ? 'dark' : 'light');
}

var VISTAS = ['vista-resumen','vista-productos','vista-guias','vista-analisis'];
var NAVS   = ['nav-resumen','nav-productos','nav-guias','nav-analisis'];
function switchVista(vistaId, navId, cb){
  VISTAS.forEach(function(v){document.getElementById(v).style.display='none';});
  NAVS.forEach(function(n){document.getElementById(n).classList.remove('nav-active');});
  var vista = document.getElementById(vistaId);
  vista.style.display='block';
  vista.classList.remove('vista-enter');
  void vista.offsetWidth;
  vista.classList.add('vista-enter');
  document.getElementById(navId).classList.add('nav-active');
  window.scrollTo(0,0);
  if(cb) cb();
}
function mostrarResumen(){ switchVista('vista-resumen','nav-resumen', renderResumen); }
function mostrarProductos(){ switchVista('vista-productos','nav-productos'); }
function mostrarGuias(){ switchVista('vista-guias','nav-guias', function(){ renderGuiaProduccion(); renderGuiaDespacho(); }); }
function mostrarRanking(){ switchVista('vista-ranking','nav-ranking', renderRanking); }
function mostrarAnalisis(){
  switchVista('vista-analisis','nav-analisis', function(){
    if(!ANA_MES && ANA_DATA.meses && ANA_DATA.meses.length > 0){
      ANA_MES = ANA_DATA.meses[ANA_DATA.meses.length - 1];
    }
    renderAnalisis();
  });
}

// ── Variables de estado de la pestaña Análisis ─────────────
var ANA_MES  = '';
var ANA_SORT = 'pct';
var ANA_DIR  = -1;
var MNA = ['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];

function mesLabel(m){
  var p = m.split('-');
  return MNA[parseInt(p[1])] + ' ' + p[0];
}
function fechaES(iso){
  if(!iso) return '';
  var p = iso.split('-');
  return parseInt(p[2],10)+' '+MNA[parseInt(p[1],10)]+' '+p[0];
}

function seleccionarMesAna(mes){
  ANA_MES = mes;
  renderAnalisis();
}

function renderAnalisis(){
  renderChipsAna();
  if(!ANA_MES || !ANA_DATA.por_mes || !ANA_DATA.por_mes[ANA_MES]){
    document.getElementById('ana-diagnostico').innerHTML = '<div style="padding:24px 16px;color:#999;font-size:13px">Sin datos para el período seleccionado.</div>';
    document.getElementById('ana-quiebres').innerHTML = '';
    document.getElementById('ana-metricas').innerHTML = '';
    document.getElementById('ana-comparativa').innerHTML = '';
    document.getElementById('ana-calendario').innerHTML = '';
    document.getElementById('ana-tabla').innerHTML = '';
    return;
  }
  var d = ANA_DATA.por_mes[ANA_MES];
  renderDiagnostico(d);
  renderMetricasAna(d);
  renderComparativaMeses();
  renderCalendario(d);
  renderTablaContrib(d);
}

function renderChipsAna(){
  var meses = ANA_DATA.meses || [];
  var html = '';
  for(var i = meses.length - 1; i >= 0; i--){
    var m = meses[i];
    var cls = m === ANA_MES ? 'chip chip-active' : 'chip';
    html += '<button class="'+cls+'" data-mes="'+m+'" onclick="seleccionarMesAna(this.dataset.mes)">'+mesLabel(m)+'</button>';
  }
  document.getElementById('ana-chips').innerHTML = html;
}

function calcularAlerta(d){
  var diasN = d.dias_datos;
  var total = d.total;
  var meses = ANA_DATA.meses || [];
  var comparables = [];
  for(var i=0; i<meses.length; i++){
    var mes = meses[i];
    if(mes === ANA_MES) continue;
    var dm = ANA_DATA.por_mes[mes];
    if(!dm || dm.incompleto) continue;
    // Excluir si tiene festividad en los primeros diasN días (infla la comparación)
    var hasFest = (dm.festividades||[]).some(function(f){ return f.dia <= diasN; });
    if(hasFest) continue;
    // Excluir si más del 50% de los primeros diasN días son vacaciones
    var vacN = (dm.vacaciones||[]).filter(function(v){ return v.dia <= diasN; }).length;
    if(vacN > diasN * 0.5) continue;
    // Ventas acumuladas hasta el día diasN usando por_dia_num (ventas por día del mes)
    var cumN = 0;
    for(var dia=1; dia<=diasN; dia++) cumN += (dm.por_dia_num[dia]||0);
    comparables.push({mes: mes, total: Math.round(cumN)});
  }
  if(comparables.length === 0) return null;
  var prom = comparables.reduce(function(s,c){ return s+c.total; },0) / comparables.length;
  var pct = Math.round((total - prom) / prom * 100);
  return {pct: pct, prom: Math.round(prom), comparables: comparables};
}

function renderDiagnostico(d){
  var total_f = d.total.toLocaleString ? d.total.toLocaleString('es-CL') : d.total;
  var html = '';

  // ── ALERTA TEMPRANA (solo mes en curso) ──────────────────
  if(d.incompleto){
    var alerta = calcularAlerta(d);
    if(alerta){
      var acls = alerta.pct <= -30 ? 'alerta-roja' : alerta.pct <= -10 ? 'alerta-amarilla' : 'alerta-verde';
      var aico = alerta.pct <= -30 ? '⚠️' : alerta.pct <= -10 ? '⚡' : '✓';
      var atit = alerta.pct <= -30 ? 'Mes por debajo del ritmo de crecimiento'
               : alerta.pct <= -10 ? 'Mes ligeramente bajo'
               : 'Mes en buen ritmo';
      var asig = alerta.pct >= 0 ? '+' : '';
      var proy_f = d.proyeccion != null ? '~'+(d.proyeccion.toLocaleString?d.proyeccion.toLocaleString('es-CL'):d.proyeccion)+' un.' : '—';
      var compTxt = alerta.comparables.map(function(c){ return mesLabel(c.mes)+': '+c.total+' un.'; }).join(' · ');
      html += '<div class="alerta-banner '+acls+'">'
        +'<div class="alerta-titulo">'+aico+' '+atit+'</div>'
        +'<div class="alerta-detalle"><strong>'+total_f+' un.</strong> en '+d.dias_datos+' días · proyección al cierre: <strong>'+proy_f+'</strong></div>'
        +'<div class="alerta-comp">Mismo período en meses anteriores → '+compTxt+'</div>'
        +'<div class="alerta-pct"><strong>'+asig+alerta.pct+'%</strong> vs ritmo esperado ('+alerta.prom+' un. promedio en mismos días)</div>'
        +'</div>';
    }
  }

  // ── ENCABEZADO ───────────────────────────────────────────
  var diff = d.diff_pct, diffSign = diff>=0?'+':'';
  var badgeCls = d.incompleto ? 'badge-mes' : (diff>=0?'badge-mes ok':'badge-mes');
  var badgeLabel = d.incompleto ? 'Mes en curso · día '+d.dias_datos : diffSign+diff+'% vs meses anteriores';

  // ── INTRO ────────────────────────────────────────────────
  var intro;
  if(d.incompleto){
    var pf = d.proyeccion != null ? (d.proyeccion.toLocaleString?d.proyeccion.toLocaleString('es-CL'):d.proyeccion) : '—';
    intro = 'Lleva <strong>'+total_f+' unidades</strong> en '+d.dias_datos+' días.'
      +(d.proyeccion ? ' Proyección al cierre (ponderada por patrón de día de semana): <strong>~'+pf+' un.</strong>' : '');
  } else {
    var suf = diff>=0?'sobre':'bajo';
    intro = 'Cerró con <strong>'+total_f+' unidades</strong> — <strong>'+diffSign+diff+'%</strong> '+suf+' el ritmo de los meses anteriores.';
  }

  // ── CONTEXTO DEL MES ────────────────────────────────────
  var ctxPartes = [];
  if(d.vacaciones && d.vacaciones.length>0){
    var vacNoms = [];
    d.vacaciones.forEach(function(v){ if(vacNoms.indexOf(v.nombre)<0) vacNoms.push(v.nombre); });
    var diasVac = d.vacaciones.length;
    var vacTxt = diasVac>=14 ? Math.round(diasVac/7)+' sem.' : diasVac+(diasVac===1?' día':' días');
    ctxPartes.push('📚 '+vacNoms.join(', ')+' ('+vacTxt+')');
  }
  if(d.feriados && d.feriados.length>0)
    ctxPartes.push('🗓 '+d.feriados.map(function(f){return f.nombre;}).join(', '));
  if(d.festividades && d.festividades.length>0)
    ctxPartes.push('⭐ '+d.festividades.map(function(f){return f.nombre;}).join(', '));
  var ctxHtml = ctxPartes.length>0
    ? '<div class="ana-contexto"><span class="ctx-label">Contexto del mes:</span> '+ctxPartes.join(' · ')+'</div>'
    : '';

  // ── QUIEBRES DE STOCK ────────────────────────────────────
  var factHtml = '';
  var conQ = d.productos.filter(function(p){return p.dias_quiebre>2;});
  conQ.sort(function(a,b){return b.pct-a.pct;});
  if(conQ.length>0){
    factHtml += '<div class="factor"><div class="factor-ico ico-red">↓</div><div class="factor-text"><strong>Quiebre de stock — '+conQ[0].nombre+':</strong> sin stock '+conQ[0].dias_quiebre+' días. Representa el '+conQ[0].pct+'% de las ventas.</div></div>';
    if(conQ.length>1)
      factHtml += '<div class="factor"><div class="factor-ico ico-red">↓</div><div class="factor-text"><strong>Quiebres adicionales:</strong> '+conQ.slice(1).map(function(p){return p.nombre;}).join(', ')+'.</div></div>';
  } else {
    factHtml = '<div class="factor"><div class="factor-ico ico-green">✓</div><div class="factor-text">Sin quiebres de stock significativos en este período.</div></div>';
  }

  document.getElementById('ana-diagnostico').innerHTML =
    '<div class="ana-section">'
    +html
    +'<div class="analisis-header">'
    +'<span class="analisis-titulo">'+mesLabel(ANA_MES)+'</span>'
    +'<span class="'+badgeCls+'">'+badgeLabel+'</span>'
    +'</div>'
    +'<p class="analisis-texto">'+intro+'</p>'
    +ctxHtml
    +'</div>';

  document.getElementById('ana-quiebres').innerHTML =
    '<div class="ana-section">'
    +'<div class="ana-section-title">Quiebre de stock</div>'
    +'<div class="factores">'+factHtml+'</div>'
    +'</div>';
}

function renderMetricasAna(d){
  var totalQ = d.productos.reduce(function(s,p){return s+p.dias_quiebre;},0);
  var prodsQ = d.productos.filter(function(p){return p.dias_quiebre>0;}).length;
  var diasEsp = (d.feriados?d.feriados.length:0)+(d.vacaciones?d.vacaciones.length:0);
  var semsAct = d.por_semana.filter(function(v){return v>0;}).length;
  var diff = d.diff_pct;
  var total_f = d.total.toLocaleString ? d.total.toLocaleString('es-CL') : d.total;
  document.getElementById('ana-metricas').innerHTML =
    '<div class="metrica"><div class="metrica-label">Unidades vendidas</div>'
    +'<div class="metrica-valor activo '+(diff>=0?'val-verde':'val-rojo')+'">'+total_f+'</div>'
    +'<div class="metrica-sub">'+(diff>=0?'+':'')+diff+'% vs promedio</div></div>'
    +'<div class="metrica"><div class="metrica-label">Días con quiebre</div>'
    +'<div class="metrica-valor '+(totalQ>0?'activo val-rojo':'val-verde')+'">'+totalQ+'</div>'
    +'<div class="metrica-sub">'+prodsQ+' producto'+(prodsQ!==1?'s':'')+' afectado'+(prodsQ!==1?'s':'')+'</div></div>'
    +'<div class="metrica"><div class="metrica-label">Días especiales</div>'
    +'<div class="metrica-valor'+(diasEsp>0?' activo':'')+'" style="color:#1960a6">'+diasEsp+'</div>'
    +'<div class="metrica-sub">feriados + vacaciones</div></div>'
    +'<div class="metrica"><div class="metrica-label">Semanas activas</div>'
    +'<div class="metrica-valor'+(semsAct>0?' activo':'')+'">'+semsAct+'</div>'
    +'<div class="metrica-sub">con ventas registradas</div></div>';
}

function renderComparativaMeses(){
  var meses   = ANA_DATA.meses || [];
  var totales = ANA_DATA.totales_mensuales || {};
  var vals    = meses.map(function(m){return totales[m]||0;});
  var maxV    = Math.max.apply(null, vals);
  if(maxV === 0){ document.getElementById('ana-comparativa').innerHTML=''; return; }
  var html = '<div class="month-wrap">';
  for(var i=0;i<meses.length;i++){
    var m  = meses[i];
    var v  = totales[m]||0;
    var h  = maxV>0 ? Math.max(4, Math.round(v/maxV*90)) : 4;
    var act = m === ANA_MES;
    var bg  = act ? '#275300' : '#c2c9b7';
    var lbl = MNA[parseInt(m.split('-')[1])];
    html += '<div class="month-col" data-mes="'+m+'" onclick="seleccionarMesAna(this.dataset.mes)">'
      +'<span class="month-val" style="'+(act?'color:#275300':'')+'">'+v+'</span>'
      +'<div class="month-track"><div class="month-bar" style="height:'+h+'px;background:'+bg+'"></div></div>'
      +'<span class="month-lab" style="'+(act?'color:#275300;font-weight:600':'')+'">'+(lbl||m)+'</span></div>';
  }
  html += '</div><div style="font-size:11px;color:#727969">Promedio: <strong style="color:var(--text-color)">'+(ANA_DATA.promedio_mensual.toLocaleString?ANA_DATA.promedio_mensual.toLocaleString('es-CL'):ANA_DATA.promedio_mensual)+' un.</strong> · Clic en barra para ver ese mes</div>';
  document.getElementById('ana-comparativa').innerHTML = html;
}

function heatColor(v, max){
  if(max===0||v===0) return '#e8e8e5';
  var t = v/max;
  if(t<0.2)  return '#d4edda';
  if(t<0.4)  return '#a3d4af';
  if(t<0.6)  return '#5aaa78';
  if(t<0.8)  return '#3b7a55';
  return '#275300';
}
function heatText(v, max){
  if(max===0||v===0) return '#b0b0a8';
  return v/max >= 0.5 ? '#fff' : '#1a3a20';
}
// Contraste del número de fecha (esquina): blanco suave en celdas oscuras,
// gris pizarra nítido en celdas claras o sin ventas.
function heatTextFecha(v, max){
  if(max===0) return '#334155';
  return v/max >= 0.5 ? 'rgba(255,255,255,0.85)' : '#334155';
}

function renderCalendario(d){
  var p     = ANA_MES.split('-');
  var yr    = parseInt(p[0]);
  var mo    = parseInt(p[1]);
  var nDias = new Date(yr, mo, 0).getDate();
  var fSet  = {};
  (d.feriados||[]).forEach(function(f){ fSet[f.dia] = f.nombre; });
  var vSet  = {};
  (d.vacaciones||[]).forEach(function(v){ vSet[v.dia] = v.nombre; });
  var festSet = {};
  (d.festividades||[]).forEach(function(f){ festSet[f.dia] = f.nombre; });

  var pdn  = d.por_dia_num || {};
  var vals = Object.values(pdn).map(Number);
  var maxV = vals.length > 0 ? Math.max.apply(null, vals) : 0;

  var MESES_NOM = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  var moNom = MESES_NOM[mo - 1];

  var primerDia = new Date(yr, mo - 1, 1).getDay();
  var offset    = (primerDia + 6) % 7;

  var HDRS = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'];
  var html = '<div class="cal-header">';
  for(var h=0;h<7;h++) html += '<div class="cal-hdr">'+HDRS[h]+'</div>';
  html += '</div><div class="cal-grid">';

  for(var b=0;b<offset;b++) html += '<div class="tday empty"></div>';

  for(var dd=1;dd<=nDias;dd++){
    var unidades = pdn[dd] || 0;
    var bg   = heatColor(unidades, maxV);
    var tc   = heatText(unidades, maxV);
    var tcF  = heatTextFecha(unidades, maxV);
    var tip  = dd+' '+moNom;
    if(unidades>0) tip += ' · '+unidades+' un.';
    var dots = '';
    if(fSet[dd])    { tip += ' · '+fSet[dd];    dots += '<span class="tday-dot feriado-dot"></span>'; }
    if(festSet[dd]) { tip += ' · '+festSet[dd]; dots += '<span class="tday-dot festividad-dot"></span>'; }
    if(vSet[dd])    { tip += ' · '+vSet[dd];    dots += '<span class="tday-dot vacacion-dot"></span>'; }
    var valHtml = unidades>0 ? '<span class="tday-val" style="color:'+tc+'">'+unidades+'</span>' : '';
    html += '<div class="tday" style="background:'+bg+'" data-tip="'+tip+'">'
      +dots
      +'<span class="tday-num" style="color:'+tcF+'">'+dd+'</span>'
      +valHtml
      +'</div>';
  }
  html += '</div>';

  var hasFest = Object.keys(festSet).length > 0;
  var hasVac  = Object.keys(vSet).length > 0;
  html += '<div class="leyenda-tl" style="margin-top:8px">'
    +'<div class="leg-tl-item"><div class="leg-tl-dot" style="background:#e8e8e5;border:1px solid #c2c9b7"></div>Sin ventas</div>'
    +'<div class="leg-tl-item"><div class="leg-tl-dot" style="background:#a3d4af"></div>Pocas</div>'
    +'<div class="leg-tl-item"><div class="leg-tl-dot" style="background:#275300"></div>Máximo</div>'
    +(fSet&&Object.keys(fSet).length?'<div class="leg-tl-item"><span class="tday-dot feriado-dot" style="position:static;display:inline-block;margin-right:2px"></span>Feriado</div>':'')
    +(hasFest?'<div class="leg-tl-item"><span class="tday-dot festividad-dot" style="position:static;display:inline-block;margin-right:2px"></span>Festividad</div>':'')
    +(hasVac?'<div class="leg-tl-item"><span class="tday-dot vacacion-dot" style="position:static;display:inline-block;margin-right:2px"></span>Vacaciones</div>':'')
    +'</div>';
  document.getElementById('ana-calendario').innerHTML = html;
}

function sortTablaAna(col){
  ANA_DIR = ANA_SORT === col ? -ANA_DIR : -1;
  ANA_SORT = col;
  renderTablaContrib(ANA_DATA.por_mes[ANA_MES]);
}

function renderTablaContrib(d){
  var prods = d.productos.slice();
  var tendVal = {'sube':2,'estable':1,'baja':0};
  prods.sort(function(a,b){
    var va, vb;
    if(ANA_SORT==='nombre') return ANA_DIR*(a.nombre>b.nombre?1:a.nombre<b.nombre?-1:0);
    if(ANA_SORT==='total'){ va=a.total; vb=b.total; }
    else if(ANA_SORT==='quiebre'){ va=a.dias_quiebre; vb=b.dias_quiebre; }
    else if(ANA_SORT==='tendencia'){ va=tendVal[a.tendencia]||0; vb=tendVal[b.tendencia]||0; }
    else { va=a.pct; vb=b.pct; }
    return ANA_DIR*(va-vb);
  });
  var arw = function(col){ return ANA_SORT===col ? (ANA_DIR===-1?' ↓':' ↑') : ' ↕'; };
  var incompleto = d.incompleto;
  var promedios  = ANA_DATA.promedios_sku || {};
  var html = '<table class="rank-table"><thead><tr>'
    +'<th data-col="nombre" onclick="sortTablaAna(this.dataset.col)" style="cursor:pointer;user-select:none">Producto'+arw('nombre')+'</th>'
    +'<th data-col="total" onclick="sortTablaAna(this.dataset.col)" style="cursor:pointer;user-select:none;text-align:right">Unidades'+arw('total')+'</th>'
    +'<th style="text-align:right">Proyec. histórica</th>'
    +'<th data-col="quiebre" onclick="sortTablaAna(this.dataset.col)" style="cursor:pointer;user-select:none;text-align:right">Quiebre'+arw('quiebre')+'</th>'
    +'<th data-col="tendencia" onclick="sortTablaAna(this.dataset.col)" style="cursor:pointer;user-select:none">vs mes ant.'+arw('tendencia')+'</th>'
    +'</tr></thead><tbody>';
  var histProy = ANA_DATA.historial_proy || {};
  var mesesHist = Object.keys(histProy).sort();
  var mesViejo = mesesHist.length >= 2 ? mesesHist[0] : null;
  var mesNuevo = mesesHist.length >= 2 ? mesesHist[mesesHist.length-1] : null;
  for(var i=0;i<prods.length;i++){
    var p = prods[i];
    var spark = '<div style="display:flex;align-items:flex-end;gap:2px;height:18px">';
    for(var j=0;j<p.spark.length;j++){
      var sh = Math.max(2, Math.round((p.spark[j]||0)/10*16));
      var sbg = j===p.spark.length-1 ? '#275300' : '#c2c9b7';
      spark += '<div style="width:5px;height:'+sh+'px;border-radius:1px;background:'+sbg+'"></div>';
    }
    spark += '</div>';
    var qColor = p.dias_quiebre>3?'#ba1a1a':p.dias_quiebre>0?'#E67E22':'#727969';
    var qText  = p.dias_quiebre>0 ? p.dias_quiebre+' d.' : '—';
    var tBadge = p.tendencia==='sube'
      ? '<span style="font-size:10px;padding:2px 8px;background:#eaf3de;color:#275300;border-radius:10px;font-weight:600">Sube ↑</span>'
      : p.tendencia==='baja'
      ? '<span style="font-size:10px;padding:2px 8px;background:#fce8e8;color:#ba1a1a;border-radius:10px;font-weight:600">Baja ↓</span>'
      : '<span style="font-size:10px;padding:2px 8px;background:#f0f0ee;border-radius:10px;color:#727969">Estable</span>';
    var dataP = DATA.find(function(x){return x.sku===p.sku;}) || {};
    var prom  = dataP.lote_sugerido || 0;
    var promTd = prom > 0 ? '<span style="color:#727969">'+prom+' un.</span>' : '<span style="color:#ccc">—</span>';
    if(mesViejo && histProy[mesViejo][p.sku] > 0){
      var vViejo = histProy[mesViejo][p.sku];
      var vNuevo = histProy[mesNuevo][p.sku] || prom;
      var difH = Math.round((vNuevo - vViejo) / vViejo * 100);
      if(Math.abs(difH) >= 10){
        var colH = difH > 0 ? '#275300' : '#ba1a1a';
        promTd += '<br><span style="font-size:9px;color:'+colH+'">'+(difH>0?'+':'')+difH+'% vs '+mesViejo+'</span>';
      }
    }
    html += '<tr>'
      +'<td>'+tituloCase(p.nombre)+'<span class="sku-tag" style="color:#aaa;font-size:10px;margin-left:4px">'+p.sku+'</span></td>'
      +'<td style="text-align:right">'+p.total+'</td>'
      +'<td style="text-align:right">'+promTd+'</td>'
      +'<td style="text-align:right;color:'+qColor+';font-weight:'+(p.dias_quiebre>2?'600':'400')+'">'+qText+'</td>'
      +'<td>'+tBadge+'</td>'
      +'</tr>';
  }
  html += '</tbody></table>';
  document.getElementById('ana-tabla').innerHTML = html;
}

// ── Alertas de cuadratura de stock ─────────────────────────
function nombreSku(sku){
  var p = DATA.find(function(x){return x.sku===sku;});
  return p ? p.nombre : sku;
}
function renderAlertas(){
  var elSal = document.getElementById('res-salidas');
  var elRec = document.getElementById('res-recepciones');
  if(!elSal || !elRec) return;
  if(!ALERTAS || ((ALERTAS.entradas||[]).length===0 && (ALERTAS.salidas||[]).length===0)){
    elSal.innerHTML = '';
    elRec.innerHTML = '<div class="card-master"><div class="res-card-title">Recepciones detectadas</div>'
      +'<div style="color:#94a3b8;font-size:13px;padding:8px 0">Sin movimientos fuera de lo esperado.</div></div>';
    return;
  }
  if((ALERTAS.salidas||[]).length>0){
    var items = ALERTAS.salidas.map(function(a){
      return '<div class="res-item"><span class="res-item-nombre">'+tituloCase(nombreSku(a.sku))
        +' <span style="color:#94a3b8;font-size:11px">('+(a.oficina==='VIT'?'Vitacura':'Pataguas')+')</span></span>'
        +'<span style="color:var(--warn-text);font-weight:700">-'+Math.round(a.cantidad)+' und</span></div>';
    }).join('');
    elSal.innerHTML = '<div class="card-master"><div class="res-card-title">Salidas sin explicación — '+(ALERTAS.fecha||'')+'</div>'
      +'<div class="card-master-desc">El stock bajó sin ventas ni guías que lo expliquen. '
      +'Puede ser consumo interno o merma; si no lo reconoces, revisa la tarjeta de existencia del producto en Bsale.</div>'
      +items+'</div>';
  } else {
    elSal.innerHTML = '';
  }
  if((ALERTAS.entradas||[]).length>0){
    var items2 = ALERTAS.entradas.map(function(a){
      return '<div class="etiqueta-item"><span class="res-item-nombre">'+tituloCase(nombreSku(a.sku))
        +' <span style="color:#94a3b8;font-size:10px">('+(a.oficina==='VIT'?'Vitacura':'Pataguas')+')</span></span>'
        +'<span style="color:var(--ok-text);font-weight:700;flex-shrink:0;margin-left:8px">+'+Math.round(a.cantidad)+' und</span></div>';
    }).join('');
    elRec.innerHTML = '<div class="card-master"><div class="res-card-title">Recepciones detectadas — '+fechaES(ALERTAS.fecha||'')+'</div>'
      +'<div class="card-master-desc">El stock subió: se registraron como producción del día. Normal si hubo recepción.</div>'
      +'<div class="lista-scroll"><div class="grid-etiquetas">'+items2+'</div></div></div>';
  } else {
    elRec.innerHTML = '<div class="card-master"><div class="res-card-title">Recepciones detectadas</div>'
      +'<div style="color:#94a3b8;font-size:13px;padding:8px 0">Sin recepciones hoy.</div></div>';
  }
}

// ── Resumen ejecutivo ──────────────────────────────────────
function renderResumen(){
  var urgentes  = DATA.filter(function(p){return p.estado==='sin_stock'||p.estado==='critico';});
  var bajos     = DATA.filter(function(p){return p.estado==='bajo';});
  var totalVit  = DATA.reduce(function(s,p){return s+p.vit;},0);
  var totalPat  = DATA.reduce(function(s,p){return s+p.pat;},0);
  var dias = 7;
  var despachos = DATA.filter(function(p){
    if(p.vel_pat<=0) return false;
    var nec  = Math.max(0, Math.ceil(p.vel_pat*dias)-p.pat);
    var res  = Math.ceil(p.vel_vit*(p.tiempo_repo||7));
    var disp = Math.max(0, p.vit-res);
    return Math.min(nec,disp)>0;
  });

  renderAlertas();

  // Fila de stats
  var sinStk = DATA.filter(function(p){return p.estado==='sin_stock';}).length;
  var crit   = DATA.filter(function(p){return p.estado==='critico';}).length;
  function statKpi(val, label, alertCls){
    var cardCls = 'card-kpi-individual'+(val>0&&alertCls?' '+alertCls:'');
    var valCls  = 'res-stat-val'+(val>0?'':' dato-cero');
    return '<div class="'+cardCls+'"><span class="'+valCls+'">'+val+'</span><span class="res-stat-label">'+label+'</span></div>';
  }
  var statsHtml = statKpi(sinStk,'Sin stock','kpi-rojo')
    +statKpi(crit,'Críticos','kpi-rojo')
    +statKpi(bajos.length,'Bajo stock','kpi-amarillo')
    +statKpi(despachos.length,'Despachos','')
    +statKpi(totalVit,'Und Vitacura','')
    +statKpi(totalPat,'Und Pataguas','');
  document.getElementById('res-stats').innerHTML = '<div class="grid-resumen-kpis">'+statsHtml+'</div>';

  // Producir urgente (tarjeta maestra, grilla de etiquetas con scroll propio)
  var urgHtml = urgentes.length===0
    ? '<div style="color:var(--ok-text);font-size:13px;padding:8px 0;font-weight:500">Sin urgencias — todo bajo control</div>'
    : '<div class="grid-etiquetas">'+urgentes.map(function(p){
        var dias_s  = p.estado==='sin_stock'?'Sin stock':Math.round(p.dias_prod)+'d';
        var badgeCls = p.estado==='sin_stock'?'danger':'warning';
        return '<div class="etiqueta-item"><span class="res-item-nombre">'+tituloCase(p.nombre)+'</span>'
          +'<span class="badge '+badgeCls+'" style="flex-shrink:0;margin-left:8px">'+dias_s+'</span></div>';
      }).join('')+'</div>';
  document.getElementById('res-alert-col').innerHTML =
    '<div class="card-master"><div class="res-card-title">Producir urgente ('+urgentes.length+')</div>'
   +'<div class="card-master-desc">Productos con stock crítico o agotado — requieren producción inmediata.</div>'
   +'<div class="lista-scroll">'+urgHtml+'</div></div>';

  // Grid de cards
  var despHtml = despachos.length===0
    ? '<div style="color:#94a3b8;font-size:13px;padding:8px 0">Sin despachos pendientes</div>'
    : despachos.map(function(p){
        var nec  = Math.max(0, Math.ceil(p.vel_pat*dias)-p.pat);
        var res  = Math.ceil(p.vel_vit*(p.tiempo_repo||7));
        var disp = Math.max(0, p.vit-res);
        var desp = Math.min(nec,disp);
        return '<div class="res-item"><span class="res-item-nombre">'+tituloCase(p.nombre)+'</span>'
          +'<span style="color:var(--ok-text);font-weight:700">+'+desp+' und</span></div>';
      }).join('');

  var bajosHtml = bajos.length===0
    ? '<div style="color:#94a3b8;font-size:13px;padding:8px 0">Ninguno</div>'
    : bajos.map(function(p){
        return '<div class="res-item"><span class="res-item-nombre">'+tituloCase(p.nombre)+'</span>'
          +'<span style="color:var(--warn-text);font-weight:600;font-size:12px">'+Math.round(p.dias_prod)+'d</span></div>';
      }).join('');

  document.getElementById('res-grid').innerHTML =
    '<div class="res-card"><div class="res-card-title">Despachar a Pataguas ('+despachos.length+')</div>'+despHtml+'</div>'
   +'<div class="res-card"><div class="res-card-title">Stock bajo ('+bajos.length+')</div>'+bajosHtml+'</div>'
   +'<div class="res-card"><div class="res-card-title">Totales en stock</div>'
   +'<div class="res-item"><span class="res-item-nombre">Vitacura</span><span class="stock-data">'+totalVit+' und</span></div>'
   +'<div class="res-item"><span class="res-item-nombre">Pataguas</span><span class="stock-data">'+totalPat+' und</span></div>'
   +'<div class="res-item"><span class="res-item-nombre">Total general</span><span class="stock-data">'+(totalVit+totalPat)+' und</span></div>'
   +'</div>';
}

// ── Ranking ────────────────────────────────────────────────
function buildMesesOpts(){
  var meses = {};
  DATA.forEach(function(p){(p.ventas_mes||[]).forEach(function(m){meses[m.mes]=1;});});
  var keys = Object.keys(meses).sort();
  var opts = '<option value="">Todos los meses</option>';
  keys.forEach(function(k){opts+='<option value="'+k+'">'+(MESES_L[k]||k)+'</option>';});
  document.getElementById('r-mes').innerHTML = opts;
}
function renderRanking(){
  var mes = document.getElementById('r-mes').value;
  var ranked = DATA.map(function(p){
    var vit=0, pat=0;
    (p.ventas_mes||[]).forEach(function(m){
      if(!mes||m.mes===mes){vit+=m.vit; pat+=m.pat;}
    });
    return {nombre:p.nombre, sku:p.sku, vit:vit, pat:pat, total:vit+pat};
  });
  ranked.sort(function(a,b){return b.total-a.total;});
  var mx = ranked[0]?ranked[0].total:1;
  var filas = ranked.map(function(p,i){
    var bw = Math.round(p.total/Math.max(mx,1)*120);
    var z  = p.total===0;
    return '<tr>'
      +'<td class="rank-num rank-'+(z?'zero':'num')+'">'+(i+1)+'</td>'
      +'<td style="font-weight:'+(z?400:500)+';color:'+(z?'#ccc':'#1A1A1A')+'">'+p.nombre+'</td>'
      +'<td style="text-align:right;color:#3B6D11;font-weight:'+(z?400:600)+'">'+p.vit+'</td>'
      +'<td style="text-align:right;color:#185FA5;font-weight:'+(z?400:600)+'">'+p.pat+'</td>'
      +'<td style="text-align:right;font-weight:'+(z?400:700)+';color:'+(z?'#ccc':'#1A1A1A')+'">'+p.total+'</td>'
      +'<td>'+(z?'':'<div class="rank-bar" style="width:'+bw+'px"></div>')+'</td>'
      +'</tr>';
  }).join('');
  document.getElementById('tabla-ranking').innerHTML = filas;
}

// ── Guías ──────────────────────────────────────────────────
var ORDEN_EST = {'sin_stock':0,'critico':1,'bajo':2,'ok':3};

function producir(p, dias){
  if(p.vel_total<=0) return 0;
  var nv = p.vel_vit*dias, np = p.vel_pat*dias;
  var vit_need    = Math.max(0, nv - p.vit);
  var vit_surplus = Math.max(0, p.vit - nv);
  var pat_need    = Math.max(0, np - p.pat);
  var pat_after   = Math.max(0, pat_need - vit_surplus);
  return Math.ceil(vit_need + pat_after);
}
function cobertura(p){
  if(p.vel_total<=0) return 99999;
  var dvit = p.vel_vit>0 ? p.vit/p.vel_vit : 99999;
  var dtot = (p.vit+p.pat)/p.vel_total;
  return Math.min(dvit, dtot);
}

function renderGuiaProduccion(){
  const coc  = document.getElementById('g-cocinero').value;
  const dias = parseInt(document.getElementById('g-dias').value)||7;
  document.getElementById('g-dias-label').textContent = dias+' días';

  var prods = DATA.filter(function(p){return coc ? p.cocinero===coc : true;});
  prods.sort(function(a,b){ return cobertura(a) - cobertura(b); });

  var filas = prods.map(function(p){
    var total   = p.vit + p.pat;
    var und     = producir(p, dias);
    var und_str = und > 0 ? und : '—';
    var cob   = cobertura(p);
    var dt    = (p.vel_total<=0 || cob>=99999) ? '—' : Math.round(cob)+'d';
    var bCls  = cob<=3 ? 'danger' : cob<=7 ? 'warning' : 'ok';
    return '<tr>'
      +'<td>'+tituloCase(p.nombre)+'</td>'
      +'<td style="width:60px"><span class="badge '+bCls+'">'+dt+'</span></td>'
      +'<td style="color:#64748b;font-size:11px">'+tituloCase(p.cocinero)+'</td>'
      +'<td style="text-align:right;color:#64748b">'+p.vit+'</td>'
      +'<td style="text-align:right;color:var(--info-text)">'+p.pat+'</td>'
      +'<td style="text-align:right;font-weight:600">'+total+'</td>'
      +'<td style="text-align:right;font-weight:700;font-size:13px;'+(und>0?'':'color:#cbd5e1')+'">'+und_str+'</td>'
      +'</tr>';
  }).join('');

  document.getElementById('tabla-produccion').innerHTML = filas||'<tr><td colspan="7" style="text-align:center;color:#aaa;padding:20px">Sin productos</td></tr>';
  document.getElementById('resumen-produccion').textContent = prods.length+' productos · para '+dias+' días';
}

// Clasifica el despacho de un producto en uno de 4 estados claros:
//  completo → Vitacura cubre los días pedidos en Pataguas
//  parcial  → manda lo que puede sin dejar Vitacura en cero
//  producir → Pataguas lo necesita pero Vitacura no tiene stock para enviar
//  ok       → Pataguas ya tiene suficiente
function estadoDespacho(p, dias){
  var necesita = Math.max(0, Math.ceil(p.vel_pat*dias) - p.pat);
  var reserva  = Math.ceil(p.vel_vit * (p.tiempo_repo||7));   // lo que Vitacura guarda para sí
  var disp     = Math.max(0, p.vit - reserva);
  var desp     = Math.min(necesita, disp);
  var estado;
  if(necesita===0)        estado='ok';
  else if(desp<=0)        estado='producir';
  else if(desp>=necesita) estado='completo';
  else                    estado='parcial';
  return {estado:estado, necesita:necesita, desp:desp};
}

function renderGuiaDespacho(){
  const dias = parseInt(document.getElementById('d-dias').value)||7;
  document.getElementById('d-dias-label').textContent = dias+' días';

  // Se muestran TODOS los productos que se venden en Pataguas; los que no
  // necesitan nada salen marcados como "OK" para que la guía se explique sola.
  var prods = DATA.filter(function(p){return p.vel_pat>0;});
  var PRIO = {completo:0, parcial:1, producir:2, ok:3};
  prods.sort(function(a,b){
    var ea = estadoDespacho(a,dias).estado, eb = estadoDespacho(b,dias).estado;
    if(PRIO[ea]!==PRIO[eb]) return PRIO[ea]-PRIO[eb];
    var da = a.pat>0&&a.vel_pat>0 ? a.pat/a.vel_pat : 999;
    var db = b.pat>0&&b.vel_pat>0 ? b.pat/b.vel_pat : 999;
    return da-db;
  });

  var despachar = 0, producir = 0;
  var filas = prods.map(function(p){
    var e = estadoDespacho(p, dias);
    if(e.estado==='completo'||e.estado==='parcial') despachar++;
    if(e.estado==='producir') producir++;
    var dpt = p.pat>0&&p.vel_pat>0 ? Math.round(p.pat/p.vel_pat)+'d' : (p.pat===0?'0d':'—');
    var cell;
    if(e.estado==='ok')            cell='<span class="badge-despacho completo">OK</span>';
    else if(e.estado==='completo') cell='<span class="badge-despacho completo">+'+e.desp+'</span>';
    else if(e.estado==='parcial')  cell='<span class="badge-despacho parcial">+'+e.desp+'</span>';
    else                           cell='<span class="badge-despacho producir">'+e.necesita+'</span>';
    return '<tr>'
      +'<td>'+tituloCase(p.nombre)+'</td>'
      +'<td style="text-align:right;color:#64748b">'+p.vit+'</td>'
      +'<td style="text-align:right;color:var(--info-text)">'+p.pat+'</td>'
      +'<td style="text-align:right;color:#64748b">'+dpt+'</td>'
      +'<td style="text-align:right">'+cell+'</td>'
      +'</tr>';
  }).join('');

  document.getElementById('tabla-despacho').innerHTML = filas||'<tr><td colspan="5" style="text-align:center;color:#aaa;padding:20px">Sin productos</td></tr>';
  document.getElementById('resumen-despacho').textContent =
    despachar+' a despachar · '+producir+' a producir · '+dias+' días';
}

// ── Ordenamiento genérico de tablas de Guías (clic en <th>) ────
function valorOrdenCelda(txt){
  txt = txt.trim();
  if(txt==='' || txt==='—' || txt.toUpperCase()==='OK') return null;
  var limpio = txt.replace(/[^0-9.\\-]/g,'');
  if(limpio !== '' && /[0-9]/.test(txt) && !isNaN(parseFloat(limpio))) return parseFloat(limpio);
  return null;
}
function ordenarTablaGuia(th){
  var ths = Array.from(th.parentElement.children);
  var colIndex = ths.indexOf(th);
  var dirNuevo = th.getAttribute('data-dir')==='asc' ? 'desc' : 'asc';
  ths.forEach(function(t){
    t.removeAttribute('data-dir');
    var ar = t.querySelector('.sort-arrow');
    if(ar) ar.textContent = '↕';
  });
  th.setAttribute('data-dir', dirNuevo);
  var arrow = th.querySelector('.sort-arrow');
  if(arrow) arrow.textContent = dirNuevo==='asc' ? '▲' : '▼';

  var tabla = th.closest('table');
  var tbody = tabla.querySelector('tbody');
  var filas = Array.from(tbody.querySelectorAll('tr'));
  if(filas.length===0 || filas[0].querySelector('td[colspan]')) return;

  var mult = dirNuevo==='asc' ? 1 : -1;
  filas.sort(function(a,b){
    var ca = a.children[colIndex], cb = b.children[colIndex];
    if(!ca || !cb) return 0;
    var ta = ca.textContent.trim(), tb = cb.textContent.trim();
    var va = valorOrdenCelda(ta), vb = valorOrdenCelda(tb);
    if(va!==null && vb!==null) return mult*(va-vb);
    if(va!==null && vb===null) return -1;
    if(va===null && vb!==null) return 1;
    return mult*ta.localeCompare(tb,'es',{sensitivity:'base'});
  });
  filas.forEach(function(tr){ tbody.appendChild(tr); });
}

function setDiasProd(d, btn){
  document.getElementById('g-dias').value = d;
  btn.parentElement.querySelectorAll('.dias-btn').forEach(function(b){b.classList.remove('dias-btn-active');});
  btn.classList.add('dias-btn-active');
  renderGuiaProduccion();
}
function setDiasDesp(d, btn){
  document.getElementById('d-dias').value = d;
  btn.parentElement.querySelectorAll('.dias-btn').forEach(function(b){b.classList.remove('dias-btn-active');});
  btn.classList.add('dias-btn-active');
  renderGuiaDespacho();
}

// ── Impresión rollo 80mm ───────────────────────────────────
var ROLLO_CSS = [
  '@import url("https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap");',
  '@page{size:80mm 297mm;margin:0}',
  '*{box-sizing:border-box;margin:0;padding:0}',
  'body{font-family:Inter,Arial,sans-serif;font-size:9pt;width:76mm;padding:4mm 3mm;color:#111}',
  '.logo{font-size:13pt;font-weight:700;letter-spacing:-0.5px;text-align:center;padding-bottom:2mm;border-bottom:2px solid #111;margin-bottom:2mm}',
  '.sub{font-size:7.5pt;color:#111;text-align:center;margin-bottom:1mm}',
  '.fecha{font-size:7pt;color:#111;text-align:center;margin-bottom:3mm}',
  'table{width:100%;border-collapse:collapse}',
  'th{font-size:7pt;font-weight:600;text-transform:uppercase;letter-spacing:0.3px;color:#111;border-bottom:1px solid #ccc;padding:1mm 1mm 1mm 0;text-align:left}',
  'th.r,td.r{text-align:right}',
  'td{font-size:8.5pt;padding:1.5mm 1mm 1.5mm 0;border-bottom:1px solid #eee;vertical-align:middle}',
  'tr:last-child td{border-bottom:none}',
  '.urg td{font-weight:700}',
  '.nom{font-weight:600;line-height:1.2}',
  '.nom.ok{font-weight:400;color:#444}',
  '.dias{font-weight:700;min-width:8mm}',
  '.prod{font-weight:700;font-size:10pt}',
  '.footer{margin-top:3mm;padding-top:2mm;border-top:1px solid #ccc;font-size:7pt;color:#111;text-align:center}',
].join('');

function abrirRollo(bodyHtml){
  var win = window.open('','_blank','width=400,height:750');
  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>La Cocina</title>'
    + '<style>' + ROLLO_CSS + '</style></head><body>' + bodyHtml
    + '<script>window.onload=function(){window.print();}<\\/script>'
    + '</body></html>';
  win.document.write(html);
  win.document.close();
}

function imprimirProduccion(){
  var coc  = document.getElementById('g-cocinero').value;
  var dias = parseInt(document.getElementById('g-dias').value)||7;
  var prods = DATA.filter(function(p){return coc ? p.cocinero===coc : true;});
  prods.sort(function(a,b){ return cobertura(a)-cobertura(b); });

  var filas = prods.map(function(p){
    var cob  = cobertura(p);
    var und  = producir(p, dias);
    var dt   = (p.vel_total<=0||cob>=99999) ? '—' : Math.round(cob)+'d';
    var urg  = cob<=3;
    return '<tr'+(urg?' class="urg"':'')+'>'
      +'<td class="nom'+(cob>7?' ok':'')+'">'+p.nombre+'</td>'
      +'<td class="r dias">'+dt+'</td>'
      +'<td class="r">'+p.vit+'</td>'
      +'<td class="r">'+p.pat+'</td>'
      +'<td class="r prod">'+( und>0 ? und : '—' )+'</td>'
      +'</tr>';
  }).join('');

  var tit = coc ? coc : 'Producción general';
  abrirRollo(
    '<div class="logo">La Cocina</div>'
   +'<div class="sub">'+tit+' · '+dias+' días</div>'
   +'<div class="fecha">FECHA_HOY_PLACEHOLDER</div>'
   +'<table><thead><tr>'
   +'<th>Producto</th><th class="r">Días</th><th class="r">VIT</th><th class="r">PAT</th><th class="r">Prod</th>'
   +'</tr></thead><tbody>'+filas+'</tbody></table>'
   +'<div class="footer">'+prods.length+' productos</div>'
  );
}

function imprimirDespacho(){
  var dias = parseInt(document.getElementById('d-dias').value)||7;
  var prods = DATA.filter(function(p){
    return p.vel_pat>0 && Math.max(0,Math.ceil(p.vel_pat*dias)-p.pat)>0;
  });
  prods.sort(function(a,b){
    var da = a.pat>0&&a.vel_pat>0 ? a.pat/a.vel_pat : 0;
    var db = b.pat>0&&b.vel_pat>0 ? b.pat/b.vel_pat : 0;
    return da-db;
  });

  var filas = prods.map(function(p){
    var desp = Math.max(0, Math.ceil(p.vel_pat*dias)-p.pat);
    return '<tr>'
      +'<td class="nom">'+p.nombre+'</td>'
      +'<td class="r">'+p.vit+'</td>'
      +'<td class="r">'+p.pat+'</td>'
      +'<td class="r prod">+'+desp+'</td>'
      +'</tr>';
  }).join('');

  abrirRollo(
    '<div class="logo">La Cocina</div>'
   +'<div class="sub">Despacho Vitacura → Pataguas · '+dias+' días</div>'
   +'<div class="fecha">FECHA_HOY_PLACEHOLDER</div>'
   +'<table><thead><tr>'
   +'<th>Producto</th><th class="r">VIT</th><th class="r">PAT</th><th class="r">Desp</th>'
   +'</tr></thead><tbody>'+filas+'</tbody></table>'
   +'<div class="footer">'+prods.length+' productos a despachar</div>'
  );
}

// ── Init ───────────────────────────────────────────────────
document.getElementById('f-cocinero').addEventListener('change', filtrar);
var fBuscar = document.getElementById('f-buscar');
var fClear  = document.getElementById('f-buscar-clear');
fBuscar.addEventListener('input', function(){
  fClear.style.display = fBuscar.value ? 'flex' : 'none';
  filtrar();
});
fClear.addEventListener('click', function(){
  fBuscar.value = '';
  fClear.style.display = 'none';
  filtrar();
  fBuscar.focus();
});
document.getElementById('r-mes').addEventListener('change', renderRanking);
renderCards(DATA);
updateMetricas();
buildMesesOpts();
renderResumen();
"""

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>La Cocina · Control de Producción</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>CSS_PLACEHOLDER</style>
</head>
<body>
<script>if(localStorage.getItem('theme')==='dark'){document.body.classList.add('dark-mode');}</script>

<div class="header">
  <div class="logo">
    <div class="logo-icon">🍽</div>
    <div><span class="logo-nombre">La Cocina</span><span class="logo-sub">· Control de Producción</span></div>
  </div>
  <div class="header-right">
    <span class="fecha">Últ. act.: ULTIMO_UPDATE_PLACEHOLDER</span>
    <button class="btn-theme-toggle" aria-label="Cambiar modo de pantalla" onclick="toggleDarkMode()" style="background: none; border: none; cursor: pointer; padding: 8px; display: inline-flex; align-items: center; justify-content: center; border-radius: 50%; transition: background-color 0.2s, transform 0.2s; color: var(--system-green);">
      <svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="12" r="4"></circle>
        <path d="M12 2v2"></path>
        <path d="M12 20v2"></path>
        <path d="m4.93 4.93 1.41 1.41"></path>
        <path d="m17.66 17.66 1.41 1.41"></path>
        <path d="M2 12h2"></path>
        <path d="M20 12h2"></path>
        <path d="m6.34 17.66-1.41 1.41"></path>
        <path d="m19.07 4.93-1.41 1.41"></path>
      </svg>
    </button>
  </div>
</div>

<nav class="navbar-cocina">
  <button class="navtab nav-active" id="nav-resumen" onclick="mostrarResumen()">Resumen</button>
  <button class="navtab" id="nav-productos" onclick="mostrarProductos()">Productos</button>
  <button class="navtab" id="nav-guias" onclick="mostrarGuias()">Guías</button>
  <button class="navtab" id="nav-analisis" onclick="mostrarAnalisis()">Análisis</button>
</nav>

<!-- VISTA RESUMEN -->
<div id="vista-resumen">
  <div id="res-stats"></div>
  <div class="grid-resumen-operaciones">
    <div id="res-alert-col"></div>
    <div id="res-recepciones"></div>
  </div>
  <div id="res-salidas"></div>
  <div class="res-grid" id="res-grid"></div>
</div>

<!-- VISTA PRODUCTOS -->
<div id="vista-productos" style="display:none">

<div class="metricas">
  <div class="metrica rojo"><div class="metrica-label">Sin stock</div><div class="metrica-valor" id="m1">—</div></div>
  <div class="metrica rojo"><div class="metrica-label">Crítico ≤3d</div><div class="metrica-valor" id="m2">—</div></div>
  <div class="metrica amarillo"><div class="metrica-label">Bajo ≤14d</div><div class="metrica-valor" id="m3">—</div></div>
  <div class="metrica verde"><div class="metrica-label">OK &gt;14d</div><div class="metrica-valor activo" id="m4">—</div></div>
</div>
  <div class="toolbar">
    <div class="search-wrap">
      <span class="search-ico">🔍</span>
      <input type="text" id="f-buscar" placeholder="Buscar producto o SKU...">
      <button type="button" id="f-buscar-clear" class="search-clear" aria-label="Borrar búsqueda" style="display:none">✕</button>
    </div>
    <div class="chips" id="chips-estado">
      <button class="chip chip-active" data-estado="" onclick="setChipEstado(this)">Todos</button>
      <button class="chip" data-estado="sin_stock" onclick="setChipEstado(this)">Sin stock</button>
      <button class="chip" data-estado="critico" onclick="setChipEstado(this)">Críticos</button>
      <button class="chip" data-estado="bajo" onclick="setChipEstado(this)">Bajos</button>
      <button class="chip" data-estado="ok" onclick="setChipEstado(this)">OK</button>
      <span class="chip" style="padding:8px 10px">👨‍🍳 <select id="f-cocinero"><option value="">Todos</option><option>CAROLINA</option><option>ADRIANA</option><option>CÉSAR</option><option>JESÚS</option></select></span>
    </div>
  </div>
  <div class="container">
    <div id="productos"></div>
    <div class="no-res" id="no-res" style="display:none">No se encontraron productos.</div>
  </div>
</div>

</div>

<!-- VISTA RANKING -->
<div id="vista-ranking" style="display:none">
  <div class="toolbar">
    <span class="toolbar-label">Mes:</span>
    <select id="r-mes"><option value="">Todos los meses</option></select>
  </div>
  <div style="padding:12px 16px">
    <table class="rank-table">
      <thead><tr>
        <th style="width:28px">#</th>
        <th>Producto</th>
        <th style="text-align:right;color:#3B6D11">Vitacura</th>
        <th style="text-align:right;color:#185FA5">Pataguas</th>
        <th style="text-align:right">Total</th>
        <th style="width:140px"></th>
      </tr></thead>
      <tbody id="tabla-ranking"></tbody>
    </table>
  </div>
</div>

<!-- VISTA GUÍAS -->
<div id="vista-guias" style="display:none">

  <!-- Guía Producción -->
  <div class="guia-section">
    <div class="guia-header">
      <div>
        <div class="guia-title">🏭 Guía de Producción</div>
        <div class="guia-sub" id="resumen-produccion">—</div>
      </div>
      <div class="guia-controls">
        <select id="g-cocinero" onchange="renderGuiaProduccion()">
          <option value="">Todos los cocineros</option>
          <option>CAROLINA</option><option>ADRIANA</option><option>CÉSAR</option><option>JESÚS</option>
        </select>
        <div class="dias-group">
          <button class="dias-btn dias-btn-active" onclick="setDiasProd(7,this)">7d</button>
          <button class="dias-btn" onclick="setDiasProd(15,this)">15d</button>
          <button class="dias-btn" onclick="setDiasProd(30,this)">30d</button>
        </div>
        <input type="number" id="g-dias" value="7" min="1" max="90" onchange="renderGuiaProduccion()" style="width:55px;text-align:center">
        <span class="guia-sub" id="g-dias-label">7 días</span>
        <button class="btn btn-primary" onclick="imprimirProduccion()">🖨 Imprimir</button>
      </div>
    </div>
    <div class="table-container-responsive">
    <table class="guia-table">
      <thead><tr>
        <th class="th-sort" onclick="ordenarTablaGuia(this)">Producto<span class="sort-arrow">↕</span></th><th class="th-sort" style="width:55px" onclick="ordenarTablaGuia(this)">Días<span class="sort-arrow">↕</span></th><th class="th-sort" onclick="ordenarTablaGuia(this)">Cocinero<span class="sort-arrow">↕</span></th>
        <th class="th-sort" style="text-align:right" onclick="ordenarTablaGuia(this)">Vitacura<span class="sort-arrow">↕</span></th><th class="th-sort" style="text-align:right" onclick="ordenarTablaGuia(this)">Pataguas<span class="sort-arrow">↕</span></th>
        <th class="th-sort" style="text-align:right" onclick="ordenarTablaGuia(this)">Total<span class="sort-arrow">↕</span></th><th class="th-sort" style="text-align:right" onclick="ordenarTablaGuia(this)">Producir<span class="sort-arrow">↕</span></th>
      </tr></thead>
      <tbody id="tabla-produccion"></tbody>
    </table>
    </div>
  </div>

  <!-- Guía Despacho -->
  <div class="guia-section">
    <div class="guia-header">
      <div>
        <div class="guia-title">🚚 Guía de Despacho — Vitacura → Pataguas</div>
        <div class="guia-sub" id="resumen-despacho">—</div>
      </div>
      <div class="guia-controls">
        <div class="dias-group">
          <button class="dias-btn dias-btn-active" onclick="setDiasDesp(7,this)">7d</button>
          <button class="dias-btn" onclick="setDiasDesp(15,this)">15d</button>
          <button class="dias-btn" onclick="setDiasDesp(30,this)">30d</button>
        </div>
        <input type="number" id="d-dias" value="7" min="1" max="90" onchange="renderGuiaDespacho()" style="width:55px;text-align:center">
        <span class="guia-sub" id="d-dias-label">7 días</span>
        <button class="btn btn-primary" onclick="imprimirDespacho()">🖨 Imprimir</button>
      </div>
    </div>
    <div style="display:flex;gap:18px;padding:8px 0 12px;font-size:12px;color:#555;flex-wrap:wrap">
      <span><span style="display:inline-block;width:12px;height:12px;background:#16a34a;border-radius:3px;margin-right:5px;vertical-align:middle"></span><b style="color:#16a34a">+N</b> Despacho completo — Vitacura cubre los días pedidos en Pataguas</span>
      <span><span style="display:inline-block;width:12px;height:12px;background:#ea580c;border-radius:3px;margin-right:5px;vertical-align:middle"></span><b style="color:#ea580c">+N</b> Despacho parcial — manda lo posible sin dejar Vitacura en cero</span>
      <span><span style="display:inline-block;width:12px;height:12px;background:#94a3b8;border-radius:3px;margin-right:5px;vertical-align:middle"></span><b style="color:#94a3b8">N</b> Producir — falta en Pataguas pero Vitacura no tiene stock</span>
      <span><b style="color:#16a34a">OK</b> — Pataguas tiene suficiente, no hace falta despachar</span>
    </div>
    <div class="table-container-responsive">
    <table class="guia-table">
      <thead><tr>
        <th class="th-sort" onclick="ordenarTablaGuia(this)">Producto<span class="sort-arrow">↕</span></th><th class="th-sort" style="text-align:right" onclick="ordenarTablaGuia(this)">Stock VIT<span class="sort-arrow">↕</span></th>
        <th class="th-sort" style="text-align:right" onclick="ordenarTablaGuia(this)">Stock PAT<span class="sort-arrow">↕</span></th>
        <th class="th-sort" style="text-align:right" onclick="ordenarTablaGuia(this)">Días PAT<span class="sort-arrow">↕</span></th>
        <th class="th-sort" style="text-align:right" onclick="ordenarTablaGuia(this)">Despachar<span class="sort-arrow">↕</span></th>
      </tr></thead>
      <tbody id="tabla-despacho"></tbody>
    </table>
    </div>
  </div>

</div>

<!-- VISTA ANÁLISIS -->
<div id="vista-analisis" style="display:none">

<div class="ana-chips" id="ana-chips"></div>

<div id="ana-diagnostico"></div>

<div id="ana-quiebres"></div>

<div class="ana-metricas" id="ana-metricas">
  <div class="metrica"><div class="metrica-label">Cargando…</div></div>
</div>

<div class="ana-grid">
  <div class="ana-card ana-card-wide">
    <div class="ana-card-title">Comparativa mensual</div>
    <div id="ana-comparativa"></div>
  </div>
  <div class="ana-card ana-card-wide">
    <div class="ana-card-title">Ventas por día del mes</div>
    <div id="ana-calendario"></div>
  </div>
</div>

<div class="ana-section ana-section-full">
  <div class="ana-section-title">Contribución por producto</div>
  <div class="tabla-wrap table-container-responsive">
    <div id="ana-tabla"></div>
  </div>
</div>

</div>

<script>
const DATA = DATA_PLACEHOLDER;
const ALERTAS = ALERTAS_PLACEHOLDER;
const ANA_DATA = ANA_DATA_PLACEHOLDER;
JS_PLACEHOLDER
</script>
</body>
</html>"""

def generar_html(datos, fecha_str, analisis=None, ultimo_update_str=None):
    data_json    = json.dumps(datos, ensure_ascii=False)
    analisis_json = json.dumps(analisis or {}, ensure_ascii=False)
    # Alertas de cuadratura del día (las escribe actualizar_diario.py)
    archivo_alertas = os.path.join(CARPETA, 'alertas_stock.json')
    if os.path.exists(archivo_alertas):
        alertas_json = json.dumps(json.load(open(archivo_alertas, encoding='utf-8-sig')), ensure_ascii=False)
    else:
        alertas_json = 'null'
    js_final  = JS.replace('FECHA_HOY_PLACEHOLDER', fecha_str)
    html = HTML_TEMPLATE
    html = html.replace('CSS_PLACEHOLDER',            CSS)
    html = html.replace('ANA_DATA_PLACEHOLDER',       analisis_json)  # antes de DATA_PLACEHOLDER
    html = html.replace('DATA_PLACEHOLDER',           data_json)
    html = html.replace('ALERTAS_PLACEHOLDER',        alertas_json)
    html = html.replace('JS_PLACEHOLDER',             js_final)
    html = html.replace('FECHA_HOY_PLACEHOLDER',      fecha_str)
    html = html.replace('ULTIMO_UPDATE_PLACEHOLDER',  ultimo_update_str or fecha_str)
    return html

# ─── MAIN ────────────────────────────────────────────────────
if __name__ == '__main__':
    print('='*50)
    print('La Cocina — Generador de Dashboard')
    print('='*50)
    datos = procesar()
    hist_proy = guardar_historial_proyecciones(datos)
    print('\nCalculando datos de análisis...')
    analisis = calcular_analisis()
    analisis['historial_proy'] = hist_proy
    print(f'  Meses disponibles: {len(analisis.get("meses", []))}')
    # Leer ultimo_update del historial para mostrar en el header
    MESES_ES = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic']
    try:
        with open(ARCHIVO_JSON, encoding='utf-8') as _f:
            _uu = json.load(_f).get('ultimo_update', '')
        _uu_ts = pd.Timestamp(_uu)
        ultimo_update_str = f'{_uu_ts.day} {MESES_ES[_uu_ts.month-1]} · {_uu_ts.strftime("%H:%M")}'
    except Exception:
        ultimo_update_str = FECHA_STR
    print(f'\nGenerando HTML con {len(datos)} productos...')
    html = generar_html(datos, FECHA_STR, analisis, ultimo_update_str)
    with open(ARCHIVO_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Dashboard: {ARCHIVO_HTML}')
    print(f'Tamanio: {len(html):,} chars')

    # Guardar velocidades para el script de email
    vel = {d['sku']: {'vel_vit': d['vel_vit'], 'vel_pat': d['vel_pat'], 'vel_total': d['vel_total']} for d in datos}
    archivo_vel = os.path.join(CARPETA, 'velocidades.json')
    with open(archivo_vel, 'w', encoding='utf-8') as f:
        json.dump(vel, f, ensure_ascii=False)
    print(f'Velocidades: {archivo_vel}')
