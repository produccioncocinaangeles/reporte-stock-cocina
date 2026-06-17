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

        for sku_p, tot_sku in por_sku.items():
            if sku_p not in NOMBRES:
                continue
            spark_vals = [int(monthly_sku.get((sm, sku_p), 0)) for sm in spark_meses_list]
            mx = max(spark_vals) if spark_vals else 1
            spark_norm = [round(v / mx * 10) if mx > 0 else 0 for v in spark_vals]
            tend = 'estable'
            if len(spark_vals) >= 2 and spark_vals[-2] > 0:
                cambio = (spark_vals[-1] - spark_vals[-2]) / spark_vals[-2]
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
        dias_datos = int(FECHA_HOY.day) if incompleto else int(dias_n)
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
            'incompleto': incompleto, 'dias_datos': dias_datos,
            'proyeccion': proyeccion, 'diff_proy': diff_proy,
            'por_dia': por_dia, 'por_semana': por_semana, 'por_dia_num': por_dia_num,
            'feriados': feriados_mes, 'vacaciones': vacaciones_mes,
            'festividades': festividades_mes,
            'productos': productos,
        }

    return {
        'meses': meses_disp,
        'por_mes': por_mes_res,
        'promedio_mensual': promedio_mensual,
        'por_dia_historico': por_dia_historico,
        'totales_mensuales': {m: totales_mes_dict.get(m, 0) for m in meses_disp},
    }

# ─── HTML (sin f-string para evitar conflictos con JS) ───────
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter',-apple-system,BlinkMacSystemFont,sans-serif;background:#f8f9fa;color:#191c1d;font-size:14px;max-width:960px;margin:0 auto}

/* ── Header ─────────────────────────────────────────────── */
.header{background:#fff;border-bottom:1px solid #c2c9b7;padding:12px 20px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;box-shadow:0 1px 4px rgba(0,0,0,0.04)}
.logo{display:flex;align-items:center;gap:10px}
.logo-icon{width:30px;height:30px;background:#275300;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;color:#fff}
.logo-nombre{font-size:15px;font-weight:700;letter-spacing:-0.03em;color:#275300}
.logo-sub{font-size:11px;font-weight:400;color:#727969;letter-spacing:0.01em;margin-left:4px}
.header-right{display:flex;gap:6px;align-items:center}
.header-nav-btns{display:flex;gap:6px}
.btn{font-size:12px;font-weight:500;padding:7px 13px;border-radius:6px;border:1px solid #c2c9b7;cursor:pointer;background:#fff;color:#42493b;font-family:inherit;transition:all 0.15s}
.btn:hover{background:#edeeef;border-color:#727969}
.btn-primary{background:#275300;color:#fff;border-color:#275300}.btn-primary:hover{background:#3b6d11}
.nav-active{background:#b8f389!important;color:#275300!important;border-color:#3b6d11!important;font-weight:600!important}
.fecha{font-size:11px;color:#727969;font-weight:400;white-space:nowrap}

/* ── Métricas ────────────────────────────────────────────── */
.metricas{display:grid;grid-template-columns:repeat(4,1fr);background:#fff;border-bottom:1px solid #c2c9b7}
.metrica{padding:14px 18px;border-right:1px solid #c2c9b7}.metrica:last-child{border-right:none}
.metrica-label{font-size:10px;font-weight:700;color:#727969;text-transform:uppercase;letter-spacing:0.07em;margin-bottom:5px}
.metrica-valor{font-size:30px;font-weight:700;line-height:1;letter-spacing:-0.02em}
.metrica-sub{font-size:11px;color:#a0a8a0;margin-top:4px}
.val-rojo{color:#ba1a1a}.val-amarillo{color:#E67E22}.val-verde{color:#275300}

/* ── Toolbar: buscador + chips ───────────────────────────── */
.toolbar{background:#f8f9fa;padding:12px 16px 4px;display:flex;flex-direction:column;gap:10px}
.toolbar-label{font-size:11px;font-weight:600;color:#727969;text-transform:uppercase;letter-spacing:0.06em}
select,input[type=text],input[type=number]{font-size:13px;font-weight:500;padding:6px 10px;border:1px solid #c2c9b7;border-radius:8px;background:#fff;color:#191c1d;font-family:inherit}
.search-wrap{position:relative}
.search-wrap .search-ico{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#727969;font-size:15px;pointer-events:none}
.search-wrap input{width:100%;height:44px;padding:0 14px 0 40px;font-size:14px;border:1px solid #c2c9b7;border-radius:10px;background:#fff}
.search-wrap input:focus{outline:none;border-color:#275300}
.chips{display:flex;gap:8px;overflow-x:auto;padding-bottom:8px;-ms-overflow-style:none;scrollbar-width:none}
.chips::-webkit-scrollbar{display:none}
.chip{display:inline-flex;align-items:center;gap:6px;padding:8px 16px;border-radius:20px;border:none;background:#e7e8e9;color:#42493b;font-size:12px;font-weight:700;font-family:inherit;cursor:pointer;white-space:nowrap;letter-spacing:0.03em;transition:all 0.15s}
.chip:active{transform:scale(0.95)}
.chip-active{background:#1960a6;color:#fff}
.chip select{border:none;background:transparent;color:inherit;font-weight:700;font-size:12px;padding:0;cursor:pointer}
.chip select:focus{outline:none}

/* ── Leyenda ─────────────────────────────────────────────── */
.leyenda{padding:8px 16px;display:flex;gap:16px;align-items:center;background:#f8f9fa;border-bottom:1px solid #c2c9b7;flex-wrap:wrap}
.leg{display:flex;align-items:center;gap:5px;font-size:11px;color:#727969}
.leg-dot{width:10px;height:10px;border-radius:2px}
.container{padding:12px 16px}

/* ── Cards ───────────────────────────────────────────────── */
.card{background:#fff;border:1px solid #c2c9b7;border-radius:8px;margin-bottom:8px;overflow:hidden;transition:box-shadow 0.15s}
.card:hover{box-shadow:0 4px 16px rgba(0,0,0,0.08)}
.card.sin_stock{border-left:6px solid #ba1a1a}
.card.critico{border-left:6px solid #ba1a1a}
.card.bajo{border-left:6px solid #E67E22}
.card.ok{border-left:6px solid #275300}
.card-top{padding:12px 14px;display:flex;align-items:center;justify-content:space-between;cursor:pointer;user-select:none}
.card-info{flex:1;min-width:0}
.card-nombre{font-size:15px;font-weight:700;letter-spacing:-0.01em;line-height:1.3}
.card-meta{font-size:11px;color:#727969;margin-top:3px;font-weight:400}
.card-badges{display:flex;align-items:center;gap:8px;flex-shrink:0;margin-left:10px}
.badge{font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px}
.badge-estado{font-size:10px;text-transform:uppercase;letter-spacing:0.05em;border-radius:4px;padding:3px 8px}
.badge-rojo{background:#ffdad6;color:#ba1a1a}
.badge-amarillo{background:#FAEEDA;color:#854F0B}
.badge-verde{background:#b8f389;color:#275300}
.badge-dist{background:#FEF3C7;color:#92400E;font-size:10px}
.chevron{font-size:12px;color:#c2c9b7;margin-left:6px;transition:transform 0.2s}
.chevron.open{transform:rotate(180deg)}
.tl-wrap{padding:0 14px 12px;border-top:1px solid #edeeef;padding-top:10px;margin:0 0}
.btn-det{font-size:11px;font-weight:500;background:none;border:none;color:#275300;padding:6px 14px;cursor:pointer;width:100%;text-align:left;border-top:1px solid #edeeef}
.btn-det:hover{background:#f8f9fa}
.detalle{display:none;border-top:1px solid #c2c9b7}
.detalle.open{display:block}

/* ── Tabs ────────────────────────────────────────────────── */
.tabs{display:flex;border-bottom:1px solid #c2c9b7;background:#f8f9fa}
.tab{font-size:11px;font-weight:600;padding:9px 16px;border:none;background:none;color:#727969;cursor:pointer;font-family:inherit;border-bottom:2px solid transparent;text-transform:uppercase;letter-spacing:0.04em}
.tab.active{color:#275300;border-bottom-color:#275300}
.tab-body{display:none;padding:14px;overflow-x:auto}
.tab-body.active{display:block}

/* ── Movimientos ─────────────────────────────────────────── */
.movs-table{width:100%;border-collapse:collapse;font-size:12px}
.movs-table th{text-align:left;color:#727969;padding:6px 10px;border-bottom:1px solid #c2c9b7;font-size:10px;text-transform:uppercase;letter-spacing:0.05em;font-weight:700}
.movs-table td{padding:7px 10px;border-bottom:1px solid #edeeef;vertical-align:middle}
.movs-table tr:last-child td{border-bottom:none}
.movs-table tr:hover td{background:#f8f9fa}
.tipo-badge{display:inline-block;font-size:10px;font-weight:700;padding:2px 8px;border-radius:4px;text-transform:uppercase;letter-spacing:0.03em}
.tipo-prod{background:#b8f389;color:#275300}
.tipo-venta{background:#ffdad6;color:#ba1a1a}
.tipo-despacho{background:#d4e3ff;color:#1960a6}
.tipo-consumo{background:#FEF3C7;color:#92400E}
.tipo-desp-rec{background:#F0FDF4;color:#275300}
.tienda-vit{font-size:10px;padding:1px 6px;border-radius:3px;background:#b8f389;color:#275300;font-weight:600}
.tienda-pat{font-size:10px;padding:1px 6px;border-radius:3px;background:#d4e3ff;color:#1960a6;font-weight:600}
.cant-pos{color:#275300;font-weight:700}
.cant-neg{color:#ba1a1a;font-weight:700}

/* ── Insights ────────────────────────────────────────────── */
.insight{background:#f8f9fa;border-radius:8px;padding:12px 14px;font-size:12px;color:#42493b;line-height:1.8;margin-bottom:10px;border:1px solid #e7e8e9}
.insight b{color:#191c1d;font-weight:600}
.insight-warn{background:#FFFBEB;border:1px solid #FDE68A;border-radius:8px;padding:12px 14px;font-size:12px;color:#92400E;line-height:1.8;margin-bottom:10px}
.insight-ok{background:#F0FDF4;border:1px solid #b8f389;border-radius:8px;padding:12px 14px;font-size:12px;color:#275300;line-height:1.8;margin-bottom:10px}
.insight-peligro{background:#fff5f5;border:1px solid #ffdad6;border-radius:8px;padding:12px 14px;font-size:12px;color:#ba1a1a;line-height:1.8;margin-bottom:10px}
.periodo-chip{display:inline-block;font-size:10px;padding:2px 8px;border-radius:4px;background:#ffdad6;color:#ba1a1a;margin:2px}
.lote-card{background:#fff;border:1px solid #c2c9b7;border-radius:6px;padding:10px 12px;font-size:11px;margin-bottom:6px}
.mes-table{width:100%;border-collapse:collapse;font-size:12px;margin-top:8px}
.mes-table th{text-align:left;color:#727969;padding:5px 10px;border-bottom:1px solid #c2c9b7;font-size:10px;text-transform:uppercase;font-weight:700}
.mes-table td{padding:6px 10px;border-bottom:1px solid #edeeef}
.mes-bar{display:inline-block;height:7px;background:#3b6d11;border-radius:3px;margin-left:6px;vertical-align:middle}
.no-res{text-align:center;color:#727969;font-size:13px;padding:40px}

/* ── Días buttons ────────────────────────────────────────── */
.dias-btn{font-size:11px;font-weight:500;padding:5px 10px;border:none;background:none;border-radius:6px;cursor:pointer;color:#42493b;font-family:inherit}
.dias-btn:hover{background:#e7e8e9}
.dias-btn-active{background:#fff;color:#275300;font-weight:700;box-shadow:0 1px 3px rgba(0,0,0,0.12)}

/* ── Guías ───────────────────────────────────────────────── */
.guia-section{background:#fff;border-bottom:1px solid #c2c9b7;padding:14px 20px}
.guia-header{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px}
.guia-title{font-size:15px;font-weight:700;margin-bottom:2px;letter-spacing:-0.01em;color:#191c1d}
.guia-sub{font-size:11px;color:#727969;font-weight:400}
.guia-controls{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.dias-group{display:flex;align-items:center;gap:4px;background:#f8f9fa;border-radius:8px;padding:4px;border:1px solid #c2c9b7}
.guia-table{width:100%;border-collapse:collapse;font-size:12px}
.guia-table th{text-align:left;color:#727969;padding:8px 10px;border-bottom:2px solid #c2c9b7;font-size:10px;text-transform:uppercase;font-weight:700;letter-spacing:0.05em}
.guia-table td{padding:8px 10px;border-bottom:1px solid #edeeef}
.guia-table tr.urgente{background:#fff5f5}

/* ── Resumen ─────────────────────────────────────────────── */
.res-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;padding:16px}
.res-card{background:#fff;border-radius:8px;padding:16px 18px;border:1px solid #c2c9b7}
.res-card-title{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:#727969;margin-bottom:12px}
.res-item{display:flex;align-items:center;justify-content:space-between;padding:7px 0;border-bottom:1px solid #edeeef;font-size:12px;font-weight:500}
.res-item:last-child{border-bottom:none}
.res-item-nombre{color:#191c1d;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:1;margin-right:8px}
.res-stat{display:flex;flex-direction:column;align-items:center;padding:0 16px;border-right:1px solid #c2c9b7}.res-stat:last-child{border-right:none}
.res-stat-val{font-size:28px;font-weight:700;letter-spacing:-0.02em}
.res-stat-label{font-size:10px;font-weight:600;color:#727969;text-transform:uppercase;letter-spacing:0.06em;margin-top:2px}
.res-stats-row{display:flex;background:#fff;border-bottom:1px solid #c2c9b7;padding:14px 20px;gap:0}

/* ── Ranking ─────────────────────────────────────────────── */
.rank-table{width:100%;border-collapse:collapse;font-size:13px}
.rank-table th{text-align:left;color:#727969;padding:9px 12px;border-bottom:2px solid #c2c9b7;font-size:10px;text-transform:uppercase;font-weight:700;letter-spacing:0.05em}
.rank-table td{padding:9px 12px;border-bottom:1px solid #edeeef;vertical-align:middle}
.rank-table tr:hover td{background:#f8f9fa}
.rank-num{color:#c2c9b7;font-weight:700;font-size:12px;width:28px}
.rank-bar{display:inline-block;height:6px;background:#3b6d11;border-radius:3px;vertical-align:middle;margin-left:8px}
.rank-zero{color:#c2c9b7}

/* ── Movimientos scroll ──────────────────────────────────── */
.movs-scroll{max-height:380px;overflow-y:auto;border:1px solid #edeeef;border-radius:6px}
.movs-scroll thead th{position:sticky;top:0;background:#fff;z-index:1}
.btn-vermas{display:block;width:100%;margin-top:8px;padding:9px;font-size:12px;font-weight:700;font-family:inherit;color:#275300;background:#F0FDF4;border:1px solid #b8f389;border-radius:8px;cursor:pointer}
.btn-vermas:hover{background:#dcfce7}

/* ── Bottom nav ──────────────────────────────────────────── */
.bottom-nav{display:none;position:fixed;bottom:0;left:0;right:0;height:64px;background:#fff;border-top:1px solid #c2c9b7;z-index:100;align-items:center;justify-content:space-around;box-shadow:0 -2px 8px rgba(0,0,0,0.06)}
.bnav-btn{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;padding:6px 12px;border:none;background:none;cursor:pointer;font-family:inherit;font-size:9px;font-weight:700;letter-spacing:0.05em;color:#727969;border-radius:12px;flex:1;text-transform:uppercase;transition:background-color 0.15s,color 0.15s}
.bnav-btn .bnav-icon{font-size:20px;line-height:1;display:block}
.bnav-btn.nav-active{background:#b8f389;color:#275300}

/* ── Móvil ───────────────────────────────────────────────── */
@media (max-width: 640px){
  .bottom-nav{display:flex}
  .header-nav-btns{display:none}
  body{padding-bottom:72px}
  .movs-table th:nth-child(3),
  .movs-table td:nth-child(3){display:none}
  .movs-table th,.movs-table td{padding:6px 6px;font-size:11px}
  .tab-body{padding:10px 8px}
  .header{flex-wrap:wrap;gap:8px}
  .res-grid{grid-template-columns:1fr}
  /* KPIs deslizables estilo Stitch */
  .metricas{display:flex;overflow-x:auto;gap:10px;padding:12px 16px;background:#f8f9fa;border-bottom:none;-ms-overflow-style:none;scrollbar-width:none}
  .metricas::-webkit-scrollbar{display:none}
  .metrica{min-width:118px;background:#fff;border:1px solid #c2c9b7;border-radius:10px;padding:12px 14px;flex-shrink:0}
  .metrica-valor{font-size:26px}
}

/* ── Pestaña Análisis ─────────────────────────────────────── */
.ana-chips{display:flex;flex-wrap:wrap;gap:8px;padding:12px 16px 4px;background:#f8f9fa;border-bottom:1px solid #c2c9b7}
.ana-metricas{display:grid;grid-template-columns:repeat(4,1fr);background:#fff;border-bottom:1px solid #c2c9b7}
.ana-metricas .metrica{border:none;border-right:1px solid #c2c9b7}
.ana-metricas .metrica:last-child{border-right:none}
.ana-section{padding:14px 16px;border-bottom:1px solid #c2c9b7;background:#fff}
.ana-section-title{font-size:11px;font-weight:700;color:#727969;text-transform:uppercase;letter-spacing:0.07em;margin-bottom:10px}
.analisis-header{display:flex;align-items:center;gap:10px;margin-bottom:6px;flex-wrap:wrap}
.analisis-titulo{font-size:17px;font-weight:700;letter-spacing:-0.02em}
.badge-mes{font-size:12px;font-weight:600;padding:3px 10px;border-radius:20px;background:#fce8e8;color:#ba1a1a}
.badge-mes.ok{background:#eaf3de;color:#275300}
.analisis-texto{font-size:13px;line-height:1.65;color:#42493b;margin-bottom:10px}
.alerta-banner{margin:0 0 2px;border-radius:10px;padding:13px 15px;border-left:4px solid}
.alerta-roja{background:#fff5f5;border-left-color:#ba1a1a}
.alerta-amarilla{background:#fffbf0;border-left-color:#E67E22}
.alerta-verde{background:#f0fdf4;border-left-color:#275300}
.alerta-titulo{font-size:13px;font-weight:700;margin-bottom:5px}
.alerta-roja .alerta-titulo{color:#ba1a1a}
.alerta-amarilla .alerta-titulo{color:#b45309}
.alerta-verde .alerta-titulo{color:#275300}
.alerta-detalle{font-size:12px;color:#42493b;margin-bottom:3px;line-height:1.5}
.alerta-comp{font-size:11px;color:#727969;margin-bottom:3px}
.alerta-pct{font-size:12px;font-weight:700;color:#42493b}
.ana-contexto{padding:7px 0 9px;font-size:11px;color:#727969;line-height:1.6;border-bottom:1px solid #e8e8e5;margin-bottom:8px}
.ctx-label{font-weight:700;color:#42493b;margin-right:4px}
.factores{display:flex;flex-direction:column;gap:5px}
.factor{display:flex;align-items:flex-start;gap:8px;padding:8px 10px;border-radius:8px;background:#f8f9fa;font-size:12px;line-height:1.5;border:1px solid #e8e8e6}
.factor-ico{width:22px;height:22px;min-width:22px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;margin-top:1px}
.factor-text{flex:1;color:#42493b}
.ico-red{background:#ba1a1a;color:#fff}
.ico-amber{background:#E67E22;color:#fff}
.ico-green{background:#275300;color:#fff}
.ico-gray{background:#c2c9b7;color:#42493b}
.ana-grid{display:grid;grid-template-columns:1fr 1fr;gap:0;border-bottom:1px solid #c2c9b7}
.ana-card{padding:14px 16px;border-right:1px solid #c2c9b7;border-bottom:1px solid #c2c9b7;background:#fff}
.ana-card:nth-child(even){border-right:none}
.ana-card-title{font-size:11px;font-weight:700;color:#727969;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:10px}
.bar-row{display:flex;align-items:center;gap:8px;margin-bottom:5px}
.bar-lbl{font-size:11px;color:#727969;width:30px;flex-shrink:0;text-align:right}
.bar-track{flex:1;height:8px;background:#f0f0ee;border-radius:4px;overflow:hidden}
.bar-fill{height:100%;border-radius:4px}
.bar-val{font-size:11px;color:#42493b;width:42px;text-align:right;flex-shrink:0}
.month-wrap{display:flex;align-items:flex-end;gap:3px;height:86px;margin-bottom:6px}
.month-col{display:flex;flex-direction:column;align-items:center;gap:2px;flex:1;cursor:pointer}
.month-bar{border-radius:3px 3px 0 0;min-height:4px;width:100%}
.month-lab{font-size:9px;color:#727969;white-space:nowrap}
.cal-header{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;margin-bottom:2px}
.cal-hdr{font-size:9px;font-weight:700;color:#727969;text-align:center;text-transform:uppercase;padding:1px 0}
.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;margin-bottom:8px;overflow:visible}
.tday{aspect-ratio:1;display:flex;align-items:center;justify-content:center;font-size:10px;border-radius:4px;position:relative;cursor:default}
.tday[data-tip]:not([data-tip=""]):hover::after{content:attr(data-tip);position:absolute;bottom:calc(100% + 6px);left:50%;transform:translateX(-50%);background:rgba(25,25,25,0.92);color:#fff;padding:4px 9px;border-radius:5px;font-size:10px;white-space:nowrap;z-index:300;pointer-events:none;font-weight:500;letter-spacing:0.02em;box-shadow:0 2px 6px rgba(0,0,0,0.25)}
.tday.empty{background:transparent}
.tday.ok{background:#e8e8e5;color:#727969}
.tday.feriado{background:#1960a6;color:#fff;font-weight:700}
.tday.festividad{background:#b8f389;color:#275300;font-weight:700}
.tday.vacacion{background:#ffd166;color:#7a5500;font-weight:500}
.leyenda-tl{display:flex;gap:10px;flex-wrap:wrap}
.leg-tl-item{display:flex;align-items:center;gap:4px;font-size:10px;color:#727969}
.leg-tl-dot{width:10px;height:10px;border-radius:3px;flex-shrink:0}
.tabla-wrap{padding:10px 14px;overflow-x:auto;background:#fff}
@media(max-width:640px){
  .ana-grid{grid-template-columns:1fr}
  .ana-card{border-right:none}
  .ana-metricas{grid-template-columns:repeat(2,1fr);font-size:12px}
  .ana-metricas .metrica:nth-child(2){border-right:none}
  .month-lab{font-size:8px}
}
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

// ── Barra de cobertura: gradiente continuo con aguja ───────
// (firma visual del diseño Stitch: rojo → ámbar → verde, marcador del
//  color del estado en la posición de los días de cobertura)
function buildBarra(p){
  const dias = (p.dias_total !== null && p.dias_total !== undefined)
    ? Math.round(p.dias_total) : (p.total>0 ? 999 : -1);
  let pct;
  if(dias <= 0) pct = 0;
  else if(dias <= 3)  pct = (dias/3)*15;
  else if(dias <= 14) pct = 15 + ((dias-3)/11)*40;
  else if(dias <= 30) pct = 55 + ((dias-14)/16)*40;
  else pct = 100;
  pct = Math.min(98, Math.max(2, pct));
  const label = dias < 0 ? 'Sin stock' : dias >= 30 ? '+30 días' : dias+' días';
  const colorMark = dias <= 3 ? '#ba1a1a' : dias <= 14 ? '#E67E22' : '#275300';
  return '<div style="padding:2px 2px 0">'
    + '<div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:5px">'
    +   '<span style="font-size:10px;font-weight:700;color:#727969;text-transform:uppercase;letter-spacing:0.05em">Cobertura estimada</span>'
    +   '<span style="font-size:13px;font-weight:700;color:'+colorMark+'">'+label+'</span>'
    + '</div>'
    + '<div style="position:relative;height:8px;border-radius:4px;background:linear-gradient(90deg,#ba1a1a 0%,#c40413 12%,#E67E22 28%,#f0c24b 45%,#9dd770 70%,#275300 100%);opacity:0.9">'
    +   '<div style="position:absolute;top:-3px;bottom:-3px;left:'+pct+'%;width:4px;transform:translateX(-50%);background:'+colorMark+';border-radius:2px;box-shadow:0 1px 3px rgba(0,0,0,0.35);border:1px solid #fff"></div>'
    + '</div>'
    + '<div style="display:flex;justify-content:space-between;font-size:9px;color:#727969;opacity:0.7;margin-top:3px;text-transform:uppercase;letter-spacing:0.06em">'
    +   '<span>0d</span><span>3d</span><span>14d</span><span>30d+</span>'
    + '</div>'
    + '</div>';
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
    return '<tr><td style="color:#888">'+m.fecha+'</td><td>'+tipoBadge(m.tipo)+'</td>'
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
  return '<div class="movs-scroll"><table class="movs-table"><thead><tr>'
    +'<th>Fecha</th><th>Tipo</th><th>Documento</th><th>Tienda</th>'
    +'<th style="text-align:right">Cant.</th><th style="text-align:right">Stock</th>'
    +'</tr></thead><tbody>'+rows+'</tbody></table></div>'+boton;
}
function verMovsCompletos(i){
  document.getElementById('tab-'+i+'-mov').innerHTML = buildMovs(DATA[i].movs, i, true);
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
function buildBloques(p){
  var cv = colorDias(p.dias_vit, p.vit);
  var cp = colorDias(p.dias_pat, p.pat);
  return '<div style="display:grid;grid-template-columns:1fr 1fr;padding:2px 14px 12px;gap:0">'
    + '<div>'
    +   '<div style="font-size:10px;font-weight:700;color:#3B6D11;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:3px"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#3B6D11;margin-right:5px"></span>Vitacura</div>'
    +   '<div style="font-size:22px;font-weight:700;letter-spacing:-0.02em;color:'+cv+'">'+p.vit+' <span style="font-size:11px;font-weight:500;color:#727969">un</span></div>'
    +   '<div style="font-size:11px;margin-top:1px;color:'+cv+'">'+textDias(p.dias_vit,p.vit)+'</div>'
    + '</div>'
    + '<div style="border-left:1px solid #c2c9b7;padding-left:16px">'
    +   '<div style="font-size:10px;font-weight:700;color:#185FA5;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:3px"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#185FA5;margin-right:5px"></span>Pataguas</div>'
    +   '<div style="font-size:22px;font-weight:700;letter-spacing:-0.02em;color:'+cp+'">'+p.pat+' <span style="font-size:11px;font-weight:500;color:#727969">un</span></div>'
    +   '<div style="font-size:11px;margin-top:1px;color:'+cp+'">'+textDias(p.dias_pat,p.pat)+'</div>'
    + '</div>'
    + '</div>';
}
function renderCards(data){
  const cont  = document.getElementById('productos');
  const noRes = document.getElementById('no-res');
  if(!data.length){cont.innerHTML='';noRes.style.display='block';return;}
  noRes.style.display='none';
  const EST_LBL = {sin_stock:'Sin stock', critico:'Crítico', bajo:'Bajo', ok:'OK'};
  cont.innerHTML = data.map(function(p,i){
    const dStr = diasStr(p);
    const bCls = badgeCls(p);
    const alHtml = p.alerta_dist ? '<span class="badge badge-dist">⚠ Distribución</span>' : '';
    const eCls = (p.estado==='sin_stock'||p.estado==='critico') ? 'badge-rojo' : p.estado==='bajo' ? 'badge-amarillo' : 'badge-verde';

    return '<div class="card '+p.estado+'">'
      +'<div class="card-top" onclick="toggleCard('+i+')">'
      +  '<div class="card-info">'
      +    '<div class="card-nombre">'+p.nombre+'</div>'
      +    '<div class="card-meta">'+p.sku+' · '+p.cocinero+' · Repo: '+p.tiempo_repo+'d · Reordenar en '+p.pto_reorden+' und</div>'
      +  '</div>'
      +  '<div class="card-badges">'+alHtml
      +    '<span class="badge badge-estado '+eCls+'">'+EST_LBL[p.estado]+'</span>'
      +    '<span class="chevron" id="chev-'+i+'">▼</span></div>'
      +'</div>'
      +buildBloques(p)
      +'<div class="tl-wrap">'+buildBarra(p)+'</div>'
      +'<button class="btn-det" onclick="toggleCard('+i+')">▼ Ver movimientos y análisis</button>'
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
function filtrar(){
  const coc = document.getElementById('f-cocinero').value;
  const est = FILTRO_ESTADO;
  const bus = document.getElementById('f-buscar').value.toLowerCase();
  const fil = DATA.filter(function(p){
    if(coc && p.cocinero!==coc) return false;
    if(est && p.estado!==est)   return false;
    if(bus && !p.nombre.toLowerCase().includes(bus) && !p.sku.toLowerCase().includes(bus)) return false;
    return true;
  });
  renderCards(fil);
  updateMetricas(fil);
}
function updateMetricas(data){
  data = data||DATA;
  document.getElementById('m1').textContent = data.filter(function(p){return p.estado==='sin_stock';}).length;
  document.getElementById('m2').textContent = data.filter(function(p){return p.estado==='critico';}).length;
  document.getElementById('m3').textContent = data.filter(function(p){return p.estado==='bajo';}).length;
  document.getElementById('m4').textContent = data.filter(function(p){return p.estado==='ok';}).length;
}

// ── Navegación ─────────────────────────────────────────────
var VISTAS = ['vista-resumen','vista-productos','vista-guias','vista-ranking','vista-analisis'];
var NAVS   = ['nav-resumen','nav-productos','nav-guias','nav-ranking','nav-analisis'];
var BNAVS  = ['bnav-resumen','bnav-productos','bnav-guias','bnav-ranking','bnav-analisis'];
function switchVista(vistaId, navId, cb){
  VISTAS.forEach(function(v){document.getElementById(v).style.display='none';});
  NAVS.forEach(function(n){document.getElementById(n).classList.remove('nav-active');});
  BNAVS.forEach(function(n){var el=document.getElementById(n);if(el)el.classList.remove('nav-active');});
  document.getElementById(vistaId).style.display='block';
  document.getElementById(navId).classList.add('nav-active');
  var bnav = document.getElementById(navId.replace('nav-','bnav-'));
  if(bnav) bnav.classList.add('nav-active');
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

function seleccionarMesAna(mes){
  ANA_MES = mes;
  renderAnalisis();
}

function renderAnalisis(){
  renderChipsAna();
  if(!ANA_MES || !ANA_DATA.por_mes || !ANA_DATA.por_mes[ANA_MES]){
    document.getElementById('ana-diagnostico').innerHTML = '<div style="padding:24px 16px;color:#999;font-size:13px">Sin datos para el período seleccionado.</div>';
    document.getElementById('ana-metricas').innerHTML = '';
    document.getElementById('ana-comparativa').innerHTML = '';
    document.getElementById('ana-por-semana').innerHTML = '';
    document.getElementById('ana-por-dia').innerHTML = '';
    document.getElementById('ana-calendario').innerHTML = '';
    document.getElementById('ana-tabla').innerHTML = '';
    return;
  }
  var d = ANA_DATA.por_mes[ANA_MES];
  renderDiagnostico(d);
  renderMetricasAna(d);
  renderComparativaMeses();
  renderPorSemana(d);
  renderPorDia(d);
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
    +'<div class="metrica-valor '+(diff>=0?'val-verde':'val-rojo')+'">'+total_f+'</div>'
    +'<div class="metrica-sub">'+(diff>=0?'+':'')+diff+'% vs promedio</div></div>'
    +'<div class="metrica"><div class="metrica-label">Días con quiebre</div>'
    +'<div class="metrica-valor '+(totalQ>0?'val-rojo':'val-verde')+'">'+totalQ+'</div>'
    +'<div class="metrica-sub">'+prodsQ+' producto'+(prodsQ!==1?'s':'')+' afectado'+(prodsQ!==1?'s':'')+'</div></div>'
    +'<div class="metrica"><div class="metrica-label">Días especiales</div>'
    +'<div class="metrica-valor" style="color:#1960a6">'+diasEsp+'</div>'
    +'<div class="metrica-sub">feriados + vacaciones</div></div>'
    +'<div class="metrica"><div class="metrica-label">Semanas activas</div>'
    +'<div class="metrica-valor">'+semsAct+'</div>'
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
    var h  = maxV>0 ? Math.max(4, Math.round(v/maxV*76)) : 4;
    var act = m === ANA_MES;
    var bg  = act ? '#275300' : '#c2c9b7';
    var lbl = MNA[parseInt(m.split('-')[1])];
    html += '<div class="month-col" data-mes="'+m+'" onclick="seleccionarMesAna(this.dataset.mes)">'
      +'<div class="month-bar" style="height:'+h+'px;background:'+bg+'"></div>'
      +'<span class="month-lab" style="'+(act?'color:#275300;font-weight:600':'')+'">'+(lbl||m)+'</span></div>';
  }
  html += '</div><div style="font-size:11px;color:#727969">Promedio: <strong style="color:#191c1d">'+(ANA_DATA.promedio_mensual.toLocaleString?ANA_DATA.promedio_mensual.toLocaleString('es-CL'):ANA_DATA.promedio_mensual)+' un.</strong> · Clic en barra para ver ese mes</div>';
  document.getElementById('ana-comparativa').innerHTML = html;
}

function renderPorSemana(d){
  var max = Math.max.apply(null, d.por_semana);
  if(max===0){ document.getElementById('ana-por-semana').innerHTML='<p style="color:#999;font-size:12px">Sin datos</p>'; return; }
  var semsPos = d.por_semana.filter(function(v){return v>0;});
  var minPos  = semsPos.length>0 ? Math.min.apply(null,semsPos) : 0;
  var html = '';
  for(var i=0;i<d.por_semana.length;i++){
    var v = d.por_semana[i];
    if(v===0) continue;
    var pct = Math.round(v/max*100);
    var bg  = v===max ? '#275300' : (v===minPos && semsPos.length>=3 ? '#ba1a1a' : '#c2c9b7');
    html += '<div class="bar-row"><span class="bar-lbl">Sem '+(i+1)+'</span>'
      +'<div class="bar-track"><div class="bar-fill" style="width:'+pct+'%;background:'+bg+'"></div></div>'
      +'<span class="bar-val">'+v+' un.</span></div>';
  }
  document.getElementById('ana-por-semana').innerHTML = html;
}

function renderPorDia(d){
  var dias   = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'];
  var colors = ['#c2c9b7','#c2c9b7','#c2c9b7','#c2c9b7','#1960a6','#275300','#1960a6'];
  var max    = Math.max.apply(null, d.por_dia);
  var html   = '';
  for(var i=0;i<7;i++){
    var v   = d.por_dia[i];
    var pct = max>0 ? Math.round(v/max*100) : 0;
    html += '<div class="bar-row"><span class="bar-lbl">'+dias[i]+'</span>'
      +'<div class="bar-track"><div class="bar-fill" style="width:'+pct+'%;background:'+colors[i]+'"></div></div>'
      +'<span class="bar-val">'+v+'%</span></div>';
  }
  document.getElementById('ana-por-dia').innerHTML = html;
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

  var MESES_NOM = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  var moNom = MESES_NOM[mo - 1];

  // getDay(): 0=Dom, 1=Lun, ..., 6=Sáb → convertir a Lun=0 … Dom=6
  var primerDia = new Date(yr, mo - 1, 1).getDay();
  var offset    = (primerDia + 6) % 7;  // Lun=0, Dom=6

  var HDRS = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'];
  var html = '<div class="cal-header">';
  for(var h=0;h<7;h++) html += '<div class="cal-hdr">'+HDRS[h]+'</div>';
  html += '</div><div class="cal-grid">';

  // Celdas vacías antes del primer día
  for(var b=0;b<offset;b++) html += '<div class="tday empty"></div>';

  for(var dd=1;dd<=nDias;dd++){
    var cls = 'tday ok', tip = '';
    if(fSet[dd]){     cls='tday feriado';   tip=dd+' de '+moNom+' · '+fSet[dd]; }
    else if(festSet[dd]){ cls='tday festividad'; tip=dd+' de '+moNom+' · '+festSet[dd]; }
    else if(vSet[dd]){ cls='tday vacacion'; tip=dd+' de '+moNom+' · '+vSet[dd]; }
    html += '<div class="'+cls+'" data-tip="'+tip+'">'+dd+'</div>';
  }
  html += '</div>';

  var hasFest = Object.keys(festSet).length > 0;
  var hasVac  = Object.keys(vSet).length  > 0;
  html += '<div class="leyenda-tl">'
    +'<div class="leg-tl-item"><div class="leg-tl-dot" style="background:#e8e8e5;border:1px solid #c2c9b7"></div>Normal</div>'
    +'<div class="leg-tl-item"><div class="leg-tl-dot" style="background:#1960a6"></div>Feriado</div>'
    +(hasFest?'<div class="leg-tl-item"><div class="leg-tl-dot" style="background:#b8f389;border:1px solid #3b6d11"></div>Festividad</div>':'')
    +(hasVac?'<div class="leg-tl-item"><div class="leg-tl-dot" style="background:#ffd166"></div>Vacaciones</div>':'')
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
  var html = '<table class="rank-table"><thead><tr>'
    +'<th data-col="nombre" onclick="sortTablaAna(this.dataset.col)" style="cursor:pointer;user-select:none">Producto'+arw('nombre')+'</th>'
    +'<th data-col="pct" onclick="sortTablaAna(this.dataset.col)" style="cursor:pointer;user-select:none;text-align:right">% Total'+arw('pct')+'</th>'
    +'<th data-col="total" onclick="sortTablaAna(this.dataset.col)" style="cursor:pointer;user-select:none;text-align:right">Unidades'+arw('total')+'</th>'
    +'<th data-col="quiebre" onclick="sortTablaAna(this.dataset.col)" style="cursor:pointer;user-select:none;text-align:right">Quiebre'+arw('quiebre')+'</th>'
    +'<th style="min-width:60px">Últ.6 meses</th>'
    +'<th data-col="tendencia" onclick="sortTablaAna(this.dataset.col)" style="cursor:pointer;user-select:none">Tendencia'+arw('tendencia')+'</th>'
    +'</tr></thead><tbody>';
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
    html += '<tr>'
      +'<td>'+p.nombre+'<span style="color:#aaa;font-size:10px;margin-left:4px">'+p.sku+'</span></td>'
      +'<td style="text-align:right;font-weight:500">'+p.pct+'%</td>'
      +'<td style="text-align:right">'+p.total+'</td>'
      +'<td style="text-align:right;color:'+qColor+';font-weight:'+(p.dias_quiebre>2?'600':'400')+'">'+qText+'</td>'
      +'<td>'+spark+'</td>'
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
  var el = document.getElementById('res-alertas');
  if(!el) return;
  if(!ALERTAS || ((ALERTAS.entradas||[]).length===0 && (ALERTAS.salidas||[]).length===0)){
    el.innerHTML = ''; return;
  }
  var h = '';
  if((ALERTAS.salidas||[]).length>0){
    var items = ALERTAS.salidas.map(function(a){
      return '<div class="res-item"><span class="res-item-nombre">'+nombreSku(a.sku)
        +' <span style="color:#888;font-size:11px">('+(a.oficina==='VIT'?'Vitacura':'Pataguas')+')</span></span>'
        +'<span style="color:#E67E22;font-weight:700">-'+Math.round(a.cantidad)+' und</span></div>';
    }).join('');
    h += '<div style="margin:16px 16px 0;background:#FFF7EE;border:1px solid #F0CFA8;border-radius:10px;padding:14px 18px">'
      +'<div style="font-weight:700;font-size:13px;color:#B9650F;margin-bottom:6px">⚠️ Salidas sin explicación — '+(ALERTAS.fecha||'')+'</div>'
      +'<div style="font-size:12px;color:#8a6a45;margin-bottom:8px">El stock bajó sin ventas ni guías que lo expliquen. '
      +'Puede ser consumo interno o merma; si no lo reconoces, revisa la tarjeta de existencia del producto en Bsale.</div>'
      +items+'</div>';
  }
  if((ALERTAS.entradas||[]).length>0){
    var items2 = ALERTAS.entradas.map(function(a){
      return '<div class="res-item"><span class="res-item-nombre">'+nombreSku(a.sku)
        +' <span style="color:#888;font-size:11px">('+(a.oficina==='VIT'?'Vitacura':'Pataguas')+')</span></span>'
        +'<span style="color:#27AE60;font-weight:700">+'+Math.round(a.cantidad)+' und</span></div>';
    }).join('');
    h += '<div style="margin:16px 16px 0;background:#F2FAF4;border:1px solid #BFE3C8;border-radius:10px;padding:14px 18px">'
      +'<div style="font-weight:700;font-size:13px;color:#1E8449;margin-bottom:6px">📦 Recepciones detectadas — '+(ALERTAS.fecha||'')+'</div>'
      +'<div style="font-size:12px;color:#5a7a62;margin-bottom:8px">El stock subió: se registraron como producción del día. Normal si hubo recepción.</div>'
      +items2+'</div>';
  }
  el.innerHTML = h;
}

// ── Resumen ejecutivo ──────────────────────────────────────
function renderResumen(){
  renderAlertas();
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

  // Fila de stats
  var sinStk = DATA.filter(function(p){return p.estado==='sin_stock';}).length;
  var crit   = DATA.filter(function(p){return p.estado==='critico';}).length;
  var statsHtml = '<div class="res-stats-row">'
    +'<div class="res-stat"><div class="res-stat-val val-rojo">'+sinStk+'</div><div class="res-stat-label">Sin stock</div></div>'
    +'<div class="res-stat"><div class="res-stat-val val-rojo">'+crit+'</div><div class="res-stat-label">Críticos</div></div>'
    +'<div class="res-stat"><div class="res-stat-val val-amarillo">'+bajos.length+'</div><div class="res-stat-label">Bajo stock</div></div>'
    +'<div class="res-stat"><div class="res-stat-val" style="color:#1A1A1A">'+despachos.length+'</div><div class="res-stat-label">Despachos</div></div>'
    +'<div class="res-stat"><div class="res-stat-val val-verde">'+totalVit+'</div><div class="res-stat-label">Und VIT</div></div>'
    +'<div class="res-stat"><div class="res-stat-val" style="color:#185FA5">'+totalPat+'</div><div class="res-stat-label">Und PAT</div></div>'
    +'</div>';
  document.getElementById('res-stats').innerHTML = statsHtml;

  // Grid de cards
  var urgHtml = urgentes.length===0
    ? '<div style="color:#27AE60;font-size:13px;padding:8px 0;font-weight:500">Sin urgencias — todo bajo control</div>'
    : urgentes.map(function(p){
        var dias_s = p.estado==='sin_stock'?'Sin stock':Math.round(p.dias_prod)+'d';
        var clr    = p.estado==='sin_stock'?'#E74C3C':'#E67E22';
        return '<div class="res-item"><span class="res-item-nombre">'+p.nombre+'</span>'
          +'<span style="color:'+clr+';font-weight:700;font-size:12px">'+dias_s+'</span></div>';
      }).join('');

  var despHtml = despachos.length===0
    ? '<div style="color:#aaa;font-size:13px;padding:8px 0">Sin despachos pendientes</div>'
    : despachos.map(function(p){
        var nec  = Math.max(0, Math.ceil(p.vel_pat*dias)-p.pat);
        var res  = Math.ceil(p.vel_vit*(p.tiempo_repo||7));
        var disp = Math.max(0, p.vit-res);
        var desp = Math.min(nec,disp);
        return '<div class="res-item"><span class="res-item-nombre">'+p.nombre+'</span>'
          +'<span style="color:#27AE60;font-weight:700">+'+desp+' und</span></div>';
      }).join('');

  var bajosHtml = bajos.length===0
    ? '<div style="color:#aaa;font-size:13px;padding:8px 0">Ninguno</div>'
    : bajos.map(function(p){
        return '<div class="res-item"><span class="res-item-nombre">'+p.nombre+'</span>'
          +'<span style="color:#E67E22;font-weight:600;font-size:12px">'+Math.round(p.dias_prod)+'d</span></div>';
      }).join('');

  document.getElementById('res-grid').innerHTML =
    '<div class="res-card"><div class="res-card-title">Producir urgente ('+urgentes.length+')</div>'+urgHtml+'</div>'
   +'<div class="res-card"><div class="res-card-title">Despachar a Pataguas ('+despachos.length+')</div>'+despHtml+'</div>'
   +'<div class="res-card"><div class="res-card-title">Stock bajo ('+bajos.length+')</div>'+bajosHtml+'</div>'
   +'<div class="res-card"><div class="res-card-title">Totales en stock</div>'
   +'<div class="res-item"><span class="res-item-nombre">Vitacura</span><span style="font-weight:700">'+totalVit+' und</span></div>'
   +'<div class="res-item"><span class="res-item-nombre">Pataguas</span><span style="color:#185FA5;font-weight:700">'+totalPat+' und</span></div>'
   +'<div class="res-item"><span class="res-item-nombre">Total general</span><span style="font-weight:700">'+(totalVit+totalPat)+' und</span></div>'
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
    var clr   = cob<=3 ? '#E74C3C' : cob<=7 ? '#E67E22' : '#27AE60';
    var urg   = cob<=3;
    return '<tr'+(urg?' class="urgente"':'')+'>'
      +'<td style="font-weight:'+(urg?700:400)+'">'+p.nombre+'</td>'
      +'<td style="color:'+clr+';font-weight:700;width:55px">'+dt+'</td>'
      +'<td style="color:#888;font-size:11px">'+p.cocinero+'</td>'
      +'<td style="text-align:right;color:#888">'+p.vit+'</td>'
      +'<td style="text-align:right;color:#185FA5">'+p.pat+'</td>'
      +'<td style="text-align:right;font-weight:600">'+total+'</td>'
      +'<td style="text-align:right;font-weight:700;font-size:13px;color:'+(und>0?clr:'#bbb')+'">'+und_str+'</td>'
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
    var urg = e.estado!=='ok' && (p.pat===0 || (p.pat/p.vel_pat)<3);
    var cell;
    if(e.estado==='ok')            cell='<span style="color:#27AE60;font-weight:700">OK</span>';
    else if(e.estado==='completo') cell='<span style="color:#27AE60;font-weight:700">+'+e.desp+'</span>';
    else if(e.estado==='parcial')  cell='<span style="color:#E67E22;font-weight:700">+'+e.desp+'</span>';
    else                           cell='<span style="color:#999;font-weight:700">'+e.necesita+'</span>';
    return '<tr'+(urg?' class="urgente"':'')+'>'
      +'<td style="font-weight:'+(urg?700:400)+'">'+p.nombre+'</td>'
      +'<td style="text-align:right;color:'+(p.vit===0?'#E74C3C':'#333')+'">'+p.vit+'</td>'
      +'<td style="text-align:right;color:#185FA5">'+p.pat+'</td>'
      +'<td style="text-align:right;color:#888">'+dpt+'</td>'
      +'<td style="text-align:right;font-size:13px">'+cell+'</td>'
      +'</tr>';
  }).join('');

  document.getElementById('tabla-despacho').innerHTML = filas||'<tr><td colspan="5" style="text-align:center;color:#aaa;padding:20px">Sin productos</td></tr>';
  document.getElementById('resumen-despacho').textContent =
    despachar+' a despachar · '+producir+' a producir · '+dias+' días';
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
document.getElementById('f-buscar').addEventListener('input', filtrar);
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

<div class="header">
  <div class="logo">
    <div class="logo-icon">🍽</div>
    <div><span class="logo-nombre">La Cocina</span><span class="logo-sub">· Control de Producción</span></div>
  </div>
  <div class="header-right">
    <div class="header-nav-btns">
      <button class="btn nav-active" id="nav-resumen" onclick="mostrarResumen()">Resumen</button>
      <button class="btn" id="nav-productos" onclick="mostrarProductos()">Productos</button>
      <button class="btn" id="nav-guias" onclick="mostrarGuias()">Guías</button>
      <button class="btn" id="nav-ranking" onclick="mostrarRanking()">Ranking</button>
      <button class="btn" id="nav-analisis" onclick="mostrarAnalisis()" style="display:none">Análisis</button>
    </div>
    <span class="fecha">FECHA_HOY_PLACEHOLDER</span>
  </div>
</div>

<nav class="bottom-nav">
  <button class="bnav-btn nav-active" id="bnav-resumen" onclick="mostrarResumen()"><span class="bnav-icon">📊</span>Resumen</button>
  <button class="bnav-btn" id="bnav-productos" onclick="mostrarProductos()"><span class="bnav-icon">📦</span>Productos</button>
  <button class="bnav-btn" id="bnav-guias" onclick="mostrarGuias()"><span class="bnav-icon">📋</span>Guías</button>
  <button class="bnav-btn" id="bnav-ranking" onclick="mostrarRanking()"><span class="bnav-icon">📈</span>Ranking</button>
  <button class="bnav-btn" id="bnav-analisis" onclick="mostrarAnalisis()" style="display:none"><span class="bnav-icon">📉</span>Análisis</button>
</nav>

<!-- VISTA RESUMEN -->
<div id="vista-resumen">
  <div id="res-alertas"></div>
  <div id="res-stats"></div>
  <div class="res-grid" id="res-grid"></div>
</div>

<!-- VISTA PRODUCTOS -->
<div id="vista-productos" style="display:none">

<div class="metricas">
  <div class="metrica"><div class="metrica-label">Sin stock</div><div class="metrica-valor val-rojo" id="m1">—</div></div>
  <div class="metrica"><div class="metrica-label">Crítico ≤3d</div><div class="metrica-valor val-rojo" id="m2">—</div></div>
  <div class="metrica"><div class="metrica-label">Bajo ≤14d</div><div class="metrica-valor val-amarillo" id="m3">—</div></div>
  <div class="metrica"><div class="metrica-label">OK &gt;14d</div><div class="metrica-valor val-verde" id="m4">—</div></div>
</div>
  <div class="toolbar">
    <div class="search-wrap">
      <span class="search-ico">🔍</span>
      <input type="text" id="f-buscar" placeholder="Buscar producto o SKU...">
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
  <div class="leyenda">
    <div class="leg"><div class="leg-dot" style="background:#E74C3C"></div>Sin stock</div>
    <div class="leg"><div class="leg-dot" style="background:#E67E22"></div>Crítico</div>
    <div class="leg"><div class="leg-dot" style="background:#F39C12"></div>Bajo stock</div>
    <div class="leg"><div class="leg-dot" style="background:#27AE60"></div>OK</div>
    <div class="leg" style="margin-left:16px"><div class="leg-dot" style="background:#EAF3DE;border:1px solid #639922"></div>Vitacura</div>
    <div class="leg"><div class="leg-dot" style="background:#E6F1FB;border:1px solid #185FA5"></div>Pataguas</div>
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
  </div>
  <div style="padding:12px 20px">
    <table class="guia-table">
      <thead><tr>
        <th>Producto</th><th style="width:55px">Días</th><th>Cocinero</th>
        <th style="text-align:right">Vitacura</th><th style="text-align:right">Pataguas</th>
        <th style="text-align:right">Total</th><th style="text-align:right">Producir</th>
      </tr></thead>
      <tbody id="tabla-produccion"></tbody>
    </table>
  </div>

  <!-- Guía Despacho -->
  <div class="guia-section" style="margin-top:12px">
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
    <div style="display:flex;gap:18px;padding:8px 0 4px;font-size:12px;color:#555;flex-wrap:wrap">
      <span><span style="display:inline-block;width:12px;height:12px;background:#27AE60;border-radius:3px;margin-right:5px;vertical-align:middle"></span><b style="color:#27AE60">+N</b> Despacho completo — Vitacura cubre los días pedidos en Pataguas</span>
      <span><span style="display:inline-block;width:12px;height:12px;background:#E67E22;border-radius:3px;margin-right:5px;vertical-align:middle"></span><b style="color:#E67E22">+N</b> Despacho parcial — manda lo posible sin dejar Vitacura en cero</span>
      <span><span style="display:inline-block;width:12px;height:12px;background:#999;border-radius:3px;margin-right:5px;vertical-align:middle"></span><b style="color:#999">N</b> Producir — falta en Pataguas pero Vitacura no tiene stock</span>
      <span><b style="color:#27AE60">OK</b> — Pataguas tiene suficiente, no hace falta despachar</span>
    </div>
  </div>
  <div style="padding:12px 20px">
    <table class="guia-table">
      <thead><tr>
        <th>Producto</th><th style="text-align:right">Stock VIT</th>
        <th style="text-align:right">Stock PAT</th>
        <th style="text-align:right">Días PAT</th>
        <th style="text-align:right">Despachar</th>
      </tr></thead>
      <tbody id="tabla-despacho"></tbody>
    </table>
  </div>

</div>

<!-- VISTA ANÁLISIS -->
<div id="vista-analisis" style="display:none">

<div class="ana-chips" id="ana-chips"></div>

<div id="ana-diagnostico"></div>

<div class="ana-metricas" id="ana-metricas">
  <div class="metrica"><div class="metrica-label">Cargando…</div></div>
</div>

<div class="ana-grid">
  <div class="ana-card">
    <div class="ana-card-title">Comparativa mensual</div>
    <div id="ana-comparativa"></div>
  </div>
  <div class="ana-card">
    <div class="ana-card-title">Por semana del mes</div>
    <div id="ana-por-semana"></div>
  </div>
  <div class="ana-card">
    <div class="ana-card-title">Por día de la semana</div>
    <div id="ana-por-dia"></div>
  </div>
  <div class="ana-card">
    <div class="ana-card-title">Días del mes</div>
    <div id="ana-calendario"></div>
  </div>
</div>

<div class="ana-section">
  <div class="ana-section-title">Contribución por producto</div>
  <div class="tabla-wrap">
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

def generar_html(datos, fecha_str, analisis=None):
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
    html = html.replace('CSS_PLACEHOLDER',         CSS)
    html = html.replace('ANA_DATA_PLACEHOLDER',    analisis_json)  # antes de DATA_PLACEHOLDER
    html = html.replace('DATA_PLACEHOLDER',        data_json)
    html = html.replace('ALERTAS_PLACEHOLDER',     alertas_json)
    html = html.replace('JS_PLACEHOLDER',          js_final)
    html = html.replace('FECHA_HOY_PLACEHOLDER',   fecha_str)
    return html

# ─── MAIN ────────────────────────────────────────────────────
if __name__ == '__main__':
    print('='*50)
    print('La Cocina — Generador de Dashboard')
    print('='*50)
    datos = procesar()
    print('\nCalculando datos de análisis...')
    analisis = calcular_analisis()
    print(f'  Meses disponibles: {len(analisis.get("meses", []))}')
    print(f'\nGenerando HTML con {len(datos)} productos...')
    html = generar_html(datos, FECHA_STR, analisis)
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
