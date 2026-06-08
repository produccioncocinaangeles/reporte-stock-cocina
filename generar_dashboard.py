#!/usr/bin/env python3
"""
La Cocina — Generador de Dashboard v2
Lee consolidados Vitacura/Pataguas + stock Bsale API
Genera dashboard.html con datos embebidos
"""
import os, json, requests, math
import pandas as pd
import numpy as np
from collections import Counter

# ─── CONFIGURACIÓN ───────────────────────────────────────────
CARPETA      = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_VIT  = os.path.join(CARPETA, 'consolidado_productos_vitacura.xlsx')
ARCHIVO_PAT  = os.path.join(CARPETA, 'consolidado_productos_pataguas.xlsx')
ARCHIVO_HTML = os.path.join(CARPETA, 'dashboard.html')
BSALE_TOKEN  = os.environ.get('BSALE_TOKEN', '')
_HOY         = pd.Timestamp.today().normalize()
FECHA_HOY    = _HOY
FECHA_STR    = _HOY.strftime('%d/%m/%Y')

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
    result  = {}
    offset  = 0
    while True:
        r = requests.get('https://api.bsale.cl/v1/stocks.json', headers=headers,
                         params={'limit':50,'offset':offset,'expand':'[variant,office]'}, timeout=30)
        if r.status_code != 200:
            break
        data  = r.json()
        items = data.get('items', [])
        for item in items:
            try:
                sku      = ' '.join(str(item['variant']['code']).split()).upper()
                sucursal = ' '.join(str(item['office']['name']).split()).upper()
                qty      = int(item.get('quantityAvailable', 0) or 0)
                if sku not in result:
                    result[sku] = {'vit':0,'pat':0}
                if   sucursal == 'VITACURA':    result[sku]['vit'] += qty
                elif sucursal == 'LAS PATAGUAS': result[sku]['pat'] += qty
            except: pass
        if offset + 50 >= data.get('count', 0): break
        offset += 50
    return result

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

def velocidad(df_sku, stock_d):
    mask = (df_sku['Salida'] > 0) & (~df_sku['Movimiento de salida'].str.contains(
        'GUÍA DE DESPACHO|Guía de Despacho|Consumo', na=False))
    total    = int(df_sku[mask]['Salida'].sum())
    con_stk  = sum(1 for v in stock_d.values() if v > 0)
    vel      = round(total / con_stk, 4) if con_stk > 0 else 0.0
    return vel, total, con_stk

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
    print("Leyendo consolidados...")
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
        trepo = moda(durs) if durs else (moda([p['dias'] for p in per_v]) if per_v else 7)

        # Velocidades
        vel_v, tot_v, dsc_v = velocidad(dv, sd_v)
        vel_p, tot_p, dsc_p = velocidad(dp, sd_p)
        vel_t = vel_v + vel_p

        # Días restantes
        total_hoy = sv_hoy + sp_hoy
        dias_v = round(sv_hoy/vel_v,1) if vel_v>0 and sv_hoy>0 else (0 if sv_hoy==0 else None)
        dias_p = round(sp_hoy/vel_p,1) if vel_p>0 and sp_hoy>0 else (0 if sp_hoy==0 else None)
        dias_t = round(total_hoy/vel_t,1) if vel_t>0 else None

        # Estado
        if   total_hoy == 0:                                   estado = 'sin_stock'
        elif dias_t is not None and dias_t <= 3:               estado = 'critico'
        elif dias_t is not None and dias_t <= 14:              estado = 'bajo'
        else:                                                  estado = 'ok'

        alerta = (sv_hoy==0 and sp_hoy>0 and vel_v>0) or (sp_hoy==0 and sv_hoy>0 and vel_p>0)
        pto_reorden   = math.ceil(vel_t * trepo)              if vel_t > 0 else 0
        lote_sugerido = math.ceil(vel_t * max(trepo+7, 14))   if vel_t > 0 else 0
        despacho      = max(0, math.ceil(vel_p * max(trepo,7)) - sp_hoy) if vel_p > 0 else 0

        lts, prom_lote, n_lotes = lotes(dv)

        # Ventas por mes combinadas
        vm_v = ventas_mes(dv)
        vm_p = ventas_mes(dp)
        meses = sorted(set(list(vm_v)+list(vm_p)))
        vm = [{'mes':m,'vit':vm_v.get(m,0),'pat':vm_p.get(m,0),'total':vm_v.get(m,0)+vm_p.get(m,0)} for m in meses]

        resultados.append({
            'sku': sku, 'nombre': NOMBRES[sku], 'cocinero': COCINEROS.get(sku,''),
            'vit': sv_hoy, 'pat': sp_hoy, 'total': total_hoy,
            'vel_vit': vel_v, 'vel_pat': vel_p, 'vel_total': round(vel_t,4),
            'dias_vit': dias_v, 'dias_pat': dias_p, 'dias_total': dias_t,
            'estado': estado, 'alerta_dist': alerta,
            'tiempo_repo': trepo, 'pto_reorden': pto_reorden,
            'lote_sugerido': lote_sugerido, 'despacho_sug': despacho,
            'total_vit': tot_v, 'total_pat': tot_p,
            'dias_stock_vit': dsc_v, 'dias_stock_pat': dsc_p,
            'n_lotes': n_lotes, 'prom_lote': prom_lote,
            'periodos_sin_stock_vit': per_v,
            'lotes': lts, 'movs': movimientos(dv, dp), 'ventas_mes': vm,
        })
        print(f"  ✓ {sku} — {NOMBRES[sku][:30]}")

    orden = {'sin_stock':0,'critico':1,'bajo':2,'ok':3}
    resultados.sort(key=lambda x: (orden[x['estado']], x['dias_total'] if x['dias_total'] is not None else 9999))
    return resultados

# ─── HTML (sin f-string para evitar conflictos con JS) ───────
CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#F5F4F0;color:#1A1A1A;font-size:14px;max-width:900px;margin:0 auto}
.header{background:#fff;border-bottom:1px solid #ddd;padding:14px 20px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}
.logo{font-size:17px;font-weight:700;letter-spacing:-0.02em}.logo span{color:#E24B4A}
.header-right{display:flex;gap:10px;align-items:center}
.btn{font-size:12px;padding:7px 16px;border-radius:6px;border:1px solid #ddd;cursor:pointer;background:#fff;color:#444;font-family:inherit}
.btn:hover{background:#f0f0f0}
.btn-primary{background:#1A1A1A;color:#fff;border-color:#1A1A1A}.btn-primary:hover{background:#333}
.nav-active{background:#1A1A1A!important;color:#fff!important;border-color:#1A1A1A!important}
.fecha{font-size:11px;color:#888}
.metricas{display:grid;grid-template-columns:repeat(4,1fr);background:#fff;border-bottom:1px solid #ddd}
.metrica{padding:16px 20px;border-right:1px solid #ddd}.metrica:last-child{border-right:none}
.metrica-label{font-size:11px;color:#888;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px}
.metrica-valor{font-size:30px;font-weight:700;line-height:1}
.metrica-sub{font-size:11px;color:#999;margin-top:4px}
.val-rojo{color:#E24B4A}.val-amarillo{color:#EF9F27}.val-verde{color:#27AE60}
.toolbar{background:#fff;border-bottom:1px solid #ddd;padding:10px 16px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.toolbar-label{font-size:11px;color:#888;text-transform:uppercase;letter-spacing:0.06em}
select,input[type=text],input[type=number]{font-size:12px;padding:6px 10px;border:1px solid #ddd;border-radius:5px;background:#fff;color:#333;font-family:inherit}
input[type=text]{width:200px}
.toolbar-sep{width:1px;height:20px;background:#ddd;margin:0 4px}
.leyenda{padding:8px 16px;display:flex;gap:16px;align-items:center;background:#F5F4F0;border-bottom:1px solid #ddd;flex-wrap:wrap}
.leg{display:flex;align-items:center;gap:5px;font-size:11px;color:#666}
.leg-dot{width:10px;height:10px;border-radius:2px}
.container{padding:12px 16px}
.card{background:#fff;border:1px solid #ddd;border-radius:10px;margin-bottom:8px;overflow:hidden;transition:box-shadow 0.15s}
.card:hover{box-shadow:0 2px 10px rgba(0,0,0,0.08)}
.card.sin_stock{border-left:4px solid #E74C3C}
.card.critico{border-left:4px solid #E74C3C}
.card.bajo{border-left:4px solid #E67E22}
.card.ok{border-left:4px solid #27AE60}
.card-top{padding:12px 14px;display:flex;align-items:center;justify-content:space-between;cursor:pointer;user-select:none}
.card-info{flex:1;min-width:0}
.card-nombre{font-size:13px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.card-meta{font-size:11px;color:#888;margin-top:2px}
.card-badges{display:flex;align-items:center;gap:8px;flex-shrink:0;margin-left:10px}
.badge{font-size:11px;font-weight:600;padding:3px 10px;border-radius:20px}
.badge-rojo{background:#FCEBEB;color:#A32D2D}
.badge-amarillo{background:#FAEEDA;color:#854F0B}
.badge-verde{background:#EAF3DE;color:#3B6D11}
.badge-dist{background:#FEF3C7;color:#92400E;font-size:10px}
.chevron{font-size:12px;color:#bbb;margin-left:6px;transition:transform 0.2s}
.chevron.open{transform:rotate(180deg)}
.tl-wrap{padding:6px 14px 10px}

.btn-det{font-size:11px;background:none;border:none;color:#2563EB;padding:6px 14px;cursor:pointer;width:100%;text-align:left;border-top:1px solid #f0f0f0}
.btn-det:hover{background:#f8f8f8}
.detalle{display:none;border-top:1px solid #eee}
.detalle.open{display:block}
.tabs{display:flex;border-bottom:1px solid #eee;background:#fafafa}
.tab{font-size:11px;padding:9px 16px;border:none;background:none;color:#888;cursor:pointer;font-family:inherit;border-bottom:2px solid transparent}
.tab.active{color:#1A1A1A;border-bottom-color:#1A1A1A;font-weight:600}
.tab-body{display:none;padding:14px}
.tab-body.active{display:block}
.movs-table{width:100%;border-collapse:collapse;font-size:12px}
.movs-table th{text-align:left;color:#888;padding:6px 10px;border-bottom:1px solid #eee;font-size:10px;text-transform:uppercase;letter-spacing:0.05em;font-weight:600}
.movs-table td{padding:7px 10px;border-bottom:1px solid #f5f5f5;vertical-align:middle}
.movs-table tr:last-child td{border-bottom:none}
.movs-table tr:hover td{background:#fafafa}
.tipo-badge{display:inline-block;font-size:10px;font-weight:600;padding:2px 8px;border-radius:4px}
.tipo-prod{background:#EAF3DE;color:#3B6D11}
.tipo-venta{background:#FEE2E2;color:#991B1B}
.tipo-despacho{background:#EFF6FF;color:#1D4ED8}
.tipo-consumo{background:#FEF3C7;color:#92400E}
.tipo-desp-rec{background:#F0FDF4;color:#166534}
.tienda-vit{font-size:10px;padding:1px 6px;border-radius:3px;background:#EAF3DE;color:#3B6D11;font-weight:500}
.tienda-pat{font-size:10px;padding:1px 6px;border-radius:3px;background:#E6F1FB;color:#185FA5;font-weight:500}
.cant-pos{color:#059669;font-weight:700}
.cant-neg{color:#DC2626;font-weight:700}
.insight{background:#F8F7F4;border-radius:8px;padding:12px 14px;font-size:12px;color:#555;line-height:1.8;margin-bottom:10px}
.insight b{color:#1A1A1A;font-weight:600}
.insight-warn{background:#FFFBEB;border:1px solid #FDE68A;border-radius:8px;padding:12px 14px;font-size:12px;color:#92400E;line-height:1.8;margin-bottom:10px}
.insight-ok{background:#EAF3DE;border-radius:8px;padding:12px 14px;font-size:12px;color:#3B6D11;line-height:1.8;margin-bottom:10px}
.insight-peligro{background:#FEE2E2;border-radius:8px;padding:12px 14px;font-size:12px;color:#991B1B;line-height:1.8;margin-bottom:10px}
.periodo-chip{display:inline-block;font-size:10px;padding:2px 8px;border-radius:4px;background:#FEE2E2;color:#991B1B;margin:2px}
.lote-card{background:#fff;border:1px solid #eee;border-radius:6px;padding:10px 12px;font-size:11px;margin-bottom:6px}
.mes-table{width:100%;border-collapse:collapse;font-size:12px;margin-top:8px}
.mes-table th{text-align:left;color:#888;padding:5px 10px;border-bottom:1px solid #eee;font-size:10px;text-transform:uppercase}
.mes-table td{padding:6px 10px;border-bottom:1px solid #f5f5f5}
.mes-bar{display:inline-block;height:8px;background:#5DCAA5;border-radius:2px;margin-left:6px;vertical-align:middle}
.no-res{text-align:center;color:#aaa;font-size:13px;padding:40px}
.dias-btn{font-size:11px;padding:5px 10px;border:none;background:none;border-radius:6px;cursor:pointer;color:#666;font-family:inherit}
.dias-btn:hover{background:#e5e5e5}
.dias-btn-active{background:#fff;color:#1A1A1A;font-weight:600;box-shadow:0 1px 3px rgba(0,0,0,0.15)}
.guia-section{background:#fff;border-bottom:1px solid #ddd;padding:14px 20px}
.guia-header{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px}
.guia-title{font-size:15px;font-weight:700;margin-bottom:2px}
.guia-sub{font-size:11px;color:#888}
.guia-controls{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.dias-group{display:flex;align-items:center;gap:6px;background:#F5F4F0;border-radius:8px;padding:4px}
.guia-table{width:100%;border-collapse:collapse;font-size:12px}
.guia-table th{text-align:left;color:#888;padding:8px 10px;border-bottom:2px solid #eee;font-size:10px;text-transform:uppercase}
.guia-table td{padding:7px 10px;border-bottom:1px solid #f5f5f5}
.guia-table tr.urgente{background:#FFF5F5}
"""

JS = """
const MESES_L = {'2026-01':'Enero','2026-02':'Febrero','2026-03':'Marzo',
  '2026-04':'Abril','2026-05':'Mayo','2026-06':'Junio'};

// ── Helpers ────────────────────────────────────────────────
function diasStr(p){
  if(p.estado==='sin_stock') return 'SIN STOCK';
  if(p.dias_total===null||p.dias_total===undefined) return '—';
  if(p.dias_total>365) return '+1 año';
  return Math.round(p.dias_total)+'d';
}
function badgeCls(e){
  if(e==='sin_stock'||e==='critico') return 'badge-rojo';
  if(e==='bajo') return 'badge-amarillo';
  return 'badge-verde';
}

// ── Barra estilo escala con marcador ───────────────────────
function buildBarra(p){
  const dias = p.dias_total !== null && p.dias_total !== undefined ? Math.round(p.dias_total) : -1;
  let pct;
  if(dias <= 0) pct = 0;
  else if(dias <= 3)  pct = (dias/3)*10;
  else if(dias <= 14) pct = 10 + ((dias-3)/11)*23;
  else if(dias <= 30) pct = 33 + ((dias-14)/16)*34;
  else pct = 100;
  pct = Math.min(98, Math.max(2, pct));
  const label = dias < 0 ? '0d' : dias === 0 ? '0d' : dias >= 30 ? '30d+' : dias+'d';
  const colorMark = dias <= 0 ? '#E74C3C' : dias <= 3 ? '#E67E22' : dias <= 14 ? '#E67E22' : '#27AE60';
  return '<div style="padding:0 2px">'
    + '<div style="height:22px;border-radius:6px;overflow:hidden;display:flex">'
    +   '<div style="width:10%;background:#E74C3C;display:flex;align-items:center;justify-content:center;font-size:9px;color:#fff;font-weight:700">0</div>'
    +   '<div style="width:23%;background:#E67E22;display:flex;align-items:center;justify-content:center;font-size:9px;color:#fff;font-weight:700">3d</div>'
    +   '<div style="width:34%;background:#F39C12;display:flex;align-items:center;justify-content:center;font-size:9px;color:#fff;font-weight:700">14d</div>'
    +   '<div style="flex:1;background:#27AE60;display:flex;align-items:center;justify-content:center;font-size:9px;color:#fff;font-weight:700">30d+</div>'
    + '</div>'
    + '<div style="position:relative;height:20px;margin-top:1px">'
    +   '<div style="position:absolute;left:'+pct+'%;top:0;transform:translateX(-50%);text-align:center">'
    +     '<div style="width:0;height:0;border-left:5px solid transparent;border-right:5px solid transparent;border-bottom:8px solid '+colorMark+';margin:0 auto"></div>'
    +     '<div style="font-size:10px;font-weight:700;color:'+colorMark+';white-space:nowrap;line-height:1.2">'+label+'</div>'
    +   '</div>'
    + '</div>'
    + '</div>';
}

// ── Tabla movimientos ──────────────────────────────────────
function tipoBadge(tipo){
  const m={'Producción':'tipo-prod','Venta':'tipo-venta','Despacho':'tipo-despacho',
    'Consumo':'tipo-consumo','Despacho recibido':'tipo-desp-rec','Entrada':'tipo-prod'};
  return '<span class="tipo-badge '+(m[tipo]||'tipo-prod')+'">'+tipo+'</span>';
}
function buildMovs(movs){
  if(!movs||movs.length===0) return '<p style="color:#bbb;font-size:12px;padding:8px 0">Sin movimientos.</p>';
  const rows = movs.map(function(m){
    const tienda = m.tienda==='Vitacura'
      ? '<span class="tienda-vit">Vitacura</span>'
      : '<span class="tienda-pat">Pataguas</span>';
    const cantCls = m.signo==='+' ? 'cant-pos' : 'cant-neg';
    return '<tr><td style="color:#888">'+m.fecha+'</td><td>'+tipoBadge(m.tipo)+'</td>'
      +'<td style="color:#666;font-size:11px">'+m.documento+'</td><td>'+tienda+'</td>'
      +'<td class="'+cantCls+'" style="text-align:right">'+m.signo+m.cantidad+'</td>'
      +'<td style="text-align:right;font-weight:600">'+m.stock+'</td></tr>';
  }).join('');
  return '<table class="movs-table"><thead><tr>'
    +'<th>Fecha</th><th>Tipo</th><th>Documento</th><th>Tienda</th>'
    +'<th style="text-align:right">Cant.</th><th style="text-align:right">Stock</th>'
    +'</tr></thead><tbody>'+rows+'</tbody></table>';
}

// ── Análisis ───────────────────────────────────────────────
function buildAnalisis(p){
  const vt = p.vel_total;
  const repo = p.tiempo_repo||7;
  let h = '';

  // Demanda real
  h += '<div class="insight"><b>📊 Demanda real del período</b><br>'
    + 'Vitacura: <b>'+p.total_vit+' und</b> en '+p.dias_stock_vit+' días con stock → <b>'+p.vel_vit.toFixed(3)+' und/día</b><br>'
    + 'Pataguas: <b>'+p.total_pat+' und</b> en '+p.dias_stock_pat+' días con stock → <b>'+p.vel_pat.toFixed(3)+' und/día</b><br>'
    + 'Total: <b>'+vt.toFixed(3)+' und/día</b> ('+(vt*30).toFixed(1)+' und/mes estimado)</div>';

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
      + 'Lote sugerido: <b>'+p.lote_sugerido+' und</b><br>'
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
  return '<div style="display:grid;grid-template-columns:1fr 1fr;border-top:0.5px solid #eee">'
    + '<div style="padding:10px 14px;border-right:0.5px solid #eee">'
    +   '<div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:0.04em;margin-bottom:4px">Vitacura</div>'
    +   '<div style="font-size:20px;font-weight:500;color:'+cv+'">'+p.vit+' und</div>'
    +   '<div style="font-size:11px;margin-top:2px;color:'+cv+'">'+textDias(p.dias_vit,p.vit)+'</div>'
    + '</div>'
    + '<div style="padding:10px 14px">'
    +   '<div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:0.04em;margin-bottom:4px">Pataguas</div>'
    +   '<div style="font-size:20px;font-weight:500;color:'+cp+'">'+p.pat+' und</div>'
    +   '<div style="font-size:11px;margin-top:2px;color:'+cp+'">'+textDias(p.dias_pat,p.pat)+'</div>'
    + '</div>'
    + '</div>';
}
function renderCards(data){
  const cont  = document.getElementById('productos');
  const noRes = document.getElementById('no-res');
  if(!data.length){cont.innerHTML='';noRes.style.display='block';return;}
  noRes.style.display='none';
  cont.innerHTML = data.map(function(p,i){
    const dStr = diasStr(p);
    const bCls = badgeCls(p.estado);
    const alHtml = p.alerta_dist ? '<span class="badge badge-dist">⚠ Distribución</span>' : '';

    return '<div class="card '+p.estado+'">'
      +'<div class="card-top" onclick="toggleCard('+i+')">'
      +  '<div class="card-info">'
      +    '<div class="card-nombre">'+p.nombre+' <span style="font-size:10px;color:#ccc;font-weight:400">'+p.sku+'</span></div>'
      +    '<div class="card-meta">'+p.cocinero+' · Repo: '+p.tiempo_repo+'d · Reordenar en '+p.pto_reorden+' und</div>'
      +  '</div>'
      +  '<div class="card-badges">'+alHtml+'<span class="badge '+bCls+'">'+dStr+'</span>'
      +    '<span class="chevron" id="chev-'+i+'">▼</span></div>'
      +'</div>'
      +'<div class="tl-wrap">'+buildBarra(p)+'</div>'
      +buildBloques(p)
      +'<button class="btn-det" onclick="toggleCard('+i+')">▼ Ver movimientos y análisis</button>'
      +'<div class="detalle" id="det-'+i+'">'
      +  '<div class="tabs">'
      +    '<button class="tab active" data-tab="mov" data-idx="'+i+'" onclick="switchTabD(this)">Movimientos</button>'
      +    '<button class="tab" data-tab="analisis" data-idx="'+i+'" onclick="switchTabD(this)">Análisis</button>'
      +  '</div>'
      +  '<div class="tab-body active" id="tab-'+i+'-mov">'+buildMovs(p.movs)+'</div>'
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
function filtrar(){
  const coc = document.getElementById('f-cocinero').value;
  const est = document.getElementById('f-estado').value;
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
function mostrarProductos(){
  document.getElementById('vista-guias').style.display='none';
  document.getElementById('vista-productos').style.display='block';
  document.getElementById('nav-guias').classList.remove('nav-active');
  document.getElementById('nav-productos').classList.add('nav-active');
}
function mostrarGuias(){
  document.getElementById('vista-productos').style.display='none';
  document.getElementById('vista-guias').style.display='block';
  document.getElementById('nav-productos').classList.remove('nav-active');
  document.getElementById('nav-guias').classList.add('nav-active');
  renderGuiaProduccion();
  renderGuiaDespacho();
}

// ── Guías ──────────────────────────────────────────────────
var ORDEN_EST = {'sin_stock':0,'critico':1,'bajo':2,'ok':3};

function renderGuiaProduccion(){
  const coc  = document.getElementById('g-cocinero').value;
  const dias = parseInt(document.getElementById('g-dias').value)||7;
  document.getElementById('g-dias-label').textContent = dias+' días';

  var prods = DATA.filter(function(p){return coc ? p.cocinero===coc : true;});
  prods.sort(function(a,b){
    var oa=ORDEN_EST[a.estado], ob=ORDEN_EST[b.estado];
    if(oa!==ob) return oa-ob;
    return (a.dias_total!==null?a.dias_total:9999)-(b.dias_total!==null?b.dias_total:9999);
  });

  var filas = prods.map(function(p){
    var und   = p.vel_total>0 ? Math.ceil(p.vel_total*dias) : 0;
    var total = p.vit + p.pat;
    var dt    = p.dias_total!==null ? Math.round(p.dias_total)+'d' : '—';
    var clr   = p.estado==='sin_stock'||p.estado==='critico' ? '#E74C3C' : p.estado==='bajo' ? '#E67E22' : '#27AE60';
    var urg   = p.estado==='sin_stock'||p.estado==='critico';
    return '<tr'+(urg?' class="urgente"':'')+'>'
      +'<td style="font-weight:'+(urg?700:400)+'">'+p.nombre+'</td>'
      +'<td style="color:'+clr+';font-weight:700;width:55px">'+dt+'</td>'
      +'<td style="color:#888;font-size:11px">'+p.cocinero+'</td>'
      +'<td style="text-align:right;color:#888">'+p.vit+'</td>'
      +'<td style="text-align:right;color:#185FA5">'+p.pat+'</td>'
      +'<td style="text-align:right;font-weight:600">'+total+'</td>'
      +'<td style="text-align:right;font-weight:700;font-size:13px">'+und+'</td>'
      +'</tr>';
  }).join('');

  document.getElementById('tabla-produccion').innerHTML = filas||'<tr><td colspan="7" style="text-align:center;color:#aaa;padding:20px">Sin productos</td></tr>';
  document.getElementById('resumen-produccion').textContent = prods.length+' productos · para '+dias+' días';
}

function renderGuiaDespacho(){
  const dias = parseInt(document.getElementById('d-dias').value)||7;
  document.getElementById('d-dias-label').textContent = dias+' días';

  var prods = DATA.filter(function(p){return p.vel_pat>0;});
  prods.sort(function(a,b){
    var da = a.pat>0&&a.vel_pat>0 ? a.pat/a.vel_pat : 0;
    var db = b.pat>0&&b.vel_pat>0 ? b.pat/b.vel_pat : 0;
    return da-db;
  });

  var despachar = 0;
  var filas = prods.map(function(p){
    var nec  = Math.ceil(p.vel_pat*dias);
    var desp = Math.max(0, nec-p.pat);
    var dpt  = p.pat>0&&p.vel_pat>0 ? Math.round(p.pat/p.vel_pat)+'d' : 'sin stock';
    var urg  = desp>0 && (p.pat===0||(p.pat/p.vel_pat)<3);
    if(desp>0) despachar++;
    return '<tr'+(urg?' class="urgente"':'')+'>'
      +'<td style="font-weight:'+(urg?700:400)+'">'+p.nombre+'</td>'
      +'<td style="text-align:right;color:#185FA5">'+p.pat+'</td>'
      +'<td style="text-align:right;color:#888">'+dpt+'</td>'
      +'<td style="text-align:right;color:#888">'+nec+'</td>'
      +'<td style="text-align:right;font-weight:700;font-size:13px;color:'+(desp>0?'#E74C3C':'#27AE60')+'">'+( desp>0?'+'+desp:'OK')+'</td>'
      +'</tr>';
  }).join('');

  document.getElementById('tabla-despacho').innerHTML = filas||'<tr><td colspan="5" style="text-align:center;color:#aaa;padding:20px">Sin productos</td></tr>';
  document.getElementById('resumen-despacho').textContent = despachar+' productos a despachar · '+dias+' días de cobertura';
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
function abrirRollo(contenido){
  var win = window.open('','_blank','width=400,height=700');
  var html = '<!DOCTYPE html><html><head><title>La Cocina</title>'
    + '<style>body{font-family:Courier New,monospace;font-size:10px;margin:4mm;width:72mm;line-height:1.4}'
    + 'pre{white-space:pre-wrap;word-break:break-word}'
    + '@media print{@page{size:80mm auto;margin:2mm 3mm}body{margin:0}}'
    + '</style></head><body><pre>' + contenido + '</pre>'
    + '<script>window.onload=function(){window.print();window.close();}<\/script>'
    + '</body></html>';
  win.document.write(html);
  win.document.close();
}

function imprimirProduccion(){
  var coc  = document.getElementById('g-cocinero').value;
  var dias = parseInt(document.getElementById('g-dias').value)||7;
  var sep  = '================================';
  var sep2 = '--------------------------------';
  var tit  = coc ? 'COCINERO: '+coc : 'PRODUCCION GENERAL';
  var prods = DATA.filter(function(p){return coc ? p.cocinero===coc : true;});
  prods.sort(function(a,b){
    var oa=ORDEN_EST[a.estado],ob=ORDEN_EST[b.estado];
    if(oa!==ob) return oa-ob;
    return (a.dias_total!==null?a.dias_total:9999)-(b.dias_total!==null?b.dias_total:9999);
  });
  var lineas = [sep,'  LA COCINA - '+tit,'  FECHA_HOY_PLACEHOLDER','  Producir para '+dias+' dias',sep2,
    'DIAS PRODUCTO               PRODUCIR',sep2];
  prods.forEach(function(p){
    var dt  = p.dias_total!==null ? ('   '+Math.round(p.dias_total)).slice(-4) : '  --';
    var nom = (p.nombre+'                      ').substring(0,22);
    var und = ('       '+( p.vel_total>0?Math.ceil(p.vel_total*dias):0 )+'u').slice(-8);
    lineas.push(dt+' '+nom+und);
  });
  lineas.push(sep2,'TOTAL: '+prods.length+' productos',sep);
  abrirRollo(lineas.join('\\n'));
}

function imprimirDespacho(){
  var dias = parseInt(document.getElementById('d-dias').value)||7;
  var sep  = '================================';
  var sep2 = '--------------------------------';
  var prods = DATA.filter(function(p){
    return p.vel_pat>0 && Math.max(0,Math.ceil(p.vel_pat*dias)-p.pat)>0;
  });
  prods.sort(function(a,b){
    var da = a.pat>0&&a.vel_pat>0 ? a.pat/a.vel_pat : 0;
    var db = b.pat>0&&b.vel_pat>0 ? b.pat/b.vel_pat : 0;
    return da-db;
  });
  var lineas = [sep,'  LA COCINA - DESPACHO','  FECHA_HOY_PLACEHOLDER',
    '  Vitacura -> Pataguas ('+dias+'d)',sep2,'PRODUCTO               PAT  DESP',sep2];
  prods.forEach(function(p){
    var nom  = (p.nombre+'                      ').substring(0,22);
    var pat  = ('    '+p.pat).slice(-4);
    var desp = ('    +'+Math.max(0,Math.ceil(p.vel_pat*dias)-p.pat)).slice(-5);
    lineas.push(nom+pat+desp);
  });
  lineas.push(sep2,'TOTAL: '+prods.length+' productos',sep);
  abrirRollo(lineas.join('\\n'));
}

// ── Init ───────────────────────────────────────────────────
document.getElementById('f-cocinero').addEventListener('change', filtrar);
document.getElementById('f-estado').addEventListener('change', filtrar);
document.getElementById('f-buscar').addEventListener('input', filtrar);
renderCards(DATA);
updateMetricas();
"""

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>La Cocina — Dashboard</title>
<style>CSS_PLACEHOLDER</style>
</head>
<body>

<div class="header">
  <div class="logo">La <span>Cocina</span> — Producción</div>
  <div class="header-right">
    <button class="btn nav-active" id="nav-productos" onclick="mostrarProductos()">📦 Productos</button>
    <button class="btn" id="nav-guias" onclick="mostrarGuias()">📋 Guías</button>
    <span class="fecha">FECHA_HOY_PLACEHOLDER · 56 productos</span>
  </div>
</div>

<div class="metricas">
  <div class="metrica"><div class="metrica-label">Sin stock</div><div class="metrica-valor val-rojo" id="m1">—</div></div>
  <div class="metrica"><div class="metrica-label">Crítico ≤3d</div><div class="metrica-valor val-rojo" id="m2">—</div></div>
  <div class="metrica"><div class="metrica-label">Bajo ≤14d</div><div class="metrica-valor val-amarillo" id="m3">—</div></div>
  <div class="metrica"><div class="metrica-label">OK &gt;14d</div><div class="metrica-valor val-verde" id="m4">—</div></div>
</div>

<!-- VISTA PRODUCTOS -->
<div id="vista-productos">
  <div class="toolbar">
    <span class="toolbar-label">Cocinero:</span>
    <select id="f-cocinero"><option value="">Todos</option><option>CAROLINA</option><option>ADRIANA</option><option>CÉSAR</option><option>JESÚS</option></select>
    <div class="toolbar-sep"></div>
    <span class="toolbar-label">Estado:</span>
    <select id="f-estado"><option value="">Todos</option><option value="sin_stock">Sin stock</option><option value="critico">Crítico</option><option value="bajo">Bajo stock</option><option value="ok">OK</option></select>
    <div class="toolbar-sep"></div>
    <input type="text" id="f-buscar" placeholder="Buscar producto...">
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
  </div>
  <div style="padding:12px 20px">
    <table class="guia-table">
      <thead><tr>
        <th>Producto</th><th style="text-align:right">Stock PAT</th>
        <th style="text-align:right">Días PAT</th><th style="text-align:right">Necesita</th>
        <th style="text-align:right">Despachar</th>
      </tr></thead>
      <tbody id="tabla-despacho"></tbody>
    </table>
  </div>

</div>

<script>
const DATA = DATA_PLACEHOLDER;
JS_PLACEHOLDER
</script>
</body>
</html>"""

def generar_html(datos, fecha_str):
    data_json = json.dumps(datos, ensure_ascii=False)
    js_final  = JS.replace('FECHA_HOY_PLACEHOLDER', fecha_str)
    html = HTML_TEMPLATE
    html = html.replace('CSS_PLACEHOLDER',         CSS)
    html = html.replace('DATA_PLACEHOLDER',        data_json)
    html = html.replace('JS_PLACEHOLDER',          js_final)
    html = html.replace('FECHA_HOY_PLACEHOLDER',   fecha_str)
    return html

# ─── MAIN ────────────────────────────────────────────────────
if __name__ == '__main__':
    print('='*50)
    print('La Cocina — Generador de Dashboard')
    print('='*50)
    datos = procesar()
    print(f'\nGenerando HTML con {len(datos)} productos...')
    html = generar_html(datos, FECHA_STR)
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
