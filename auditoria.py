# Auditoría producto a producto del cálculo de velocidad y calidad de datos.
# Replica exactamente la lógica de generar_dashboard.py sobre historial.json
# y marca señales de alerta por producto. No modifica nada.
import json, statistics, sys
import pandas as pd
from collections import defaultdict

CARPETA = r'C:\Users\alamo\OneDrive\Escritorio\reporte-stock-PRUEBAS'
data = json.load(open(f'{CARPETA}\\historial.json', encoding='utf-8'))
FECHA_HOY  = pd.Timestamp.now().normalize()
MES_ACTUAL = pd.Period(FECHA_HOY, 'M')
F_HIST     = (MES_ACTUAL - 3).start_time
UMBRAL     = 7

NOMBRES = {}
for m in data['movimientos']:
    NOMBRES.setdefault(m['sku'], m['sku'])

def analizar_oficina(sku, oficina):
    movs = sorted([m for m in data['movimientos'] if m['sku'] == sku and m['oficina'] == oficina],
                  key=lambda m: m['fecha'])
    if not movs:
        return None
    # stock diario reconstruido + ventas "clampeadas" (venta con stock reconstruido en 0)
    byday = defaultdict(lambda: [0.0, 0.0])
    for m in movs:
        if m['tipo'] == 'produccion':
            byday[pd.Timestamp(m['fecha'])][0] += m['cantidad']
        else:
            byday[pd.Timestamp(m['fecha'])][1] += m['cantidad']
    dias = pd.date_range(pd.Timestamp(movs[0]['fecha']), FECHA_HOY)
    st, sd, clamped = 0.0, {}, 0.0
    for d in dias:
        if d in byday:
            ent, sal = byday[d]
            nuevo = st + ent - sal
            if nuevo < 0:
                clamped += -nuevo
                nuevo = 0
            st = nuevo
        sd[d] = st
    # ventana 3 meses completos
    dias_mes, ventas_mes = defaultdict(int), defaultdict(float)
    for d, v in sd.items():
        if d >= F_HIST and v > 0 and d.to_period('M') != MES_ACTUAL:
            dias_mes[d.to_period('M')] += 1
    for m in movs:
        if m['tipo'] == 'venta' and pd.Timestamp(m['fecha']) >= F_HIST:
            ventas_mes[pd.Timestamp(m['fecha']).to_period('M')] += m['cantidad']
    validos = {mes: d for mes, d in dias_mes.items() if d >= UMBRAL}
    tv = sum(ventas_mes.get(mes, 0) for mes in validos)
    td = sum(validos.values())
    vel = tv / td if td else 0.0
    # ventas en la ventana que NO se contaron (meses con quiebre o mes actual)
    ventas_no_contadas = sum(v for mes, v in ventas_mes.items()
                             if mes not in validos and mes != MES_ACTUAL)
    ventas_mes_actual = ventas_mes.get(MES_ACTUAL, 0)
    # tasas mensuales para detectar tendencia
    tasas = {str(mes): (ventas_mes.get(mes, 0) / d) for mes, d in sorted(validos.items())}
    # quiebres (toda la historia) para tiempo de reposición
    per, ini = [], None
    for d in sorted(sd):
        if sd[d] <= 0:
            if ini is None: ini = d
        else:
            if ini is not None:
                dur = (d - ini).days
                if dur > 0: per.append(dur)
                ini = None
    quiebre_abierto = (FECHA_HOY - ini).days if ini is not None else 0
    return {
        'vel': vel, 'tv': tv, 'td': td,
        'n_meses_validos': len(validos),
        'ventas_no_contadas': ventas_no_contadas,
        'ventas_mes_actual': ventas_mes_actual,
        'tasas': tasas,
        'clamped': clamped,
        'quiebres': per,
        'quiebre_abierto': quiebre_abierto,
        'total_ventas_hist': sum(m['cantidad'] for m in movs if m['tipo'] == 'venta'),
    }

skus = sorted(set(m['sku'] for m in data['movimientos']))
informe = []
for sku in skus:
    v = analizar_oficina(sku, 'VIT')
    p = analizar_oficina(sku, 'PAT')
    flags = []
    for nombre_of, r in (('VIT', v), ('PAT', p)):
        if r is None:
            continue
        if r['total_ventas_hist'] == 0:
            continue
        if r['vel'] == 0 and (r['ventas_no_contadas'] > 0 or r['ventas_mes_actual'] > 0):
            flags.append(f"{nombre_of}: VELOCIDAD 0 pese a ventas recientes "
                         f"({r['ventas_no_contadas']:.0f} en meses con quiebre, "
                         f"{r['ventas_mes_actual']:.0f} este mes) — producto invisible para las guías")
        elif r['n_meses_validos'] == 1 and r['vel'] > 0:
            flags.append(f"{nombre_of}: velocidad basada en UN solo mes válido "
                         f"({r['tv']:.0f} ventas / {r['td']} días) — confianza baja")
        if r['ventas_no_contadas'] >= 5 and r['vel'] > 0:
            flags.append(f"{nombre_of}: {r['ventas_no_contadas']:.0f} ventas en meses con quiebre "
                         f"quedaron fuera del promedio — posible subestimación")
        if r['clamped'] >= 3:
            flags.append(f"{nombre_of}: {r['clamped']:.0f} ventas ocurrieron con stock reconstruido en 0 "
                         f"— faltan producciones en el historial, días con stock subcontados")
        if r['quiebre_abierto'] >= 14:
            flags.append(f"{nombre_of}: lleva {r['quiebre_abierto']} días seguidos sin stock (quiebre abierto)")
        if len(r['tasas']) >= 2:
            vals = list(r['tasas'].values())
            if min(vals) > 0 and max(vals) / min(vals) >= 3:
                flags.append(f"{nombre_of}: demanda muy variable entre meses "
                             f"({ {k: round(x,2) for k,x in r['tasas'].items()} }) — promedio menos fiable")
    vel_v = v['vel'] if v else 0
    vel_p = p['vel'] if p else 0
    informe.append({
        'sku': sku,
        'vel_vit': round(vel_v, 3), 'vel_pat': round(vel_p, 3),
        'vel_total': round(vel_v + vel_p, 3),
        'meses_v': (v['n_meses_validos'] if v else 0, p['n_meses_validos'] if p else 0),
        'flags': flags,
    })

ok      = [i for i in informe if not i['flags']]
revisar = [i for i in informe if i['flags']]

out = []
out.append(f"AUDITORIA {FECHA_HOY.date()} — ventana {F_HIST.date()} a fin de {MES_ACTUAL-1}")
out.append(f"Productos analizados: {len(informe)} | OK: {len(ok)} | Con alertas: {len(revisar)}")
out.append("")
out.append("=== CON ALERTAS ===")
for i in sorted(revisar, key=lambda x: -len(x['flags'])):
    out.append(f"\n{i['sku']}  (vel VIT {i['vel_vit']} + PAT {i['vel_pat']} = {i['vel_total']}/dia, "
               f"meses validos VIT/PAT: {i['meses_v'][0]}/{i['meses_v'][1]})")
    for f in i['flags']:
        out.append(f"   - {f}")
out.append("")
out.append("=== OK (sin alertas) ===")
for i in ok:
    out.append(f"{i['sku']:8} vel total {i['vel_total']}/dia  (meses VIT/PAT: {i['meses_v'][0]}/{i['meses_v'][1]})")

with open(f'{CARPETA}\\auditoria_resultado.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(out))
print(f"OK -> auditoria_resultado.txt ({len(revisar)} con alertas, {len(ok)} ok)")
