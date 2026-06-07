"""
Producción — Cocina Ángeles Álamos
Descarga stock actual desde Bsale API y genera index.html
Corre en GitHub Actions todos los días a las 00:00 UTC (21:00 Chile)
"""

import os
import math
import requests
import time
from datetime import datetime

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────
BSALE_TOKEN = os.environ["BSALE_TOKEN"]
BSALE_BASE  = "https://api.bsale.cl/v1"
HEADERS     = {"access_token": BSALE_TOKEN}
OFFICE_VIT  = 1
OFFICE_PAT  = 3

# ─────────────────────────────────────────────────────────────
# PRODUCTOS — SKU → {nombre, promedio_mes, cocinero}
# Promedios desde cocina_STOCK_NUEVA.xlsx (Jun 2026)
# promedio_mes=None → sin histórico suficiente
# ─────────────────────────────────────────────────────────────
PRODUCTOS = {
    'OP':     {'nombre': 'Ostión a la parmesana',              'prom': 10.0,  'cocinero': 'Carolina'},
    'RPP':    {'nombre': 'Rollo pollo pimentón',               'prom': 5.0,   'cocinero': 'César'},
    'MCL':    {'nombre': 'Mini chupe loco',                    'prom': 30.0,  'cocinero': 'Jesús'},
    'CJ':     {'nombre': 'Chupe jaiba grande',                 'prom': 15.0,  'cocinero': 'Jesús'},
    'CLC':    {'nombre': 'Chupe loco camarón grande',          'prom': 45.0,  'cocinero': 'Jesús'},
    'MCJ':    {'nombre': 'Mini chupe jaiba',                   'prom': 20.0,  'cocinero': 'Jesús'},
    'RC':     {'nombre': 'Rollo camarón',                      'prom': 10.0,  'cocinero': 'César'},
    'LA':     {'nombre': 'Lomito acaramelado',                 'prom': 7.0,   'cocinero': 'César'},
    'RM':     {'nombre': 'Rollo mechada',                      'prom': 25.0,  'cocinero': 'César'},
    'LCM':    {'nombre': 'Lasaña carne mechada',               'prom': 7.0,   'cocinero': 'Adriana'},
    'TLS':    {'nombre': 'Tequeños lomo saltado',              'prom': 20.0,  'cocinero': 'Adriana'},
    'CL':     {'nombre': 'Carpaccio de locos',                 'prom': 20.0,  'cocinero': 'Jesús'},
    'LS':     {'nombre': 'Lomo saltado',                       'prom': 18.0,  'cocinero': 'César'},
    'LJC':    {'nombre': 'Lasaña jaiba camarón',               'prom': 5.0,   'cocinero': 'Adriana'},
    'PC':     {'nombre': 'Puré de camote',                     'prom': 12.0,  'cocinero': 'Carolina'},
    'MMR':    {'nombre': 'Mix masitas rellenas',               'prom': 11.0,  'cocinero': 'Adriana'},
    'TS':     {'nombre': 'Tártaro salmón',                     'prom': 20.0,  'cocinero': 'César'},
    'ÑJ':     {'nombre': 'Ñoquis con jamón serrano',           'prom': 15.0,  'cocinero': 'Carolina'},
    'TSA':    {'nombre': 'Tallarín salmón ahumado',            'prom': 15.0,  'cocinero': 'Adriana'},
    'RB':     {'nombre': 'Roast beef',                         'prom': 30.0,  'cocinero': 'Jesús'},
    'CF':     {'nombre': 'Carpaccio de filete',                'prom': 15.1,  'cocinero': 'Jesús'},
    'RS':     {'nombre': 'Rollo salmón',                       'prom': 20.0,  'cocinero': 'César'},
    'FCT':    {'nombre': 'Filete champiñón tocino',            'prom': 5.0,   'cocinero': 'Jesús'},
    'PCM':    {'nombre': 'Pastel de choclo',                   'prom': 10.0,  'cocinero': 'Carolina'},
    'ÑP':     {'nombre': 'Ñoquis pesto tomate cherry',         'prom': 9.9,   'cocinero': 'Carolina'},
    'TF':     {'nombre': 'Tártaro filete',                     'prom': 5.0,   'cocinero': 'César'},
    'MCC':    {'nombre': 'Mini chupe camarón',                 'prom': 12.0,  'cocinero': 'Jesús'},
    'CAMA':   {'nombre': 'Camarones apanados',                 'prom': 35.0,  'cocinero': 'Adriana'},
    'TPP':    {'nombre': 'Tallarines pollo pimentón',          'prom': 7.0,   'cocinero': 'Adriana'},
    'RA':     {'nombre': 'Rollo alcachofa',                    'prom': 6.0,   'cocinero': 'César'},
    'AS':     {'nombre': 'Arroz salvaje',                      'prom': 2.0,   'cocinero': 'Jesús'},
    'PAC':    {'nombre': 'Milhojas de papas',                  'prom': 20.0,  'cocinero': 'César'},
    'EM':     {'nombre': 'Empanaditas mechada',                'prom': 11.0,  'cocinero': 'Carolina'},
    'ÑC':     {'nombre': 'Ñoquis de camarón',                  'prom': 11.0,  'cocinero': 'Carolina'},
    'EJ':     {'nombre': 'Empanaditas jamón serrano',          'prom': 11.0,  'cocinero': 'Carolina'},
    'BAP':    {'nombre': 'Berenjenas a la parmesana',          'prom': 10.0,  'cocinero': 'Adriana'},
    'MIGNON': {'nombre': 'Mignon de pollo',                    'prom': 10.0,  'cocinero': 'César'},
    'PV':     {'nombre': 'Caja de postres en vasito',          'prom': 15.0,  'cocinero': 'Carolina'},
    'TA':     {'nombre': 'Tártaro atún',                       'prom': 20.0,  'cocinero': 'César'},
    'MMRM':   {'nombre': 'Mix masitas rellenas del mar',       'prom': 8.0,   'cocinero': 'Adriana'},
    'LR':     {'nombre': 'Lomo relleno',                       'prom': 4.0,   'cocinero': 'Jesús'},
    'CP':     {'nombre': 'Carpaccio pulpo con salsa al olivo', 'prom': 10.0,  'cocinero': 'Jesús'},
    'LSA':    {'nombre': 'Lasaña salmón',                      'prom': 6.1,   'cocinero': 'Adriana'},
    'RJ':     {'nombre': 'Rollo jamón serrano',                'prom': 5.0,   'cocinero': 'César'},
    'CM':     {'nombre': 'Carne mechada',                      'prom': 10.0,  'cocinero': 'César'},
    'TCM':    {'nombre': 'Tallarín carne mechada',             'prom': 5.0,   'cocinero': 'Adriana'},
    'PCP':    {'nombre': 'Paté con peras',                     'prom': 45.0,  'cocinero': 'Jesús'},
    'EC':     {'nombre': 'Empanaditas camarón',                'prom': 11.0,  'cocinero': 'Carolina'},
    'CC':     {'nombre': 'Chupe centolla grande',              'prom': 5.0,   'cocinero': 'Jesús'},
    'SSA':    {'nombre': 'Salmón con salsa de alcaparras',     'prom': 2.0,   'cocinero': 'Jesús'},
    'BAR':    {'nombre': 'Barquillos',                         'prom': None,  'cocinero': 'Carolina'},
    'LM':     {'nombre': 'Lasaña mediterránea',                'prom': 4.4,   'cocinero': 'Adriana'},
    'CAC':    {'nombre': 'Choclo a la crema',                  'prom': 5.0,   'cocinero': 'Jesús'},
    'TPU':    {'nombre': 'Tequeños de pulpo con salsa',        'prom': 4.0,   'cocinero': 'Adriana'},
    'SA':     {'nombre': 'Salsa en frasco',                    'prom': None,  'cocinero': 'Adriana'},
    'CAMAC':  {'nombre': 'Camarones cocidos con salsa',        'prom': None,  'cocinero': 'César'},
}

# ─────────────────────────────────────────────────────────────
# BSALE API
# ─────────────────────────────────────────────────────────────
def bsale_get(path, params=None):
    url = f"{BSALE_BASE}{path}"
    p = {"limit": 50, "offset": 0}
    if params:
        p.update(params)
    items = []
    while True:
        for intento in range(3):
            try:
                r = requests.get(url, headers=HEADERS, params=p, timeout=30)
                r.raise_for_status()
                data = r.json()
                break
            except Exception as e:
                print(f"  ⚠ Error {path}: {e} (intento {intento+1})")
                time.sleep(2)
        else:
            break
        items.extend(data.get("items", []))
        if p["offset"] + p["limit"] >= data.get("count", 0):
            break
        p["offset"] += p["limit"]
        time.sleep(0.15)
    return items

def norm(s):
    return " ".join(str(s).split()).upper()

def obtener_stock():
    """Descarga stock actual de Bsale para los 56 productos."""
    print("⏳ Descargando stock desde Bsale...")
    items = bsale_get("/stocks.json", {"expand": "[variant,office]"})
    stock = {}
    for item in items:
        try:
            sku = norm(item["variant"]["code"])
        except (KeyError, TypeError):
            continue
        if sku not in PRODUCTOS:
            continue
        try:
            office_id = int(item["office"]["id"])
        except (KeyError, TypeError, ValueError):
            continue
        cantidad = int(item.get("quantityAvailable", 0) or 0)
        if sku not in stock:
            stock[sku] = {"vit": 0, "pat": 0}
        if office_id == OFFICE_VIT:
            stock[sku]["vit"] += cantidad
        elif office_id == OFFICE_PAT:
            stock[sku]["pat"] += cantidad
    print(f"   → {len(stock)} SKUs con stock")
    return stock

# ─────────────────────────────────────────────────────────────
# CÁLCULOS
# ─────────────────────────────────────────────────────────────
def calcular_dias(stock_total, prom_mes):
    if not prom_mes or prom_mes <= 0 or stock_total <= 0:
        return None
    return round(stock_total / (prom_mes / 30), 1)

def calcular_despacho(stock_pat, prom_mes):
    """Unidades sugeridas para cubrir 7 días en Pataguas."""
    if not prom_mes or prom_mes <= 0:
        return 0
    necesita = math.ceil((prom_mes / 30) * 7)
    return max(0, necesita - stock_pat)

def clasificar(vit, pat, dias):
    if vit == 0 and pat == 0:
        return "sin_stock"
    if dias is None:
        return "ok"
    if dias <= 3:
        return "critico"
    if dias <= 14:
        return "bajo"
    return "ok"

def pct_termometro(dias):
    """Convierte días a porcentaje para el termómetro (0-100)."""
    if dias is None or dias <= 0:
        return 0
    if dias >= 30:
        return 100
    if dias <= 3:
        return round(dias / 3 * 20)
    if dias <= 14:
        return round(20 + (dias - 3) / 11 * 40)
    return round(60 + (dias - 14) / 16 * 40)

# ─────────────────────────────────────────────────────────────
# GENERAR HTML
# ─────────────────────────────────────────────────────────────
def generar_html(filas, fecha_str):
    conteo = {"sin_stock": 0, "critico": 0, "bajo": 0, "ok": 0}
    for f in filas:
        conteo[f["estado"]] += 1

    # Serializar datos para JS
    import json
    data_js = json.dumps(filas, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Producción — Cocina Ángeles Álamos</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f3;color:#1a1a1a;min-height:100vh}}
.header{{background:#fff;border-bottom:1px solid #e5e5e5;padding:16px 24px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}}
.header-left{{display:flex;align-items:center;gap:10px}}
.dot{{width:8px;height:8px;border-radius:50%;background:#639922;flex-shrink:0}}
.header-stack{{display:flex;flex-direction:column;gap:1px}}
.header-sub{{font-size:10px;color:#888;letter-spacing:0.06em;text-transform:uppercase}}
.header-title{{font-size:17px;font-weight:500;color:#1a1a1a}}
.header-fecha{{font-size:12px;color:#888}}
.resumen{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;padding:16px 24px}}
.stat{{background:#fff;border:1px solid #e5e5e5;border-radius:10px;padding:12px 16px;text-align:center}}
.stat-n{{font-size:22px;font-weight:500}}
.stat-l{{font-size:11px;color:#888;margin-top:2px}}
.s-rojo{{color:#A32D2D}}.s-amber{{color:#854F0B}}.s-azul{{color:#185FA5}}.s-verde{{color:#3B6D11}}
.filtros{{padding:0 24px 12px;display:flex;gap:8px;flex-wrap:wrap}}
.filtros select,.filtros input{{font-size:12px;padding:6px 10px;border-radius:8px;border:1px solid #e0e0e0;background:#fff;color:#1a1a1a;outline:none}}
.filtros select:focus,.filtros input:focus{{border-color:#639922}}
.filtros input{{flex:1;min-width:160px}}
.cards{{padding:0 24px 24px;display:flex;flex-direction:column;gap:8px}}
.card{{background:#fff;border:1px solid #e5e5e5;border-radius:12px;overflow:hidden;transition:box-shadow .15s}}
.card:hover{{box-shadow:0 2px 8px rgba(0,0,0,.06)}}
.card-header{{display:flex;align-items:center;gap:10px;padding:12px 14px;cursor:pointer;user-select:none}}
.badge{{font-size:10px;font-weight:500;padding:3px 8px;border-radius:999px;white-space:nowrap;flex-shrink:0}}
.b-rojo{{background:#FCEBEB;color:#A32D2D}}
.b-amber{{background:#FAEEDA;color:#854F0B}}
.b-azul{{background:#E6F1FB;color:#185FA5}}
.b-verde{{background:#EAF3DE;color:#3B6D11}}
.card-nombre{{font-size:13px;font-weight:500;flex:1}}
.card-cocinero{{font-size:11px;color:#888}}
.card-dias{{font-size:12px;font-weight:500;white-space:nowrap;margin-left:8px}}
.chevron{{font-size:12px;color:#ccc;transition:transform .2s;flex-shrink:0}}
.chevron.open{{transform:rotate(180deg)}}
.term-wrap{{padding:4px 14px 10px;display:flex;align-items:center;gap:8px}}
.term-labels{{display:flex;justify-content:space-between;font-size:9px;margin-bottom:3px}}
.term-track{{position:relative;height:7px;border-radius:4px;flex:1}}
.term-bg{{position:absolute;inset:0;border-radius:4px;background:linear-gradient(to right,#E24B4A 0%,#E24B4A 10%,#EF9F27 10%,#EF9F27 30%,#378ADD 30%,#378ADD 70%,#639922 70%,#639922 100%)}}
.term-marker{{position:absolute;top:50%;width:11px;height:11px;border-radius:50%;background:#fff;border:2px solid #333;transform:translate(-50%,-50%);z-index:2}}
.term-nums{{display:flex;justify-content:space-between;font-size:9px;color:#aaa;margin-top:3px}}
.stats-row{{display:grid;grid-template-columns:repeat(4,1fr);border-top:1px solid #f0f0f0}}
.sc{{padding:8px 12px;text-align:center;border-right:1px solid #f0f0f0}}
.sc:last-child{{border-right:none}}
.sc .v{{font-size:13px;font-weight:500}}
.sc .l{{font-size:10px;color:#aaa;margin-top:1px}}
.detalle{{display:none;border-top:1px solid #f0f0f0}}
.detalle.open{{display:block}}
.tabs{{display:flex;border-bottom:1px solid #f0f0f0}}
.tab{{font-size:11px;padding:8px 14px;cursor:pointer;color:#aaa;border-bottom:2px solid transparent}}
.tab.active{{color:#3B6D11;border-bottom-color:#3B6D11;font-weight:500}}
.tc{{padding:12px 14px;display:none}}
.tc.active{{display:block}}
.movs{{width:100%;font-size:11px;border-collapse:collapse}}
.movs th{{text-align:left;color:#aaa;padding:4px 6px;font-weight:500;border-bottom:1px solid #f0f0f0}}
.movs td{{padding:5px 6px;border-bottom:1px solid #f5f5f5;color:#333}}
.movs tr:last-child td{{border-bottom:none}}
.vit{{color:#3B6D11;font-weight:500}}.pat{{color:#185FA5;font-weight:500}}
.pos{{color:#3B6D11}}.neg{{color:#A32D2D}}
.ib{{background:#f9f9f9;border-radius:8px;padding:10px 12px;font-size:11px;color:#666;line-height:1.8;margin-bottom:8px}}
.ib b{{color:#1a1a1a;font-weight:500}}
.alerta{{font-size:11px;color:#854F0B;background:#FAEEDA;padding:6px 10px;border-radius:6px;margin-bottom:8px}}
.guias-btn{{margin:0 24px 24px;display:flex;gap:10px}}
.btn{{padding:9px 20px;border-radius:8px;border:1px solid #e0e0e0;background:#fff;font-size:12px;cursor:pointer;color:#1a1a1a;font-weight:500}}
.btn:hover{{background:#f5f5f3;border-color:#ccc}}
.btn-primary{{background:#3B6D11;color:#fff;border-color:#3B6D11}}
.btn-primary:hover{{background:#2d5a0d}}
@media print{{body{{background:#fff}}.header{{position:static}}.filtros,.guias-btn,.chevron,.tabs{{display:none!important}}.detalle{{display:block!important}}.card{{break-inside:avoid;box-shadow:none;border:1px solid #ddd}}}}
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <div class="dot"></div>
    <div class="header-stack">
      <div class="header-sub">Sistema de producción · Vitacura &amp; Las Pataguas</div>
      <div class="header-title">Cocina Ángeles Álamos</div>
    </div>
  </div>
  <div class="header-fecha">Actualizado: {fecha_str}</div>
</div>

<div class="resumen">
  <div class="stat"><div class="stat-n s-rojo" id="cnt-sinstock">-</div><div class="stat-l">Sin stock</div></div>
  <div class="stat"><div class="stat-n s-amber" id="cnt-critico">-</div><div class="stat-l">Crítico ≤3d</div></div>
  <div class="stat"><div class="stat-n s-azul" id="cnt-bajo">-</div><div class="stat-l">Bajo ≤14d</div></div>
  <div class="stat"><div class="stat-n s-verde" id="cnt-ok">-</div><div class="stat-l">OK</div></div>
</div>

<div class="filtros">
  <select id="f-estado" onchange="render()">
    <option value="">Todos los estados</option>
    <option value="sin_stock">Sin stock</option>
    <option value="critico">Crítico</option>
    <option value="bajo">Bajo stock</option>
    <option value="ok">OK</option>
  </select>
  <select id="f-cocinero" onchange="render()">
    <option value="">Todos los cocineros</option>
    <option>Carolina</option>
    <option>Adriana</option>
    <option>César</option>
    <option>Jesús</option>
  </select>
  <input id="f-buscar" placeholder="Buscar producto..." oninput="render()">
</div>

<div class="cards" id="cards"></div>

<div class="guias-btn">
  <button class="btn btn-primary" onclick="imprimirGuias()">Imprimir guías</button>
</div>

<script>
const DATA = {data_js};

const EC = {{
  sin_stock:{{label:'Sin stock',cls:'b-rojo',color:'#A32D2D',orden:0}},
  critico:  {{label:'Crítico',  cls:'b-amber',color:'#854F0B',orden:1}},
  bajo:     {{label:'Bajo',     cls:'b-azul', color:'#185FA5',orden:2}},
  ok:       {{label:'OK',       cls:'b-verde',color:'#3B6D11',orden:3}},
}};

function pct(d){{
  if(!d||d<=0) return 0;
  if(d>=30) return 100;
  if(d<=3) return Math.round(d/3*20);
  if(d<=14) return Math.round(20+(d-3)/11*40);
  return Math.round(60+(d-14)/16*40);
}}

function diasStr(d,est){{
  if(est==='sin_stock') return 'Sin stock';
  if(!d) return '—';
  return d+'d restantes';
}}

function buildCard(p){{
  const ec = EC[p.estado];
  const mc = ec.color;
  const pt = pct(p.dias);
  const dvit = p.vit===0?'sin stock':(p.dias_vit?p.dias_vit+'d':'—');
  const dpat = p.pat===0?'sin stock':(p.dias_pat?p.dias_pat+'d':'—');
  const alerta = (p.vit===0&&p.pat>0)||(p.pat===0&&p.vit>0)
    ? `<div class="alerta">Revisar distribución — una sucursal sin stock</div>` : '';
  const desp = p.despacho>0
    ? `<span class="pos">+${{p.despacho}} und</span>`
    : `<span style="color:#aaa">No necesita</span>`;

  return `<div class="card" data-estado="${{p.estado}}" data-cocinero="${{p.cocinero.toLowerCase()}}" data-nombre="${{p.nombre.toLowerCase()}}">
    <div class="card-header" onclick="toggle(this)">
      <span class="badge ${{ec.cls}}">${{ec.label}}</span>
      <span class="card-nombre">${{p.nombre}}</span>
      <span class="card-cocinero">${{p.cocinero}}</span>
      <span class="card-dias" style="color:${{mc}}">${{diasStr(p.dias,p.estado)}}</span>
      <span class="chevron">▼</span>
    </div>
    <div class="term-wrap">
      <span style="font-size:9px;color:#A32D2D">0d</span>
      <div style="flex:1">
        <div class="term-labels">
          <span style="color:#A32D2D">Sin stock</span>
          <span style="color:#854F0B">Crítico</span>
          <span style="color:#185FA5">Bajo stock</span>
          <span style="color:#3B6D11">OK</span>
        </div>
        <div class="term-track">
          <div class="term-bg"></div>
          <div class="term-marker" style="left:${{pt}}%;border-color:${{mc}}"></div>
        </div>
        <div class="term-nums"><span>0</span><span>3d</span><span>14d</span><span>30d+</span></div>
      </div>
      <span style="font-size:9px;color:#3B6D11">30d+</span>
    </div>
    <div class="stats-row">
      <div class="sc"><div class="v">${{p.vit}}</div><div class="l">Vitacura</div></div>
      <div class="sc"><div class="v">${{p.pat}}</div><div class="l">Pataguas</div></div>
      <div class="sc"><div class="v">${{p.vit+p.pat}}</div><div class="l">Total</div></div>
      <div class="sc"><div class="v">${{p.despacho>0?'+'+p.despacho:'OK'}}</div><div class="l">Despacho Pat.</div></div>
    </div>
    <div class="detalle">
      <div class="tabs">
        <div class="tab active" onclick="switchTab(this,'m-${{p.sku}}','a-${{p.sku}}')">Movimientos</div>
        <div class="tab" onclick="switchTab(this,'a-${{p.sku}}','m-${{p.sku}}')">Análisis</div>
      </div>
      <div class="tc active" id="m-${{p.sku}}">
        <p style="font-size:11px;color:#aaa;text-align:center;padding:16px 0">Los movimientos detallados estarán disponibles cuando se habiliten los endpoints de Bsale.</p>
      </div>
      <div class="tc" id="a-${{p.sku}}">
        ${{alerta}}
        <div class="ib">
          <b>Stock actual:</b> Vitacura ${{p.vit}} und (${{dvit}}) · Pataguas ${{p.pat}} und (${{dpat}})<br>
          <b>Promedio mensual:</b> ${{p.prom?p.prom+' und/mes':'Sin datos históricos'}}<br>
          <b>Velocidad diaria:</b> ${{p.prom?Math.round(p.prom/30*100)/100+' und/día':'—'}}
        </div>
        <div class="ib">
          <b>Despacho sugerido a Pataguas (7 días):</b> ${{desp}}<br>
          <b>Lote sugerido de producción:</b> ${{p.lote?p.lote+' und':'—'}}
        </div>
      </div>
    </div>
  </div>`;
}}

function toggle(h){{
  const d=h.parentElement.querySelector('.detalle');
  const c=h.querySelector('.chevron');
  d.classList.toggle('open');
  c.classList.toggle('open');
}}

function switchTab(tab,showId,hideId){{
  const card=tab.closest('.card');
  card.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  tab.classList.add('active');
  document.getElementById(showId).classList.add('active');
  document.getElementById(hideId).classList.remove('active');
}}

function render(){{
  const est = document.getElementById('f-estado').value;
  const coc = document.getElementById('f-cocinero').value.toLowerCase();
  const bus = document.getElementById('f-buscar').value.toLowerCase();
  const cnt = {{sin_stock:0,critico:0,bajo:0,ok:0}};
  DATA.forEach(p=>cnt[p.estado]++);
  document.getElementById('cnt-sinstock').textContent = cnt.sin_stock;
  document.getElementById('cnt-critico').textContent  = cnt.critico;
  document.getElementById('cnt-bajo').textContent     = cnt.bajo;
  document.getElementById('cnt-ok').textContent       = cnt.ok;
  const filtrado = DATA
    .filter(p=>{{
      if(est && p.estado!==est) return false;
      if(coc && p.cocinero.toLowerCase()!==coc) return false;
      if(bus && !p.nombre.toLowerCase().includes(bus)) return false;
      return true;
    }});
  document.getElementById('cards').innerHTML = filtrado.map(buildCard).join('');
}}

function imprimirGuias(){{
  const hoy = new Date().toLocaleDateString('es-CL');
  const cocineros = ['Carolina','Adriana','César','Jesús'];

  let html = `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
  <title>Guías de producción</title>
  <style>
    body{{font-family:monospace;font-size:11px;margin:20px;color:#000}}
    h1{{font-size:14px;font-weight:bold;margin-bottom:4px}}
    h2{{font-size:12px;font-weight:bold;margin:16px 0 4px;border-bottom:1px solid #000;padding-bottom:2px}}
    table{{width:100%;border-collapse:collapse;margin-bottom:12px}}
    th{{text-align:left;border-bottom:1px solid #000;padding:3px 6px;font-size:10px}}
    td{{padding:3px 6px;border-bottom:1px solid #eee;font-size:11px}}
    .urgente{{font-weight:bold}}
    .page-break{{page-break-before:always}}
    @media print{{.no-print{{display:none}}}}
  </style></head><body>
  <button class="no-print" onclick="window.print()" style="margin-bottom:16px;padding:6px 14px;cursor:pointer">Imprimir</button>`;

  // GUÍA GENERAL
  html += `<h1>Producción — Cocina Ángeles Álamos</h1>
  <p style="font-size:10px;color:#666;margin-bottom:12px">Generado: ${{hoy}} · Ordenado por urgencia</p>
  <h2>Guía general — todos los productos</h2>
  <table><thead><tr><th>Días</th><th>Producto</th><th>Cocinero</th><th>Vit</th><th>Pat</th><th>Total</th><th>Lote sug.</th></tr></thead><tbody>`;

  DATA.forEach(p=>{{
    const cls = p.dias!==null&&p.dias<=3?'urgente':'';
    const diasTxt = p.estado==='sin_stock'?'0':p.dias?p.dias+'d':'—';
    html += `<tr class="${{cls}}"><td>${{diasTxt}}</td><td>${{p.nombre}}</td><td>${{p.cocinero}}</td><td>${{p.vit}}</td><td>${{p.pat}}</td><td>${{p.vit+p.pat}}</td><td>${{p.lote||'—'}}</td></tr>`;
  }});
  html += `</tbody></table>`;

  // GUÍAS POR COCINERO
  cocineros.forEach(coc=>{{
    const prods = DATA.filter(p=>p.cocinero===coc);
    html += `<div class="page-break"></div><h1>Producción — Cocina Ángeles Álamos</h1>
    <p style="font-size:10px;color:#666;margin-bottom:12px">Generado: ${{hoy}} · Cocinero: ${{coc}}</p>
    <h2>${{coc}} — ${{prods.length}} productos</h2>
    <table><thead><tr><th>Días</th><th>Producto</th><th>Vit</th><th>Pat</th><th>Total</th><th>Lote sug.</th></tr></thead><tbody>`;
    prods.forEach(p=>{{
      const cls = p.dias!==null&&p.dias<=3?'urgente':'';
      const diasTxt = p.estado==='sin_stock'?'0':p.dias?p.dias+'d':'—';
      html += `<tr class="${{cls}}"><td>${{diasTxt}}</td><td>${{p.nombre}}</td><td>${{p.vit}}</td><td>${{p.pat}}</td><td>${{p.vit+p.pat}}</td><td>${{p.lote||'—'}}</td></tr>`;
    }});
    html += `</tbody></table>`;
  }});

  // GUÍA DESPACHO PATAGUAS
  const despachos = DATA.filter(p=>p.despacho>0);
  html += `<div class="page-break"></div><h1>Producción — Cocina Ángeles Álamos</h1>
  <p style="font-size:10px;color:#666;margin-bottom:12px">Generado: ${{hoy}} · Despacho a Las Pataguas (7 días)</p>
  <h2>Guía de despacho — Vitacura → Las Pataguas</h2>`;
  if(despachos.length===0){{
    html += `<p style="color:#666">No hay despachos sugeridos — Pataguas tiene stock suficiente para 7 días.</p>`;
  }}else{{
    html += `<table><thead><tr><th>Producto</th><th>Stock Pat.</th><th>Despacho sug.</th></tr></thead><tbody>`;
    despachos.forEach(p=>{{
      html += `<tr><td>${{p.nombre}}</td><td>${{p.pat}}</td><td class="urgente">+${{p.despacho}}</td></tr>`;
    }});
    html += `</tbody></table>`;
  }}

  html += `</body></html>`;
  const w = window.open('','_blank');
  w.document.write(html);
  w.document.close();
}}

render();
</script>
</body>
</html>"""
    return html

# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
def main():
    print("🚀 Generando dashboard — Cocina Ángeles Álamos")

    stock = obtener_stock()
    fecha_str = datetime.now().strftime("%d/%m/%Y %H:%M")

    filas = []
    for sku, meta in PRODUCTOS.items():
        s    = stock.get(sku, {"vit": 0, "pat": 0})
        vit  = s["vit"]
        pat  = s["pat"]
        prom = meta["prom"]
        dias = calcular_dias(vit + pat, prom)
        dias_vit = calcular_dias(vit, prom * 0.6) if prom else None  # aprox proporción Vitacura
        dias_pat = calcular_dias(pat, prom * 0.4) if prom else None  # aprox proporción Pataguas
        desp = calcular_despacho(pat, prom)
        lote = math.ceil(prom / 30 * 14) if prom else None  # 14 días de cobertura
        est  = clasificar(vit, pat, dias)

        filas.append({
            "sku":      sku,
            "nombre":   meta["nombre"],
            "cocinero": meta["cocinero"],
            "estado":   est,
            "dias":     dias,
            "dias_vit": dias_vit,
            "dias_pat": dias_pat,
            "vit":      vit,
            "pat":      pat,
            "prom":     prom,
            "despacho": desp,
            "lote":     lote,
        })

    # Ordenar por urgencia: sin_stock → critico → bajo → ok, luego por días asc
    orden = {"sin_stock": 0, "critico": 1, "bajo": 2, "ok": 3}
    filas.sort(key=lambda f: (orden[f["estado"]], f["dias"] if f["dias"] is not None else 9999))

    html = generar_html(filas, fecha_str)

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ index.html generado — {len(filas)} productos · {fecha_str}")

if __name__ == "__main__":
    main()
