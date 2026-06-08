"""
Cocina Ángeles Álamos — Dashboard de producción
Lee datos históricos de Google Sheets + stock actual de Bsale
Calcula velocidades reales excluyendo días sin stock
"""

import os
import json
import math
import requests
import time
from datetime import datetime, date, timedelta
from collections import Counter, defaultdict

import gspread
from google.oauth2.service_account import Credentials

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────
BSALE_TOKEN = os.environ["BSALE_TOKEN"]
BSALE_BASE  = "https://api.bsale.cl/v1"
HEADERS     = {"access_token": BSALE_TOKEN}
OFFICE_VIT  = 1
OFFICE_PAT  = 3
SHEET_URL   = "https://docs.google.com/spreadsheets/d/1bB0ze4WAHFSl5hWmPTIsXgnKjhFiy6zfMtPCIyn503Y/edit"
FECHA_INI   = date(2026, 1, 2)
FECHA_FIN   = date(2026, 6, 5)

# ─────────────────────────────────────────────────────────────
# PRODUCTOS
# ─────────────────────────────────────────────────────────────
PRODUCTOS = {
    'OP':{'nombre':'Ostión a la parmesana','cocinero':'Carolina'},
    'RPP':{'nombre':'Rollo pollo pimentón','cocinero':'César'},
    'MCL':{'nombre':'Mini chupe loco','cocinero':'Jesús'},
    'CJ':{'nombre':'Chupe jaiba grande','cocinero':'Jesús'},
    'CLC':{'nombre':'Chupe loco camarón grande','cocinero':'Jesús'},
    'MCJ':{'nombre':'Mini chupe jaiba','cocinero':'Jesús'},
    'RC':{'nombre':'Rollo camarón','cocinero':'César'},
    'LA':{'nombre':'Lomito acaramelado','cocinero':'César'},
    'RM':{'nombre':'Rollo mechada','cocinero':'César'},
    'LCM':{'nombre':'Lasaña carne mechada','cocinero':'Adriana'},
    'TLS':{'nombre':'Tequeños lomo saltado','cocinero':'Adriana'},
    'CL':{'nombre':'Carpaccio de locos','cocinero':'Jesús'},
    'LS':{'nombre':'Lomo saltado','cocinero':'César'},
    'LJC':{'nombre':'Lasaña jaiba camarón','cocinero':'Adriana'},
    'PC':{'nombre':'Puré de camote','cocinero':'Carolina'},
    'MMR':{'nombre':'Mix masitas rellenas','cocinero':'Adriana'},
    'TS':{'nombre':'Tártaro salmón','cocinero':'César'},
    'ÑJ':{'nombre':'Ñoquis con jamón serrano','cocinero':'Carolina'},
    'TSA':{'nombre':'Tallarín salmón ahumado','cocinero':'Adriana'},
    'RB':{'nombre':'Roast beef','cocinero':'Jesús'},
    'CF':{'nombre':'Carpaccio de filete','cocinero':'Jesús'},
    'RS':{'nombre':'Rollo salmón','cocinero':'César'},
    'FCT':{'nombre':'Filete champiñón tocino','cocinero':'Jesús'},
    'PCM':{'nombre':'Pastel de choclo','cocinero':'Carolina'},
    'ÑP':{'nombre':'Ñoquis pesto tomate cherry','cocinero':'Carolina'},
    'TF':{'nombre':'Tártaro filete','cocinero':'César'},
    'MCC':{'nombre':'Mini chupe camarón','cocinero':'Jesús'},
    'CAMA':{'nombre':'Camarones apanados','cocinero':'Adriana'},
    'TPP':{'nombre':'Tallarines pollo pimentón','cocinero':'Adriana'},
    'RA':{'nombre':'Rollo alcachofa','cocinero':'César'},
    'AS':{'nombre':'Arroz salvaje','cocinero':'Jesús'},
    'PAC':{'nombre':'Milhojas de papas','cocinero':'César'},
    'EM':{'nombre':'Empanaditas mechada','cocinero':'Carolina'},
    'ÑC':{'nombre':'Ñoquis de camarón','cocinero':'Carolina'},
    'EJ':{'nombre':'Empanaditas jamón serrano','cocinero':'Carolina'},
    'BAP':{'nombre':'Berenjenas a la parmesana','cocinero':'Adriana'},
    'MIGNON':{'nombre':'Mignon de pollo','cocinero':'César'},
    'PV':{'nombre':'Caja de postres en vasito','cocinero':'Carolina'},
    'TA':{'nombre':'Tártaro atún','cocinero':'César'},
    'MMRM':{'nombre':'Mix masitas rellenas del mar','cocinero':'Adriana'},
    'LR':{'nombre':'Lomo relleno','cocinero':'Jesús'},
    'CP':{'nombre':'Carpaccio pulpo con salsa al olivo','cocinero':'Jesús'},
    'LSA':{'nombre':'Lasaña salmón','cocinero':'Adriana'},
    'RJ':{'nombre':'Rollo jamón serrano','cocinero':'César'},
    'CM':{'nombre':'Carne mechada','cocinero':'César'},
    'TCM':{'nombre':'Tallarín carne mechada','cocinero':'Adriana'},
    'PCP':{'nombre':'Paté con peras','cocinero':'Jesús'},
    'EC':{'nombre':'Empanaditas camarón','cocinero':'Carolina'},
    'CC':{'nombre':'Chupe centolla grande','cocinero':'Jesús'},
    'SSA':{'nombre':'Salmón con salsa de alcaparras','cocinero':'Jesús'},
    'BAR':{'nombre':'Barquillos','cocinero':'Carolina'},
    'LM':{'nombre':'Lasaña mediterránea','cocinero':'Adriana'},
    'CAC':{'nombre':'Choclo a la crema','cocinero':'Jesús'},
    'TPU':{'nombre':'Tequeños de pulpo con salsa','cocinero':'Adriana'},
    'SA':{'nombre':'Salsa en frasco','cocinero':'Adriana'},
    'CAMAC':{'nombre':'Camarones cocidos con salsa','cocinero':'César'},
}

def norm(s):
    return ' '.join(str(s).split()).upper()

def all_dates(fi, ff):
    d = fi
    while d <= ff:
        yield d
        d += timedelta(days=1)

# ─────────────────────────────────────────────────────────────
# GOOGLE SHEETS
# ─────────────────────────────────────────────────────────────
def conectar_sheets():
    creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_url(SHEET_URL)

def leer_hoja(sh, nombre):
    ws = sh.worksheet(nombre)
    rows = ws.get_all_values()
    if len(rows) < 2:
        return []
    h = rows[0]
    return [dict(zip(h, r)) for r in rows[1:]]

# ─────────────────────────────────────────────────────────────
# BSALE
# ─────────────────────────────────────────────────────────────
def bsale_get(path, params=None):
    url = f"{BSALE_BASE}{path}"
    p = {"limit": 50, "offset": 0}
    if params:
        p.update(params)
    items = []
    while True:
        for i in range(3):
            try:
                r = requests.get(url, headers=HEADERS, params=p, timeout=30)
                r.raise_for_status()
                data = r.json()
                break
            except Exception as e:
                print(f"  ⚠ {e} (intento {i+1})")
                time.sleep(2)
        else:
            break
        items.extend(data.get("items", []))
        if p["offset"] + p["limit"] >= data.get("count", 0):
            break
        p["offset"] += p["limit"]
        time.sleep(0.15)
    return items

def obtener_stock():
    print("⏳ Descargando stock desde Bsale...")
    items = bsale_get("/stocks.json", {"expand": "[variant,office]"})
    stock = {}
    for item in items:
        try:
            sku = norm(item["variant"]["code"])
        except:
            continue
        if sku not in PRODUCTOS:
            continue
        try:
            office_id = int(item["office"]["id"])
        except:
            continue
        qty = int(item.get("quantityAvailable", 0) or 0)
        if sku not in stock:
            stock[sku] = {"vit": 0, "pat": 0}
        if office_id == OFFICE_VIT:
            stock[sku]["vit"] += qty
        elif office_id == OFFICE_PAT:
            stock[sku]["pat"] += qty
    print(f"   → {len(stock)} SKUs con stock")
    return stock

# ─────────────────────────────────────────────────────────────
# CÁLCULOS
# ─────────────────────────────────────────────────────────────
def calcular_analisis(ventas, mov_vit, mov_pat, stock_actual):
    print("🧮 Calculando velocidades reales...")

    # Ventas por SKU/sucursal/fecha
    v_vit = defaultdict(lambda: defaultdict(float))
    v_pat = defaultdict(lambda: defaultdict(float))
    for r in ventas:
        try:
            d = date.fromisoformat(r['Fecha'])
            sku = r['SKU']
            c = float(r['Cantidad'] or 0)
            if r['Sucursal'] == 'VITACURA':
                v_vit[sku][d] += c
            else:
                v_pat[sku][d] += c
        except:
            continue

    # Stock Vitacura por SKU/fecha
    stock_vit_h = defaultdict(dict)
    for r in mov_vit:
        try:
            d = date.fromisoformat(r['Fecha'])
            sku = r['SKU']
            s = float(r['Stock'] or 0)
            stock_vit_h[sku][d] = s
        except:
            continue

    # Entradas Pataguas por SKU/fecha
    ent_pat = defaultdict(lambda: defaultdict(float))
    for r in mov_pat:
        try:
            d = date.fromisoformat(r['Fecha'])
            sku = r['SKU']
            c = float(r['Cantidad'] or 0)
            ent_pat[sku][d] += c
        except:
            continue

    dias = list(all_dates(FECHA_INI, FECHA_FIN))
    analisis = {}

    for sku in PRODUCTOS:
        # VITACURA: usar stock del consolidado
        dias_con_vit = 0
        dias_sin_vit = 0
        periodos_sin = []
        en_zero = False
        inicio_zero = None

        for d in dias:
            s = stock_vit_h[sku].get(d, None)
            if s is not None:
                if s > 0:
                    dias_con_vit += 1
                    if en_zero:
                        dur = (d - inicio_zero).days
                        if dur > 0:
                            periodos_sin.append(dur)
                        en_zero = False
                else:
                    dias_sin_vit += 1
                    if not en_zero:
                        en_zero = True
                        inicio_zero = d

        total_vend_vit = sum(v_vit[sku].values())
        vel_vit = total_vend_vit / dias_con_vit if dias_con_vit > 0 else 0

        # PATAGUAS: reconstruir desde stock actual
        s_act_pat = stock_actual.get(sku, {}).get('pat', 0)
        stock_pat_d = {}
        s = s_act_pat
        for d in reversed(dias):
            stock_pat_d[d] = s
            s = max(0, s + v_pat[sku].get(d, 0) - ent_pat[sku].get(d, 0))

        dias_con_pat = sum(1 for d in dias if stock_pat_d[d] > 0)
        total_vend_pat = sum(v_pat[sku].values())
        vel_pat = total_vend_pat / dias_con_pat if dias_con_pat > 0 else 0

        # Tiempo reposición (moda)
        moda = Counter(periodos_sin).most_common(1)[0][0] if periodos_sin else 7

        # Stock actual y días restantes
        s_act_vit = stock_actual.get(sku, {}).get('vit', 0)
        vel_total = vel_vit + vel_pat

        dr_vit = round(s_act_vit / vel_vit, 1) if vel_vit > 0 and s_act_vit > 0 else None
        dr_pat = round(s_act_pat / vel_pat, 1) if vel_pat > 0 and s_act_pat > 0 else None
        candidatos = [d for d in [dr_vit, dr_pat] if d is not None]
        dias_rest = min(candidatos) if candidatos else None

        despacho = max(0, math.ceil(vel_pat * 7) - s_act_pat) if vel_pat > 0 else 0
        lote = math.ceil(vel_total * (moda + 7)) if vel_total > 0 else None

        if s_act_vit == 0 and s_act_pat == 0:
            estado = 'sin_stock'
        elif dias_rest is not None and dias_rest <= 3:
            estado = 'critico'
        elif dias_rest is not None and dias_rest <= 14:
            estado = 'bajo'
        else:
            estado = 'ok'

        analisis[sku] = {
            'vel_vit': round(vel_vit, 4), 'vel_pat': round(vel_pat, 4),
            'vel_total': round(vel_total, 4),
            'dias_con_vit': dias_con_vit, 'dias_sin_vit': dias_sin_vit,
            'dias_con_pat': dias_con_pat,
            'total_vend_vit': int(total_vend_vit), 'total_vend_pat': int(total_vend_pat),
            'tiempo_repo': moda, 'periodos_sin': periodos_sin,
            'dias_rest': dias_rest, 'dias_rest_vit': dr_vit, 'dias_rest_pat': dr_pat,
            'despacho': despacho, 'lote': lote, 'estado': estado,
        }

    print(f"   → {len(analisis)} productos calculados")
    return analisis

# ─────────────────────────────────────────────────────────────
# HTML
# ─────────────────────────────────────────────────────────────
def generar_html(filas, fecha_str):
    data_js = json.dumps(filas, ensure_ascii=False)
    return f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Producción — Cocina Ángeles Álamos</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f3;color:#1a1a1a}}
.header{{background:#fff;border-bottom:1px solid #e5e5e5;padding:16px 24px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
.hl{{display:flex;align-items:center;gap:10px}}
.dot{{width:8px;height:8px;border-radius:50%;background:#639922;flex-shrink:0}}
.hs{{display:flex;flex-direction:column;gap:1px}}
.hsub{{font-size:10px;color:#888;letter-spacing:.06em;text-transform:uppercase}}
.htit{{font-size:17px;font-weight:500;color:#1a1a1a}}
.hf{{font-size:12px;color:#888}}
.wrap{{max-width:900px;margin:0 auto;padding:0 24px}}
.resumen{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;padding:16px 0}}
.stat{{background:#fff;border:1px solid #e5e5e5;border-radius:10px;padding:12px 16px;text-align:center}}
.sn{{font-size:22px;font-weight:500}}.sl{{font-size:11px;color:#888;margin-top:2px}}
.sr{{color:#A32D2D}}.sa{{color:#854F0B}}.sb{{color:#185FA5}}.sv{{color:#3B6D11}}
.filtros{{display:flex;gap:8px;flex-wrap:wrap;padding-bottom:12px}}
.filtros select,.filtros input{{font-size:12px;padding:6px 10px;border-radius:8px;border:1px solid #e0e0e0;background:#fff;color:#1a1a1a;outline:none}}
.filtros input{{flex:1;min-width:160px}}
.cards{{display:flex;flex-direction:column;gap:8px;padding-bottom:24px}}
.card{{background:#fff;border:1px solid #e5e5e5;border-radius:12px;overflow:hidden;transition:box-shadow .15s}}
.card:hover{{box-shadow:0 2px 8px rgba(0,0,0,.06)}}
.ch{{display:flex;align-items:center;gap:10px;padding:12px 14px;cursor:pointer;user-select:none}}
.badge{{font-size:10px;font-weight:500;padding:3px 8px;border-radius:999px;white-space:nowrap;flex-shrink:0}}
.br{{background:#FCEBEB;color:#A32D2D}}.ba{{background:#FAEEDA;color:#854F0B}}
.bb{{background:#E6F1FB;color:#185FA5}}.bv{{background:#EAF3DE;color:#3B6D11}}
.cn{{font-size:13px;font-weight:500;flex:1}}
.cc{{font-size:11px;color:#888}}
.cd{{font-size:12px;font-weight:500;white-space:nowrap;margin-left:8px}}
.chev{{font-size:12px;color:#ccc;transition:transform .2s;flex-shrink:0}}
.chev.open{{transform:rotate(180deg)}}
.tw{{padding:4px 14px 10px;display:flex;align-items:center;gap:8px}}
.tl{{display:flex;justify-content:space-between;font-size:9px;margin-bottom:3px}}
.tt{{position:relative;height:7px;border-radius:4px;flex:1}}
.tbg{{position:absolute;inset:0;border-radius:4px;background:linear-gradient(to right,#E24B4A 0%,#E24B4A 10%,#EF9F27 10%,#EF9F27 30%,#378ADD 30%,#378ADD 70%,#639922 70%,#639922 100%)}}
.tm{{position:absolute;top:50%;width:11px;height:11px;border-radius:50%;background:#fff;border:2px solid #333;transform:translate(-50%,-50%);z-index:2}}
.tn{{display:flex;justify-content:space-between;font-size:9px;color:#aaa;margin-top:3px}}
.sr2{{display:grid;grid-template-columns:repeat(4,1fr);border-top:1px solid #f0f0f0}}
.sc{{padding:8px 12px;text-align:center;border-right:1px solid #f0f0f0}}
.sc:last-child{{border-right:none}}
.sc .v{{font-size:13px;font-weight:500}}.sc .l{{font-size:10px;color:#aaa;margin-top:1px}}
.det{{display:none;border-top:1px solid #f0f0f0}}
.det.open{{display:block}}
.tc{{padding:12px 14px}}
.ib{{background:#f9f9f9;border-radius:8px;padding:10px 12px;font-size:11px;color:#666;line-height:1.8;margin-bottom:8px}}
.ib b{{color:#1a1a1a;font-weight:500}}
.al{{font-size:11px;color:#854F0B;background:#FAEEDA;padding:6px 10px;border-radius:6px;margin-bottom:8px}}
.gbtn{{padding-bottom:24px}}
.btn{{padding:9px 20px;border-radius:8px;border:1px solid #3B6D11;background:#3B6D11;color:#fff;font-size:12px;cursor:pointer;font-weight:500}}
.btn:hover{{background:#2d5a0d}}
</style></head><body>
<div class="header">
  <div class="hl"><div class="dot"></div>
    <div class="hs">
      <div class="hsub">Sistema de producción · Vitacura &amp; Las Pataguas</div>
      <div class="htit">Cocina Ángeles Álamos</div>
    </div>
  </div>
  <div class="hf">Actualizado: {fecha_str}</div>
</div>
<div class="wrap">
<div class="resumen">
  <div class="stat"><div class="sn sr" id="css">-</div><div class="sl">Sin stock</div></div>
  <div class="stat"><div class="sn sa" id="ccr">-</div><div class="sl">Crítico ≤3d</div></div>
  <div class="stat"><div class="sn sb" id="cbj">-</div><div class="sl">Bajo ≤14d</div></div>
  <div class="stat"><div class="sn sv" id="cok">-</div><div class="sl">OK</div></div>
</div>
<div class="filtros">
  <select id="fe" onchange="render()">
    <option value="">Todos los estados</option>
    <option value="sin_stock">Sin stock</option>
    <option value="critico">Crítico</option>
    <option value="bajo">Bajo stock</option>
    <option value="ok">OK</option>
  </select>
  <select id="fc" onchange="render()">
    <option value="">Todos los cocineros</option>
    <option>Carolina</option><option>Adriana</option><option>César</option><option>Jesús</option>
  </select>
  <input id="fb" placeholder="Buscar producto..." oninput="render()">
</div>
<div class="cards" id="cards"></div>
<div class="gbtn"><button class="btn" onclick="imprimirGuias()">Imprimir guías</button></div>
</div>
<script>
const DATA={data_js};
const EC={{
  sin_stock:{{label:'Sin stock',cls:'br',color:'#A32D2D',orden:0}},
  critico:{{label:'Crítico',cls:'ba',color:'#854F0B',orden:1}},
  bajo:{{label:'Bajo',cls:'bb',color:'#185FA5',orden:2}},
  ok:{{label:'OK',cls:'bv',color:'#3B6D11',orden:3}},
}};
function pct(d){{
  if(!d||d<=0)return 0;if(d>=30)return 100;
  if(d<=3)return Math.round(d/3*20);
  if(d<=14)return Math.round(20+(d-3)/11*40);
  return Math.round(60+(d-14)/16*40);
}}
function fd(d,e){{
  if(e==='sin_stock')return'Sin stock';
  if(!d&&d!==0)return'—';
  return Math.round(d)+'d restantes';
}}
function buildCard(p){{
  const ec=EC[p.estado],pt=pct(p.dias_rest),mc=ec.color;
  const dv=p.vit===0?'sin stock':(p.dias_rest_vit?Math.round(p.dias_rest_vit)+'d':'—');
  const dp=p.pat===0?'sin stock':(p.dias_rest_pat?Math.round(p.dias_rest_pat)+'d':'—');
  const al=(p.vit===0&&p.pat>0)||(p.pat===0&&p.vit>0)?`<div class="al">⚠ Revisar distribución — una sucursal sin stock</div>`:'';
  const ds=p.despacho>0?`<span style="color:#3B6D11;font-weight:500">+${{p.despacho}} und sugeridas</span>`:`<span style="color:#aaa">No necesita</span>`;
  return`<div class="card" data-estado="${{p.estado}}" data-cocinero="${{p.cocinero.toLowerCase()}}" data-nombre="${{p.nombre.toLowerCase()}}">
    <div class="ch" onclick="toggle(this)">
      <span class="badge ${{ec.cls}}">${{ec.label}}</span>
      <span class="cn">${{p.nombre}}</span>
      <span class="cc">${{p.cocinero}}</span>
      <span class="cd" style="color:${{mc}}">${{fd(p.dias_rest,p.estado)}}</span>
      <span class="chev">▼</span>
    </div>
    <div class="tw">
      <span style="font-size:9px;color:#A32D2D">0d</span>
      <div style="flex:1">
        <div class="tl">
          <span style="color:#A32D2D">Sin stock</span><span style="color:#854F0B">Crítico</span>
          <span style="color:#185FA5">Bajo stock</span><span style="color:#3B6D11">OK</span>
        </div>
        <div class="tt"><div class="tbg"></div><div class="tm" style="left:${{pt}}%;border-color:${{mc}}"></div></div>
        <div class="tn"><span>0</span><span>3d</span><span>14d</span><span>30d+</span></div>
      </div>
      <span style="font-size:9px;color:#3B6D11">30d+</span>
    </div>
    <div class="sr2">
      <div class="sc"><div class="v">${{p.vit}}</div><div class="l">Vitacura (${{dv}})</div></div>
      <div class="sc"><div class="v">${{p.pat}}</div><div class="l">Pataguas (${{dp}})</div></div>
      <div class="sc"><div class="v">${{p.vit+p.pat}}</div><div class="l">Total</div></div>
      <div class="sc"><div class="v">${{p.despacho>0?'+'+p.despacho:'OK'}}</div><div class="l">Despacho Pat.</div></div>
    </div>
    <div class="det">
      <div class="tc">
        ${{al}}
        <div class="ib">
          <b>Velocidad real de venta:</b><br>
          Vitacura: ${{p.vel_vit.toFixed(3)}} und/día (${{p.total_vend_vit}} und en ${{p.dias_con_vit}} días con stock)<br>
          Pataguas: ${{p.vel_pat.toFixed(3)}} und/día (${{p.total_vend_pat}} und en ${{p.dias_con_pat}} días con stock)<br>
          Total: ${{p.vel_total.toFixed(3)}} und/día
        </div>
        <div class="ib">
          <b>Días sin stock en el período:</b> ${{p.dias_sin_vit}} días en Vitacura<br>
          <b>Tiempo de reposición estimado:</b> ${{p.tiempo_repo}} día${{p.tiempo_repo!==1?'s':''}} (moda de períodos sin stock)<br>
          <b>Períodos sin stock detectados:</b> ${{p.periodos_sin.length>0?p.periodos_sin.join(', ')+'d':'Ninguno'}}
        </div>
        <div class="ib">
          <b>Despacho sugerido a Pataguas (7 días):</b> ${{ds}}<br>
          <b>Lote sugerido de producción:</b> ${{p.lote?p.lote+' und (reposición + 7 días colchón)':'—'}}
        </div>
      </div>
    </div>
  </div>`;
}}
function toggle(h){{
  const d=h.parentElement.querySelector('.det');
  const c=h.querySelector('.chev');
  d.classList.toggle('open');c.classList.toggle('open');
}}
function render(){{
  const est=document.getElementById('fe').value;
  const coc=document.getElementById('fc').value.toLowerCase();
  const bus=document.getElementById('fb').value.toLowerCase();
  const cnt={{sin_stock:0,critico:0,bajo:0,ok:0}};
  DATA.forEach(p=>cnt[p.estado]++);
  document.getElementById('css').textContent=cnt.sin_stock;
  document.getElementById('ccr').textContent=cnt.critico;
  document.getElementById('cbj').textContent=cnt.bajo;
  document.getElementById('cok').textContent=cnt.ok;
  document.getElementById('cards').innerHTML=DATA
    .filter(p=>{{
      if(est&&p.estado!==est)return false;
      if(coc&&p.cocinero.toLowerCase()!==coc)return false;
      if(bus&&!p.nombre.toLowerCase().includes(bus))return false;
      return true;
    }}).map(buildCard).join('');
}}
function imprimirGuias(){{
  const hoy=new Date().toLocaleDateString('es-CL');
  const cocineros=['Carolina','Adriana','César','Jesús'];
  let html=`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Guías de producción</title>
  <style>body{{font-family:monospace;font-size:11px;margin:20px}}h1{{font-size:13px;font-weight:bold;margin-bottom:2px}}
  h2{{font-size:12px;font-weight:bold;margin:16px 0 4px;border-bottom:1px solid #000;padding-bottom:2px}}
  p.sub{{font-size:10px;color:#666;margin-bottom:10px}}
  table{{width:100%;border-collapse:collapse;margin-bottom:12px}}
  th{{text-align:left;border-bottom:1px solid #000;padding:3px 6px;font-size:10px}}
  td{{padding:3px 6px;border-bottom:1px solid #eee}}
  .u{{font-weight:bold;color:#A32D2D}}.pb{{page-break-before:always}}
  .np{{display:none}}@media print{{.np{{display:none}}}}</style></head><body>
  <button class="np" onclick="window.print()" style="margin-bottom:16px;padding:6px 14px;cursor:pointer">🖨 Imprimir</button>`;
  html+=`<h1>Producción — Cocina Ángeles Álamos</h1><p class="sub">Generado: ${{hoy}} · Ordenado por urgencia</p>
  <h2>Guía general — todos los productos</h2>
  <table><thead><tr><th>Días</th><th>Producto</th><th>Cocinero</th><th>Vit</th><th>Pat</th><th>Total</th><th>Lote</th><th>Despacho Pat.</th></tr></thead><tbody>`;
  DATA.forEach(p=>{{
    const cls=p.dias_rest!==null&&p.dias_rest<=3?'u':'';
    const d=p.estado==='sin_stock'?'0d':p.dias_rest?Math.round(p.dias_rest)+'d':'—';
    html+=`<tr class="${{cls}}"><td>${{d}}</td><td>${{p.nombre}}</td><td>${{p.cocinero}}</td><td>${{p.vit}}</td><td>${{p.pat}}</td><td>${{p.vit+p.pat}}</td><td>${{p.lote||'—'}}</td><td>${{p.despacho>0?'+'+p.despacho:'—'}}</td></tr>`;
  }});
  html+=`</tbody></table>`;
  cocineros.forEach(coc=>{{
    const prods=DATA.filter(p=>p.cocinero===coc);
    html+=`<div class="pb"></div><h1>Producción — Cocina Ángeles Álamos</h1><p class="sub">Generado: ${{hoy}} · Cocinero: ${{coc}}</p>
    <h2>${{coc}} — ${{prods.length}} productos</h2>
    <table><thead><tr><th>Días</th><th>Producto</th><th>Vit</th><th>Pat</th><th>Total</th><th>Lote</th></tr></thead><tbody>`;
    prods.forEach(p=>{{
      const cls=p.dias_rest!==null&&p.dias_rest<=3?'u':'';
      const d=p.estado==='sin_stock'?'0d':p.dias_rest?Math.round(p.dias_rest)+'d':'—';
      html+=`<tr class="${{cls}}"><td>${{d}}</td><td>${{p.nombre}}</td><td>${{p.vit}}</td><td>${{p.pat}}</td><td>${{p.vit+p.pat}}</td><td>${{p.lote||'—'}}</td></tr>`;
    }});
    html+=`</tbody></table>`;
  }});
  const ds=DATA.filter(p=>p.despacho>0);
  html+=`<div class="pb"></div><h1>Producción — Cocina Ángeles Álamos</h1><p class="sub">Generado: ${{hoy}} · Despacho Vitacura → Las Pataguas (7 días)</p>
  <h2>Guía de despacho a Las Pataguas</h2>`;
  if(ds.length===0){{html+=`<p>Pataguas tiene stock suficiente para 7 días.</p>`;}}
  else{{
    html+=`<table><thead><tr><th>Producto</th><th>Stock Pat.</th><th>Vel. Pat.</th><th>Despacho sug.</th></tr></thead><tbody>`;
    ds.forEach(p=>{{html+=`<tr><td>${{p.nombre}}</td><td>${{p.pat}}</td><td>${{p.vel_pat.toFixed(2)}} und/día</td><td class="u">+${{p.despacho}}</td></tr>`;}});
    html+=`</tbody></table>`;
  }}
  html+=`</body></html>`;
  const w=window.open('','_blank');w.document.write(html);w.document.close();
}}
render();
</script></body></html>"""

# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
def main():
    print("🚀 Generando dashboard — Cocina Ángeles Álamos")

    print("\n📊 Leyendo datos históricos de Google Sheets...")
    sh = conectar_sheets()
    ventas  = leer_hoja(sh, 'VENTAS')
    mov_vit = leer_hoja(sh, 'MOV_VITACURA')
    mov_pat = leer_hoja(sh, 'MOV_PATAGUAS')
    print(f"   → Ventas: {len(ventas)} | Vit: {len(mov_vit)} | Pat: {len(mov_pat)}")

    stock_actual = obtener_stock()
    analisis = calcular_analisis(ventas, mov_vit, mov_pat, stock_actual)

    filas = []
    for sku, meta in PRODUCTOS.items():
        s = stock_actual.get(sku, {"vit": 0, "pat": 0})
        a = analisis.get(sku, {})
        filas.append({
            "sku": sku, "nombre": meta['nombre'], "cocinero": meta['cocinero'],
            "vit": s['vit'], "pat": s['pat'],
            "estado": a.get('estado','ok'),
            "dias_rest": a.get('dias_rest'),
            "dias_rest_vit": a.get('dias_rest_vit'),
            "dias_rest_pat": a.get('dias_rest_pat'),
            "vel_vit": a.get('vel_vit',0), "vel_pat": a.get('vel_pat',0),
            "vel_total": a.get('vel_total',0),
            "dias_con_vit": a.get('dias_con_vit',0),
            "dias_con_pat": a.get('dias_con_pat',0),
            "dias_sin_vit": a.get('dias_sin_vit',0),
            "total_vend_vit": a.get('total_vend_vit',0),
            "total_vend_pat": a.get('total_vend_pat',0),
            "tiempo_repo": a.get('tiempo_repo',0),
            "periodos_sin": a.get('periodos_sin',[]),
            "despacho": a.get('despacho',0),
            "lote": a.get('lote'),
        })

    orden = {"sin_stock":0,"critico":1,"bajo":2,"ok":3}
    filas.sort(key=lambda f: (orden[f["estado"]], f["dias_rest"] if f["dias_rest"] is not None else 9999))

    fecha_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    html = generar_html(filas, fecha_str)
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n✅ index.html generado — {len(filas)} productos · {fecha_str}")

if __name__ == "__main__":
    main()
