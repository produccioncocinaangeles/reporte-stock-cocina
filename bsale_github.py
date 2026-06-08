import requests
import smtplib
import os
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# ============================================================
# LISTA FIJA DE PRODUCTOS DE PRODUCCION
# ============================================================

PRODUCTOS = [
    {"cm": "OP",     "nombre": "OSTION A LA PARMESANA",           "promedio": 10.0, "cocinero": "CAROLINA"},
    {"cm": "RPP",    "nombre": "ROLLO POLLO PIMENTON",            "promedio": 5.0,  "cocinero": "CESAR"},
    {"cm": "MCL",    "nombre": "MINI CHUPE LOCO",                 "promedio": 30.0, "cocinero": "JESUS"},
    {"cm": "CJ",     "nombre": "CHUPE JAIBA GRANDE",              "promedio": 15.0, "cocinero": "JESUS"},
    {"cm": "CLC",    "nombre": "CHUPE LOCO CAMARON GRANDE",       "promedio": 45.0, "cocinero": "JESUS"},
    {"cm": "MCJ",    "nombre": "MINI CHUPE JAIBA",                "promedio": 20.0, "cocinero": "JESUS"},
    {"cm": "RC",     "nombre": "ROLLO CAMARON",                   "promedio": 10.0, "cocinero": "CESAR"},
    {"cm": "RM",     "nombre": "ROLLO MECHADA",                   "promedio": 25.0, "cocinero": "CESAR"},
    {"cm": "LA",     "nombre": "LOMITO ACARAMELADO",              "promedio": 7.0,  "cocinero": "CESAR"},
    {"cm": "LCM",    "nombre": "LASAÑA CARNE MECHADA",            "promedio": 7.0,  "cocinero": "ADRIANA"},
    {"cm": "TLS",    "nombre": "TEQUEÑOS LOMO SALTADO",           "promedio": 20.0, "cocinero": "ADRIANA"},
    {"cm": "CL",     "nombre": "CARPACCIO DE LOCOS",              "promedio": 20.0, "cocinero": "JESUS"},
    {"cm": "LS",     "nombre": "LOMO SALTADO",                    "promedio": 18.0, "cocinero": "CESAR"},
    {"cm": "LJC",    "nombre": "LASAÑA JAIBA CAMARON",            "promedio": 5.0,  "cocinero": "ADRIANA"},
    {"cm": "PC",     "nombre": "PURE DE CAMOTE",                  "promedio": 12.0, "cocinero": "CAROLINA"},
    {"cm": "MMR",    "nombre": "MIX MASITAS RELLENAS",            "promedio": 11.0, "cocinero": "ADRIANA"},
    {"cm": "ÑJ",     "nombre": "ÑOQUIS CON JAMON SERRANO",        "promedio": 15.0, "cocinero": "CAROLINA"},
    {"cm": "TSA",    "nombre": "TALLARIN SALMON AHUMADO",         "promedio": 15.0, "cocinero": "ADRIANA"},
    {"cm": "TS",     "nombre": "TARTARO SALMON",                  "promedio": 20.0, "cocinero": "CESAR"},
    {"cm": "RB",     "nombre": "ROAST BEEF",                      "promedio": 30.0, "cocinero": "JESUS"},
    {"cm": "CF",     "nombre": "CARPACCIO DE FILETE",             "promedio": 15.1, "cocinero": "JESUS"},
    {"cm": "RS",     "nombre": "ROLLO SALMON",                    "promedio": 20.0, "cocinero": "CESAR"},
    {"cm": "FCT",    "nombre": "FILETE CHAMPIÑON TOCINO",         "promedio": 5.0,  "cocinero": "JESUS"},
    {"cm": "PCM",    "nombre": "PASTEL DE CHOCLO",                "promedio": 10.0, "cocinero": "CAROLINA"},
    {"cm": "ÑP",     "nombre": "ÑOQUIS PESTO TOMATE CHERRY",      "promedio": 9.9,  "cocinero": "CAROLINA"},
    {"cm": "TF",     "nombre": "TARTARO FILETE",                  "promedio": 5.0,  "cocinero": "CESAR"},
    {"cm": "MCC",    "nombre": "MINI CHUPE CAMARON",              "promedio": 12.0, "cocinero": "JESUS"},
    {"cm": "CAMA",   "nombre": "CAMARONES APANADOS",              "promedio": 35.0, "cocinero": "ADRIANA"},
    {"cm": "TPP",    "nombre": "TALLARINES POLLO PIMENTON",       "promedio": 7.0,  "cocinero": "ADRIANA"},
    {"cm": "RA",     "nombre": "ROLLO ALCACHOFA",                 "promedio": 6.0,  "cocinero": "CESAR"},
    {"cm": "PAC",    "nombre": "MILHOJAS DE PAPAS",               "promedio": 20.0, "cocinero": "CESAR"},
    {"cm": "EM",     "nombre": "EMPANADITAS MECHADA",             "promedio": 11.0, "cocinero": "CAROLINA"},
    {"cm": "ÑC",     "nombre": "ÑOQUIS DE CAMARON",               "promedio": 11.0, "cocinero": "CAROLINA"},
    {"cm": "EJ",     "nombre": "EMPANADITAS JAMON SERRANO",       "promedio": 11.0, "cocinero": "CAROLINA"},
    {"cm": "BAP",    "nombre": "BERENJENAS A LA PARMESANA",       "promedio": 10.0, "cocinero": "ADRIANA"},
    {"cm": "PV",     "nombre": "CAJA DE POSTRES EN VASITO",       "promedio": 15.0, "cocinero": "CAROLINA"},
    {"cm": "MIGNON", "nombre": "MIGNON DE POLLO",                 "promedio": 10.0, "cocinero": "CESAR"},
    {"cm": "TA",     "nombre": "TARTARO ATUN",                    "promedio": 20.0, "cocinero": "CESAR"},
    {"cm": "MMRM",   "nombre": "MIX MASITAS RELLENAS DEL MAR",    "promedio": 8.0,  "cocinero": "ADRIANA"},
    {"cm": "LR",     "nombre": "LOMO RELLENO",                    "promedio": 4.0,  "cocinero": "JESUS"},
    {"cm": "CP",     "nombre": "CARPACCIO PULPO CON SALSA",       "promedio": 10.0, "cocinero": "JESUS"},
    {"cm": "LSA",    "nombre": "LASAÑA SALMON",                   "promedio": 6.1,  "cocinero": "ADRIANA"},
    {"cm": "RJ",     "nombre": "ROLLO JAMON SERRANO",             "promedio": 5.0,  "cocinero": "CESAR"},
    {"cm": "AS",     "nombre": "ARROZ SALVAJE",                   "promedio": 2.0,  "cocinero": "PROVEEDOR"},
    {"cm": "TCM",    "nombre": "TALLARIN CARNE MECHADA",          "promedio": 5.0,  "cocinero": "ADRIANA"},
    {"cm": "CM",     "nombre": "CARNE MECHADA",                   "promedio": 10.0, "cocinero": "CESAR"},
    {"cm": "PCP",    "nombre": "PATE CON PERAS",                  "promedio": 45.0, "cocinero": "JESUS"},
    {"cm": "EC",     "nombre": "EMPANADITAS CAMARON",             "promedio": 11.0, "cocinero": "CAROLINA"},
    {"cm": "CC",     "nombre": "CHUPE CENTOLLA GRANDE",           "promedio": 5.0,  "cocinero": "JESUS"},
    {"cm": "SSA",    "nombre": "SALMON CON SALSA DE ALCAPARRAS",  "promedio": 2.0,  "cocinero": "JESUS"},
    {"cm": "LM",     "nombre": "LASAÑA MEDITERRANEA",             "promedio": 4.4,  "cocinero": "ADRIANA"},
    {"cm": "CAC",    "nombre": "CHOCLO A LA CREMA",               "promedio": 5.0,  "cocinero": "JESUS"},
    {"cm": "TPU",    "nombre": "TEQUEÑOS DE PULPO CON SALSA",     "promedio": 4.0,  "cocinero": "ADRIANA"},
    {"cm": "BAR",    "nombre": "BARQUILLOS",                      "promedio": 3.0,  "cocinero": "CAROLINA"},
    {"cm": "SA",     "nombre": "SALSA EN FRASCO",                 "promedio": 0.0,  "cocinero": "ADRIANA"},
    {"cm": "CAMAC",  "nombre": "CAMARONES COCIDOS CON SALSA",     "promedio": 0.0,  "cocinero": "CESAR"},
]

# ============================================================
# PASO 1: BAJAR STOCK DESDE BSALE
# ============================================================

def normalizar(texto):
    return " ".join(str(texto).split()).upper()

def obtener_stock_bsale(token):
    print("Bajando stock desde Bsale...")
    headers = {"access_token": token}
    stock = {}
    offset = 0
    limit = 50

    while True:
        url = f"https://api.bsale.cl/v1/stocks.json?limit={limit}&offset={offset}&expand=[variant,office]"
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            print(f"Error Bsale: {response.status_code}")
            break
        data = response.json()
        items = data.get("items", [])
        if not items:
            break
        for item in items:
            try:
                variant  = item.get("variant", {})
                office   = item.get("office", {})
                codigo   = normalizar(variant.get("code", ""))
                sucursal = normalizar(office.get("name", ""))
                cantidad = float(item.get("quantityAvailable", 0) or 0)
                if codigo:
                    key = (codigo, sucursal)
                    stock[key] = stock.get(key, 0) + cantidad
            except:
                continue
        offset += limit
        if offset >= data.get("count", 0):
            break

    print(f"Stock bajado: {len(stock)} registros")
    return stock

# ============================================================
# PASO 2: CALCULAR STOCK POR PRODUCTO
# ============================================================

def cargar_velocidades():
    ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "velocidades.json")
    if os.path.exists(ruta):
        with open(ruta, encoding="utf-8") as f:
            return json.load(f)
    return {}

def calcular_productos(stock_bsale):
    ORDEN = {"sin_stock": 0, "critico": 1, "bajo": 2, "ok": 3}
    velocidades = cargar_velocidades()
    resultado = []
    for p in PRODUCTOS:
        cm       = normalizar(p["cm"])
        vitacura = int(stock_bsale.get((cm, "VITACURA"), 0))
        pataguas = int(stock_bsale.get((cm, "LAS PATAGUAS"), 0))
        total    = vitacura + pataguas

        # Velocidad real del dashboard, fallback a promedio fijo
        vel_vit = velocidades.get(p["cm"], {}).get("vel_vit") or (p["promedio"] / 30 if p["promedio"] > 0 else 0)

        if vitacura == 0:
            dias_vit = 0
            estado   = "sin_stock"
        elif vel_vit > 0:
            dias_vit = round(vitacura / vel_vit, 1)
            if   dias_vit <= 3:  estado = "critico"
            elif dias_vit <= 14: estado = "bajo"
            else:                estado = "ok"
        else:
            dias_vit = None
            estado   = "ok"

        resultado.append({
            "nombre":   p["nombre"],
            "cocinero": p["cocinero"],
            "vitacura": vitacura,
            "pataguas": pataguas,
            "total":    total,
            "dias_vit": dias_vit,
            "estado":   estado,
        })

    resultado.sort(key=lambda x: (
        ORDEN[x["estado"]],
        x["dias_vit"] if x["dias_vit"] is not None else 9999
    ))
    return resultado

# ============================================================
# PASO 3: ARMAR CORREO HTML
# ============================================================

def armar_html(productos):
    fecha = datetime.now().strftime("%d/%m/%Y · %H:%M")

    sin_stock = [p for p in productos if p["estado"] == "sin_stock"]
    criticos  = [p for p in productos if p["estado"] == "critico"]
    bajos     = [p for p in productos if p["estado"] == "bajo"]
    ok        = [p for p in productos if p["estado"] == "ok"]

    BG    = {"sin_stock":"#FCEBEB","critico":"#FAEEDA","bajo":"#FAF3DA","ok":"#E1F5EE"}
    COLOR = {"sin_stock":"#A32D2D","critico":"#854F0B","bajo":"#7A6010","ok":"#0F6E56"}
    LABEL = {"sin_stock":"Sin stock","critico":"Crítico","bajo":"Bajo stock","ok":"OK"}

    def badge(estado):
        return (f'<span style="font-size:10px;font-weight:600;padding:2px 7px;border-radius:99px;'
                f'background:{BG[estado]};color:{COLOR[estado]}">{LABEL[estado]}</span>')

    def dias_str(p):
        if p["estado"] == "sin_stock": return "Sin stock"
        if p["dias_vit"] is None:      return "—"
        return f'{round(p["dias_vit"])}d'

    def fila(p):
        clr = COLOR[p["estado"]]
        return (f'<tr>'
                f'<td style="padding:7px 10px;font-size:12px;border-bottom:0.5px solid #eee">{badge(p["estado"])} {p["nombre"]}</td>'
                f'<td style="padding:7px 10px;font-size:12px;text-align:right;border-bottom:0.5px solid #eee">{p["vitacura"]}</td>'
                f'<td style="padding:7px 10px;font-size:12px;text-align:right;color:#185FA5;border-bottom:0.5px solid #eee">{p["pataguas"]}</td>'
                f'<td style="padding:7px 10px;font-size:12px;text-align:right;font-weight:600;color:{clr};border-bottom:0.5px solid #eee">{dias_str(p)}</td>'
                f'</tr>')

    def tabla(lista):
        if not lista: return '<p style="font-size:12px;color:#aaa;padding:8px 0">Sin productos</p>'
        filas = "".join(fila(p) for p in lista)
        return (f'<table style="width:100%;border-collapse:collapse">'
                f'<thead><tr style="background:#f5f4f0">'
                f'<th style="padding:6px 10px;font-size:10px;font-weight:600;color:#888;text-align:left;border-bottom:1px solid #eee;text-transform:uppercase">Producto</th>'
                f'<th style="padding:6px 10px;font-size:10px;font-weight:600;color:#888;text-align:right;border-bottom:1px solid #eee;text-transform:uppercase">Vitacura</th>'
                f'<th style="padding:6px 10px;font-size:10px;font-weight:600;color:#185FA5;text-align:right;border-bottom:1px solid #eee;text-transform:uppercase">Pataguas</th>'
                f'<th style="padding:6px 10px;font-size:10px;font-weight:600;color:#888;text-align:right;border-bottom:1px solid #eee;text-transform:uppercase">Días VIT</th>'
                f'</tr></thead><tbody>{filas}</tbody></table>')

    def seccion(titulo, lista, color="#1A1A1A"):
        return (f'<div style="margin-bottom:16px">'
                f'<p style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:{color};margin:0 0 6px;padding-bottom:4px;border-bottom:2px solid {color}">{titulo}</p>'
                f'{tabla(lista)}</div>')

    # Secciones por cocinero (ordenados por urgencia dentro de cada uno)
    COCINEROS = ["CAROLINA", "CESAR", "JESUS", "ADRIANA"]
    secciones_cocinero = ""
    for coc in COCINEROS:
        prods_coc = [p for p in productos if p["cocinero"] == coc]
        if prods_coc:
            urgentes = [p for p in prods_coc if p["estado"] in ("sin_stock","critico")]
            color_titulo = "#A32D2D" if urgentes else "#444"
            secciones_cocinero += seccion(coc, prods_coc, color_titulo)

    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:20px;background:#f5f4f0;font-family:Arial,sans-serif;">
<div style="max-width:640px;margin:0 auto;">

  <div style="background:#fff;border:0.5px solid #e5e5e2;border-radius:12px;padding:18px 24px;margin-bottom:10px;">
    <p style="font-size:17px;font-weight:700;margin:0 0 2px;color:#1A1A1A;">La Cocina — Reporte de Producción</p>
    <p style="font-size:12px;color:#888;margin:0">{fecha} · Prioridad por stock Vitacura</p>
  </div>

  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:10px;">
    <div style="background:#FCEBEB;border-radius:8px;padding:12px;text-align:center">
      <div style="font-size:26px;font-weight:700;color:#A32D2D">{len(sin_stock)}</div>
      <div style="font-size:10px;font-weight:600;color:#A32D2D;margin-top:2px">Sin stock</div>
    </div>
    <div style="background:#FAEEDA;border-radius:8px;padding:12px;text-align:center">
      <div style="font-size:26px;font-weight:700;color:#854F0B">{len(criticos)}</div>
      <div style="font-size:10px;font-weight:600;color:#854F0B;margin-top:2px">Crítico ≤3d</div>
    </div>
    <div style="background:#FAF3DA;border-radius:8px;padding:12px;text-align:center">
      <div style="font-size:26px;font-weight:700;color:#7A6010">{len(bajos)}</div>
      <div style="font-size:10px;font-weight:600;color:#7A6010;margin-top:2px">Bajo ≤14d</div>
    </div>
    <div style="background:#E1F5EE;border-radius:8px;padding:12px;text-align:center">
      <div style="font-size:26px;font-weight:700;color:#0F6E56">{len(ok)}</div>
      <div style="font-size:10px;font-weight:600;color:#0F6E56;margin-top:2px">OK</div>
    </div>
  </div>

  <div style="background:#fff;border:0.5px solid #e5e5e2;border-radius:12px;padding:16px 20px;margin-bottom:10px;">
    <p style="font-size:13px;font-weight:700;margin:0 0 14px;color:#1A1A1A">Por cocinero</p>
    {secciones_cocinero}
  </div>

  <div style="background:#fff;border:0.5px solid #e5e5e2;border-radius:12px;padding:16px 20px;">
    <p style="font-size:13px;font-weight:700;margin:0 0 10px;color:#1A1A1A">Lista completa · {len(productos)} productos</p>
    {tabla(productos)}
  </div>

  <p style="font-size:11px;color:#aaa;text-align:center;margin-top:10px">Generado automáticamente · Stock desde Bsale</p>
</div>
</body></html>"""
    return html

# ============================================================
# PASO 4: ENVIAR CORREO
# ============================================================

def enviar_correo(html):
    remitente     = os.environ["GMAIL_REMITENTE"]
    password      = os.environ["GMAIL_PASSWORD"]
    destinatarios = [c.strip() for c in os.environ["CORREOS_DESTINO"].split(",")]
    fecha         = datetime.now().strftime("%d/%m/%Y")

    msg = MIMEMultipart("alternative")
    msg["From"]    = remitente
    msg["To"]      = ", ".join(destinatarios)
    msg["Subject"] = f"Reporte de Produccion - {fecha}"
    msg.attach(MIMEText(html, "html", "utf-8"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(remitente, password)
        server.sendmail(remitente, destinatarios, msg.as_string())

    print(f"Correo enviado a: {', '.join(destinatarios)}")

# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 50)
    print("INICIANDO REPORTE DE STOCK")
    print("=" * 50)

    token       = os.environ["BSALE_TOKEN"]
    stock_bsale = obtener_stock_bsale(token)
    productos   = calcular_productos(stock_bsale)
    html        = armar_html(productos)
    enviar_correo(html)

    print("=" * 50)
    print("PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 50)

if __name__ == "__main__":
    main()
