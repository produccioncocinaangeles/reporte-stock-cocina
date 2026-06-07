import requests
import smtplib
import os
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

def calcular_productos(stock_bsale):
    resultado = []
    for p in PRODUCTOS:
        cm = normalizar(p["cm"])
        vitacura = int(stock_bsale.get((cm, "VITACURA"), 0))
        pataguas = int(stock_bsale.get((cm, "LAS PATAGUAS"), 0))
        total    = vitacura + pataguas
        prom     = p["promedio"]

        if prom > 0:
            dias = round(total / (prom / 30), 1)
        else:
            dias = None

        resultado.append({
            "nombre":   p["nombre"],
            "cocinero": p["cocinero"],
            "vitacura": vitacura,
            "pataguas": pataguas,
            "total":    total,
            "dias":     dias,
        })

    resultado.sort(key=lambda x: (x["dias"] is None, x["dias"] if x["dias"] is not None else 9999))
    return resultado

# ============================================================
# PASO 3: ARMAR CORREO HTML
# ============================================================

def dias_label(dias):
    if dias is None: return "-"
    if dias == 0:    return "0 dias"
    if dias == 1:    return "1 dia"
    return f"{dias} dias"

def color_dias(dias):
    if dias is None:  return "#888780"
    if dias <= 1:     return "#A32D2D"
    if dias <= 3:     return "#854F0B"
    return "#3B6D11"

def armar_html(productos):
    fecha = datetime.now().strftime("%d de %B de %Y · %H:%M")

    sin_stock = [p for p in productos if p["total"] == 0]
    criticos  = [p for p in productos if p["total"] > 0 and p["dias"] is not None and p["dias"] <= 3]
    bajos     = [p for p in productos if p["total"] > 0 and (p["dias"] is None or p["dias"] > 3) and p["total"] <= 8]
    ok        = [p for p in productos if p["total"] > 8 and (p["dias"] is None or p["dias"] > 3)]

    def badge(tipo):
        estilos = {
            "sin_stock": "background:#FCEBEB;color:#A32D2D;",
            "critico":   "background:#FAEEDA;color:#854F0B;",
            "bajo":      "background:#FAF3DA;color:#7A6010;",
        }
        textos = {
            "sin_stock": "Sin stock",
            "critico":   "Critico",
            "bajo":      "Bajo stock",
        }
        return f'<span style="display:inline-block;font-size:11px;font-weight:500;padding:2px 8px;border-radius:99px;{estilos[tipo]}">{textos[tipo]}</span>'

    def filas(lista, tipo):
        html = ""
        for p in lista:
            cd = color_dias(p["dias"])
            b  = badge(tipo) + "&nbsp; " if tipo != "ok" else ""
            html += f"""<tr>
          <td style="padding:8px 10px;font-size:13px;border-bottom:0.5px solid #e5e5e2;">{b}{p['nombre']}</td>
          <td style="padding:8px 10px;font-size:13px;text-align:right;border-bottom:0.5px solid #e5e5e2;">{p['vitacura']}</td>
          <td style="padding:8px 10px;font-size:13px;text-align:right;border-bottom:0.5px solid #e5e5e2;">{p['pataguas']}</td>
          <td style="padding:8px 10px;font-size:13px;text-align:right;font-weight:500;border-bottom:0.5px solid #e5e5e2;">{p['total']}</td>
          <td style="padding:8px 10px;font-size:13px;text-align:right;color:{cd};font-weight:500;border-bottom:0.5px solid #e5e5e2;">{dias_label(p['dias'])}</td>
          <td style="padding:8px 10px;font-size:13px;border-bottom:0.5px solid #e5e5e2;">{p['cocinero']}</td>
        </tr>"""
        return html

    def tabla(titulo, lista, tipo, color):
        if not lista: return ""
        return f"""
    <p style="font-size:12px;font-weight:500;text-transform:uppercase;letter-spacing:0.05em;color:{color};margin:20px 0 6px;">{titulo}</p>
    <table style="width:100%;border-collapse:collapse;border:0.5px solid #e5e5e2;border-radius:8px;overflow:hidden;">
      <thead><tr style="background:#f5f4f0;">
        <th style="padding:8px 10px;font-size:11px;font-weight:500;color:#888780;text-align:left;border-bottom:0.5px solid #e5e5e2;">Producto</th>
        <th style="padding:8px 10px;font-size:11px;font-weight:500;color:#888780;text-align:right;border-bottom:0.5px solid #e5e5e2;">Vitacura</th>
        <th style="padding:8px 10px;font-size:11px;font-weight:500;color:#888780;text-align:right;border-bottom:0.5px solid #e5e5e2;">Pataguas</th>
        <th style="padding:8px 10px;font-size:11px;font-weight:500;color:#888780;text-align:right;border-bottom:0.5px solid #e5e5e2;">Total</th>
        <th style="padding:8px 10px;font-size:11px;font-weight:500;color:#888780;text-align:right;border-bottom:0.5px solid #e5e5e2;">Dias</th>
        <th style="padding:8px 10px;font-size:11px;font-weight:500;color:#888780;text-align:left;border-bottom:0.5px solid #e5e5e2;">Cocinero</th>
      </tr></thead>
      <tbody>{filas(lista, tipo)}</tbody>
    </table>"""

    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:20px;background:#f5f4f0;font-family:Arial,sans-serif;">
<div style="max-width:640px;margin:0 auto;">

  <div style="background:#ffffff;border:0.5px solid #e5e5e2;border-radius:12px;padding:20px 24px;margin-bottom:12px;">
    <p style="font-size:18px;font-weight:500;margin:0 0 4px;color:#2c2c2a;">Reporte de produccion</p>
    <p style="font-size:13px;color:#888780;margin:0;">{fecha} &nbsp;·&nbsp; Vitacura y Pataguas</p>
  </div>

  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:12px;">
    <div style="background:#FCEBEB;border-radius:8px;padding:12px;text-align:center;">
      <div style="font-size:24px;font-weight:500;color:#A32D2D;">{len(sin_stock)}</div>
      <div style="font-size:11px;color:#A32D2D;margin-top:2px;">Sin stock</div>
    </div>
    <div style="background:#FAEEDA;border-radius:8px;padding:12px;text-align:center;">
      <div style="font-size:24px;font-weight:500;color:#854F0B;">{len(criticos)}</div>
      <div style="font-size:11px;color:#854F0B;margin-top:2px;">Critico</div>
    </div>
    <div style="background:#FAF3DA;border-radius:8px;padding:12px;text-align:center;">
      <div style="font-size:24px;font-weight:500;color:#7A6010;">{len(bajos)}</div>
      <div style="font-size:11px;color:#7A6010;margin-top:2px;">Bajo stock</div>
    </div>
    <div style="background:#E1F5EE;border-radius:8px;padding:12px;text-align:center;">
      <div style="font-size:24px;font-weight:500;color:#0F6E56;">{len(ok)}</div>
      <div style="font-size:11px;color:#0F6E56;margin-top:2px;">OK</div>
    </div>
  </div>

  <div style="background:#ffffff;border:0.5px solid #e5e5e2;border-radius:12px;padding:16px 20px;">
    {tabla("Sin stock y critico — producir hoy", sin_stock + criticos, "sin_stock", "#A32D2D")}
    {tabla("Bajo stock — producir esta semana", bajos, "bajo", "#854F0B")}
    {tabla("OK — sin urgencia", ok, "ok", "#0F6E56")}
  </div>

  <p style="font-size:12px;color:#888780;text-align:center;margin-top:12px;">Generado automaticamente · Stock actualizado desde Bsale</p>
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
