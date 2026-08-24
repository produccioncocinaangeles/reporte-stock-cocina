# -*- coding: utf-8 -*-
"""Correo diario de produccion.

NO calcula nada: lee datos_reporte.json que deja generar_dashboard.py y solo
arma y envia el HTML. Antes este script bajaba el stock de Bsale por su cuenta
y recalculaba los dias con su propia formula y sus propios umbrales, asi que el
correo y el dashboard se contradecian (41 de 56 productos mostraban dias
distintos y 20 tenian semaforo distinto). Ahora la unica fuente de verdad es el
dashboard.
"""

import smtplib
import os
import json
import sys
import unicodedata
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_DATOS = os.path.join(CARPETA, 'datos_reporte.json')

# Orden en que se muestran las secciones por cocinero.
COCINEROS = ["CAROLINA", "CESAR", "JESUS", "ADRIANA"]


def sin_tildes(texto):
    """CESAR y CESAR con tilde son el mismo cocinero (el dashboard usa tildes)."""
    return "".join(
        c for c in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(c) != "Mn"
    ).upper()


# ============================================================
# PASO 1: LEER LOS DATOS DEL DASHBOARD
# ============================================================

def cargar_datos():
    """Datos ya calculados por generar_dashboard.py (paso 1 del workflow).

    Si el archivo no esta o quedo de un dia anterior, cortamos: el paso 1 fallo
    y mandar el correo igual significaria enviarle a la cocina numeros viejos
    con fecha de hoy.
    """
    if not os.path.exists(ARCHIVO_DATOS):
        sys.exit(
            "ERROR: no existe datos_reporte.json.\n"
            "       Lo genera generar_dashboard.py; si no esta, ese paso fallo.\n"
            "       No se envia el correo para no mandar datos incorrectos."
        )

    with open(ARCHIVO_DATOS, encoding="utf-8") as f:
        datos = json.load(f)

    productos = datos.get("productos", [])
    if not productos:
        sys.exit("ERROR: datos_reporte.json no tiene productos. No se envia el correo.")

    generado = datetime.fromisoformat(datos["generado"])
    if (datetime.now() - generado).days >= 1:
        sys.exit(
            "ERROR: datos_reporte.json es del {}, no de hoy.\n"
            "       El dashboard no se actualizo. No se envia el correo.".format(datos["fecha"])
        )

    print("Datos del dashboard: {} productos ({})".format(len(productos), datos["fecha"]))
    return datos


def ordenar(productos):
    """Mas urgente primero, igual que el dashboard."""
    ORDEN = {"sin_stock": 0, "critico": 1, "bajo": 2, "ok": 3}
    return sorted(
        productos,
        key=lambda p: (
            ORDEN.get(p["estado"], 9),
            p["dias"] if p["dias"] is not None else 9999,
        ),
    )


# ============================================================
# PASO 2: ARMAR CORREO HTML
# ============================================================
#
# El diseno sigue el mismo sistema visual del dashboard (mismos colores,
# tipografia y tarjetas), pero maquetado con <table> y CSS inline porque
# display:grid y las hojas de estilo externas no funcionan en Outlook ni en
# varios clientes de correo.

FUENTE = ("'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,"
          "Helvetica,Arial,sans-serif")

# Paleta del dashboard (bloque :root de generar_dashboard.py)
TEXTO    = "#1e293b"   # --text-color
SUAVE    = "#64748b"   # texto secundario
TENUE    = "#94a3b8"   # labels de tabla y KPI
CERO     = "#cbd5e1"   # --zero-color
BORDE    = "#e2e8f0"   # --border-color
LINEA    = "#f1f5f9"   # separador de filas
FONDO    = "#f8fafc"   # --body-bg
VERDE    = "#275300"   # verde del logo
AZUL     = "#1d4ed8"   # --info-text (Pataguas)

# Barra lateral de estado: mismos colores que .card.sin_stock/.critico/... del
# dashboard, donde cada tarjeta lleva un border-left de 6px.
BARRA = {"sin_stock":"#ef4444","critico":"#f97316","bajo":"#facc15","ok":BORDE}
# Color del numero de dias: mismo criterio que los badges de la Guia de
# Produccion (danger <=3, warning <=7, ok mas arriba).
DIAS_COLOR = {"sin_stock":"#b91c1c","critico":"#b91c1c","bajo":"#c2410c","ok":"#166534"}


def titulo_case(texto):
    """Igual que tituloCase() en el dashboard: solo la primera en mayuscula."""
    if not texto:
        return texto
    return texto[0].upper() + texto[1:].lower()


def armar_html(productos, fecha_datos):
    fecha = datetime.now().strftime("%d/%m/%Y · %H:%M")

    criticos      = [p for p in productos if p["estado"] == "critico"]
    bajos         = [p for p in productos if p["estado"] == "bajo"]
    ok            = [p for p in productos if p["estado"] == "ok"]
    sin_stock_vit = [p for p in productos if p["vit"] == 0]
    sin_stock_pat = [p for p in productos if p["pat"] == 0]

    def dias_str(p):
        if p["estado"] == "sin_stock": return "Sin stock"
        if p["dias"] is None:          return "—"
        return f'{round(p["dias"])}d'

    def dias_total_str(p):
        if p["dias_total"] is None: return "—"
        return f'{round(p["dias_total"])}d'

    def num(valor, color=TEXTO):
        """Los ceros van atenuados, como .num-zero en el dashboard."""
        return f'<span style="color:{CERO}">0</span>' if valor == 0 else \
               f'<span style="color:{color}">{valor}</span>'

    # ── Tabla de productos ──────────────────────────────────
    TD  = f'padding:11px 12px;font-size:13px;border-bottom:1px solid {LINEA}'
    TDR = TD + ';text-align:right'
    TH  = (f'padding:10px 12px;font-size:10px;font-weight:700;color:{TENUE};'
           f'text-transform:uppercase;letter-spacing:0.05em;'
           f'border-bottom:1px solid {BORDE}')

    def fila(p):
        clr = DIAS_COLOR[p["estado"]]
        return (
            f'<tr>'
            f'<td style="{TD};border-left:4px solid {BARRA[p["estado"]]};'
            f'color:{TEXTO};font-weight:500">{titulo_case(p["nombre"])}</td>'
            f'<td style="{TDR}">{num(p["vit"])}</td>'
            f'<td style="{TDR}">{num(p["pat"], AZUL)}</td>'
            f'<td style="{TDR};font-weight:700;color:{clr}">{dias_str(p)}</td>'
            f'<td style="{TDR};color:{SUAVE}">{dias_total_str(p)}</td>'
            f'</tr>')

    def tabla(lista):
        if not lista:
            return f'<p style="font-size:13px;color:{TENUE};padding:10px 0;margin:0">Sin productos</p>'
        return (
            f'<table role="presentation" cellpadding="0" cellspacing="0" border="0" '
            f'style="width:100%;border-collapse:collapse;font-family:{FUENTE}">'
            f'<thead><tr>'
            f'<th style="{TH};text-align:left">Producto</th>'
            f'<th style="{TH};text-align:right">Vitacura</th>'
            f'<th style="{TH};text-align:right;color:{AZUL}">Pataguas</th>'
            f'<th style="{TH};text-align:right">Días</th>'
            f'<th style="{TH};text-align:right">Días total</th>'
            f'</tr></thead>'
            f'<tbody>{"".join(fila(p) for p in lista)}</tbody></table>')

    # ── Tarjeta blanca, equivalente a .card del dashboard ────
    def card(contenido, padding="18px 20px"):
        return (
            f'<table role="presentation" cellpadding="0" cellspacing="0" border="0" '
            f'style="width:100%;border-collapse:separate;margin-bottom:14px">'
            f'<tr><td style="background:#ffffff;border:1px solid {BORDE};'
            f'border-radius:12px;padding:{padding}">{contenido}</td></tr></table>')

    # ── KPI, mismo estilo que .card-kpi-individual ───────────
    def kpi(valor, etiqueta, color):
        val = color if valor else CERO
        return (
            f'<td width="20%" style="padding:0 4px" valign="top">'
            f'<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%">'
            f'<tr><td style="background:#ffffff;border:1px solid {BORDE};border-radius:12px;'
            f'padding:14px 8px;text-align:center">'
            f'<div style="font-size:26px;font-weight:800;letter-spacing:-0.02em;'
            f'line-height:1;color:{val};font-family:{FUENTE}">{valor}</div>'
            f'<div style="font-size:10px;font-weight:600;color:{TENUE};margin-top:6px;'
            f'font-family:{FUENTE}">{etiqueta}</div>'
            f'</td></tr></table></td>')

    kpis = (
        f'<table role="presentation" cellpadding="0" cellspacing="0" border="0" '
        f'style="width:100%;border-collapse:separate;margin:0 -4px 14px"><tr>'
        + kpi(len(sin_stock_vit), "Sin stock VIT", "#b91c1c")
        + kpi(len(sin_stock_pat), "Sin stock PAT", "#b91c1c")
        + kpi(len(criticos),      "Crítico ≤3d",   "#c2410c")
        + kpi(len(bajos),         "Bajo ≤7d",      "#c2410c")
        + kpi(len(ok),            "OK",            "#166534")
        + '</tr></table>')

    # ── Secciones por cocinero ──────────────────────────────
    def seccion(titulo, lista, hay_urgentes):
        color = "#b91c1c" if hay_urgentes else SUAVE
        return (
            f'<div style="margin-bottom:20px">'
            f'<div style="font-size:11px;font-weight:700;text-transform:uppercase;'
            f'letter-spacing:0.07em;color:{color};margin-bottom:8px;padding-bottom:6px;'
            f'border-bottom:2px solid {color};font-family:{FUENTE}">{titulo_case(titulo)}</div>'
            f'{tabla(lista)}</div>')

    secciones_cocinero = ""
    for coc in COCINEROS:
        prods_coc = [p for p in productos if sin_tildes(p["cocinero"]) == coc]
        if prods_coc:
            urgentes = any(p["estado"] in ("sin_stock", "critico") for p in prods_coc)
            # El titulo sale del dato, no de la constante, para conservar las
            # tildes que usa el dashboard (Cesar -> Cesar con tilde).
            secciones_cocinero += seccion(prods_coc[0]["cocinero"], prods_coc, urgentes)

    titulo_card = (f'font-size:14px;font-weight:700;color:{TEXTO};margin:0 0 14px;'
                   f'font-family:{FUENTE}')

    # ── Header con el logo del dashboard ────────────────────
    header = (
        f'<table role="presentation" cellpadding="0" cellspacing="0" border="0" '
        f'style="width:100%;margin-bottom:16px"><tr>'
        f'<td valign="middle" style="width:38px;padding-right:10px">'
        f'<div style="width:34px;height:34px;background:{VERDE};border-radius:8px;'
        f'text-align:center;line-height:34px;font-size:17px;color:#fff">&#127869;</div></td>'
        f'<td valign="middle">'
        f'<div style="font-size:16px;font-weight:800;letter-spacing:-0.03em;color:{TEXTO};'
        f'font-family:{FUENTE}">La Cocina '
        f'<span style="font-size:12px;font-weight:400;color:{SUAVE};letter-spacing:0">'
        f'· Reporte de Producción</span></div>'
        f'<div style="font-size:11px;color:{SUAVE};margin-top:3px;font-family:{FUENTE}">'
        f'{fecha} · Prioridad por stock Vitacura</div></td>'
        f'</tr></table>')

    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
</style>
</head>
<body style="margin:0;padding:24px 16px;background:{FONDO};font-family:{FUENTE};color:{TEXTO}">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%">
<tr><td align="center">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:660px">
<tr><td>

{header}

{kpis}

{card(f'<p style="{titulo_card}">Por cocinero</p>{secciones_cocinero}')}

{card(f'<p style="{titulo_card}">Lista completa · {len(productos)} productos</p>{tabla(productos)}')}

<p style="font-size:11px;color:{TENUE};text-align:center;margin:4px 0 0;line-height:1.7;font-family:{FUENTE}">
  Generado automáticamente · Mismos datos del dashboard ({fecha_datos})<br>
  <b style="color:{SUAVE}">Días</b>: cuánto dura el stock de Vitacura al ritmo de venta de las dos tiendas<br>
  <b style="color:{SUAVE}">Días total</b>: lo mismo contando también el stock de Pataguas
</p>

</td></tr></table>
</td></tr></table>
</body></html>"""
    return html


# ============================================================
# PASO 3: ENVIAR CORREO
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
    print("REPORTE DE PRODUCCION (datos del dashboard)")
    print("=" * 50)

    datos     = cargar_datos()
    productos = ordenar(datos["productos"])
    html      = armar_html(productos, datos["fecha"])
    enviar_correo(html)

    print("=" * 50)
    print("PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 50)


if __name__ == "__main__":
    main()
