# -*- coding: utf-8 -*-
"""Backfill de recepciones ('produccion') de los 14 postres de proveedor en historial.json.
Tramo 1: copia desde historial_completo.json. Tramo 2: baja de Bsale desde 2026-08-28.
Idempotente. Luego compara el stock reconstruido contra /v1/stocks.json."""
import json, os, shutil, sys
from collections import defaultdict
import requests
import descargar_historial_completo as d  # sin efectos secundarios (main protegido)

C = os.path.dirname(os.path.abspath(__file__))
# Uso: python backfill_postres.py [ruta\historial.json]  (por defecto el de esta carpeta).
# Necesita historial_completo.json y descargar_historial_completo.py junto al script
# (están en reporte-stock-REDISENO), por eso se corre desde ahí.
H = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.join(C, "historial.json")
HC = os.path.join(C, "historial_completo.json")
BK = os.path.join(os.path.dirname(H), "historial_backup_antes_postres.json")
SKUS = set("CBN CDL CDM CHF FC GHD P TBN TCF TCH THM TMF THF HDV".split())
CORTE = "2026-08-28"

def clave(m):
    return (m["fecha"], m["sku"], m["oficina"], m["tipo"], m["doc_id"])

def main():
    if not os.path.exists(BK):  # respaldo único del estado original
        shutil.copy2(H, BK)
    with open(H, encoding="utf-8") as f:
        hist = json.load(f)
    movs = hist["movimientos"]
    fmax = max(m["fecha"] for m in movs)
    with open(HC, encoding="utf-8") as f:
        comp = json.load(f)["movimientos"]

    cand = [m for m in comp if m["tipo"] == "produccion" and m["sku"] in SKUS and m["fecha"] < CORTE]
    # historial.json solo tiene ventas/despachos de estos SKUs desde la primera
    # fecha en que aparecen; lo anterior (ene-may) se trae también de completo,
    # si no las entradas de ese tramo quedan sin salidas y el stock se infla.
    primera = min((m["fecha"] for m in movs if m["sku"] in SKUS and m["tipo"] != "produccion"), default=CORTE)
    cand += [m for m in comp if m["tipo"] in ("venta", "despacho") and m["sku"] in SKUS and m["fecha"] < primera]
    n1 = len(cand)
    headers = {"access_token": d.leer_token()}
    from datetime import datetime, timezone
    mapa = d.construir_mapa_variantes(headers)
    desde = datetime.strptime(CORTE, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    hasta = datetime.strptime(fmax, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    bsale = [m for m in d.descargar_recepciones(headers, desde, hasta, mapa) if m["sku"] in SKUS]
    # tramo copiado de completo que ya cubre >= CORTE se ignora a favor de Bsale (mismo dato)
    cand += bsale

    existentes = {clave(m) for m in movs}
    agregados = defaultdict(int)
    nuevos = []
    for m in cand:
        k = clave(m)
        if k in existentes:
            continue
        existentes.add(k)
        nuevos.append(m)
        agregados[(m["sku"], m["oficina"])] += 1
    movs.extend(nuevos)
    movs.sort(key=lambda m: m["fecha"])  # sort estable: conserva orden relativo
    with open(H, "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=2)
    print(f"Tramo 1 (historial_completo: entradas <{CORTE}, salidas <{primera}): {n1} candidatas; tramo 2 (Bsale {CORTE}..{fmax}): {len(bsale)}")
    print(f"Agregadas: {len(nuevos)}")
    for (s, o), n in sorted(agregados.items()):
        print(f"  {s:4s} {o}: {n}")

    # chequeo contra stock real
    real = defaultdict(float)
    for oid, of in d.OFICINAS.items():
        for it in d.get_paginado(f"{d.BASE}/stocks.json", headers, {"officeid": oid, "expand": "[variant,office]"}):
            v = it.get("variant") or {}
            code = str(v.get("code", "")).strip().upper() or mapa.get(str(v.get("id", "")), "")
            if code in SKUS:
                real[(code, of)] += float(it.get("quantity", 0))
    rec = {}
    for of in d.OFICINAS.values():
        acum = defaultdict(float)
        for m in sorted([x for x in movs if x["oficina"] == of and x["sku"] in SKUS],
                        key=lambda m: (m["fecha"], m["tipo"] != "produccion")):
            e = m["cantidad"] if m["tipo"] == "produccion" else 0
            s = m["cantidad"] if m["tipo"] in ("venta", "despacho", "consumo") else 0
            acum[m["sku"]] = max(0, acum[m["sku"]] + e - s)
        for s in SKUS:
            rec[(s, of)] = acum.get(s, 0)
    print("\nSKU  OF  reconstruido  real  dif")
    dif = 0
    for k in sorted(rec):
        r, b = rec[k], real.get(k, 0)
        if abs(r - b) > 1e-6:
            dif += 1
        print(f"{k[0]:4s} {k[1]} {r:10.1f} {b:8.1f} {r-b:+8.1f}")
    print(f"Pares con diferencia: {dif} de {len(rec)}")

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
