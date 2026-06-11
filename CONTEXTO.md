# Contexto del proyecto — Reporte de Stock La Cocina

> Este archivo resume el estado y las decisiones del proyecto para que cualquier
> sesión de Claude (en cualquier PC) pueda retomar el trabajo. Mantenerlo
> actualizado cuando haya cambios importantes.

## Qué es esto

Dashboard de control de producción y stock para **La Cocina de los Ángeles**,
restaurante chileno con 2 sucursales activas: **Vitacura** (cocina central,
produce) y **Las Pataguas** (recibe despachos). Los datos vienen de la API de
**Bsale** (punto de venta chileno).

- Dashboard publicado: rama `gh-pages` (GitHub Pages), se ve desde el teléfono.
- `historial.json` es la única fuente de verdad de todos los movimientos
  (ventas, producción, despachos, consumos).

## Cómo funciona

1. **GitHub Actions** corre `actualizar_diario.py` todos los días a las 19:00
   hora de Chile (el plan gratuito lo retrasa 1–3 h sin patrón — es normal).
2. El script descarga ventas/guías nuevas de Bsale, infiere producción y
   consumos comparando contra el stock real, y actualiza `historial.json`.
3. `generar_dashboard.py` genera `dashboard.html` (CSS/JS/HTML embebidos como
   strings de Python) y se publica en gh-pages.

## Decisiones importantes (no re-discutir sin motivo)

- **Hora de Chile siempre** (`zoneinfo America/Santiago`) en ambos scripts:
  GitHub corre en UTC y las corridas retrasadas marcaban fechas del día
  siguiente. Corregido 11-jun-2026.
- **Reintentos de API**: `bsale_stock()` y `get_all()` intentan 3 veces con
  30 s de espera; si fallan los 3, abortan SIN tocar datos.
- **CAMAC y BAR son productos de ocasión** (Navidad / Día de la Madre):
  no evaluarlos por velocidad de venta ni alertar por baja rotación.
- **Zapallar (3ª sucursal) cerró a mediados de enero 2026**: sus despachos
  históricos no son anomalías.
- **Guía de despacho**: verde `#27AE60` = despacho completo, naranja `#E67E22`
  = parcial, gris = hay que producir (Vitacura sin stock disponible),
  OK = Pataguas ya cubierta. Si Vitacura está sin stock, todo cae en gris:
  es el comportamiento correcto.
- **Diseño móvil** (jun-2026, basado en proyecto Stitch 5695411200869491903):
  paleta verde `#275300`, azul `#1960a6`, rojo `#ba1a1a`, bordes `#c2c9b7`;
  chips de filtro, tarjetas con stock en 2 columnas, barra de cobertura con
  gradiente y aguja, navegación inferior en móvil.

## Forma de trabajo del usuario

- Comunicación **siempre en español**. No es programador: explicar el porqué
  de las cosas en lenguaje simple antes de ejecutar.
- **Cambios riesgosos se prueban primero en un clon local** (carpeta
  `reporte-stock-PRUEBAS`, copia sin git) y **NO se sube nada a GitHub sin
  su aprobación explícita** ("listo súbelo" / "ok subamos").
- Validar cálculos mostrando el impacto en todos los productos antes de
  darlos por buenos.
- Flujo entre PCs: GitHub es la fuente de verdad; `git pull` antes de
  trabajar, push (con aprobación) al terminar.

## Estado al 11-jun-2026

Subido a producción (commit 2440c39): rediseño móvil + reintentos de API +
fecha en hora de Chile.

**Pendiente de verificar:**
- Que el dashboard publicado se vea bien en el teléfono tras la corrida nocturna.
- Que las fechas de despachos/recepciones salgan correctas los próximos días
  (antes salían con fecha del día siguiente cuando GitHub se retrasaba).

## Proyecto futuro

Dashboard de stock para la **pizzería delivery del hermano** del usuario,
reutilizando la metodología de este proyecto (skill `analisis-stock-restaurante`).
