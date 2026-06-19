# Contexto del proyecto — Reporte de Stock La Cocina

> Este archivo es el puente entre sesiones y entre PCs. Mantenerlo actualizado
> cuando haya cambios importantes. Leerlo siempre al inicio de una sesión nueva.

## Qué es esto

Dashboard de control de producción y stock para **La Cocina de los Ángeles**,
restaurante chileno con 2 sucursales activas: **Vitacura** (cocina central,
produce) y **Las Pataguas** (recibe despachos). Los datos vienen de la API de
**Bsale** (punto de venta chileno).

- Dashboard publicado: rama `gh-pages` (GitHub Pages), se ve desde el teléfono.
- `historial.json` es la única fuente de verdad de todos los movimientos
  (ventas, producción, despachos, consumos).

---

## Estructura de carpetas (siempre trabajar desde Google Drive)

| Carpeta | Ruta | Propósito |
|---------|------|-----------|
| **Producción** | `G:\Mi unidad\repo-temp\` | Repo con git. Push a GitHub = activa Actions = actualiza dashboard publicado. **NUNCA pushear sin aprobación explícita del usuario.** |
| **Pruebas** | `G:\Mi unidad\reporte-stock-PRUEBAS\` | Copia sin git. Aquí se desarrolla y prueba todo. Cuando está listo, se copia a repo-temp y se sube. |

> OneDrive (`C:\Users\alamo\OneDrive\Escritorio\`) ya no se usa. Todo va en Google Drive.

---

## Cómo funciona el sistema

1. **GitHub Actions** corre `actualizar_diario.py` todos los días a las 19:00
   hora de Chile (el plan gratuito lo retrasa 1–3 h — es normal).
2. El script descarga ventas/guías nuevas de Bsale, infiere producción y
   consumos comparando contra el stock real, y actualiza `historial.json`.
3. `generar_dashboard.py` genera `dashboard.html` (CSS + JS + HTML embebidos
   como strings de Python) y se publica en gh-pages.

---

## Decisiones importantes (no re-discutir sin motivo)

**Datos:**
- `historial.json` — lista de movimientos con `{fecha, sku, oficina, tipo, cantidad, doc_id}`. Tipos: `venta`, `produccion`, `despacho`, `consumo`.
- Ventas reales = `tipo == 'venta'` (equivale a `Movimiento de salida == 'BOLETA'` en la API).
- **Zapallar (3ª sucursal) cerró a mediados de enero 2026**: sus despachos históricos no son anomalías.
- **CAMAC y BAR** son productos de ocasión (Navidad / Día de la Madre): no evaluar por velocidad de venta.

**Técnico — generar_dashboard.py:**
- El HTML se genera con placeholders (`CSS_PLACEHOLDER`, `DATA_PLACEHOLDER`, `ANA_DATA_PLACEHOLDER`, `JS_PLACEHOLDER`, `FECHA_HOY_PLACEHOLDER`).
- `ANA_DATA_PLACEHOLDER` debe reemplazarse ANTES que `DATA_PLACEHOLDER` (es substring del otro).
- Dentro de strings `"""..."""` de Python, `\'` se procesa como `'` (no como `\'`). Para onclick en JS usar atributos `data-*` en vez de parámetros entre comillas simples.
- **Hora de Chile siempre** (`zoneinfo America/Santiago`): GitHub corre en UTC y las corridas retrasadas marcaban el día siguiente. Corregido 11-jun-2026.
- **Reintentos de API**: 3 intentos con 30 s de espera; si fallan los 3, abortan sin tocar datos.
- **Buscador de productos ignora tildes y eñes** (12-jun-2026): la función `norm()` normaliza texto antes de comparar, así "salmon" encuentra SALMÓN y "lasana" encuentra LASAÑA. Aplica al nombre y al SKU. Botón ✕ dentro del buscador borra el campo con un toque.
- **Tabla de movimientos — filtro por tienda** (11-jun-2026, commit b74a70b): botones Todas / Vitacura / Pataguas. Stock coloreado verde `#275300` para Vitacura, azul `#1960a6` para Pataguas. Filas con stock = 0 en fondo rojo suave.

**Diseño (jun-2026, basado en proyecto Stitch 5695411200869491903):**
- Paleta: verde `#275300`, azul `#1960a6`, rojo `#ba1a1a`, bordes `#c2c9b7`.
- Navegación: tabs superiores + barra inferior móvil.
- Vistas: Resumen, Productos, Guías, Ranking, **Análisis**.

**Guía de despacho:**
- Verde `#27AE60` = despacho completo, naranja `#E67E22` = parcial, gris = hay que producir, OK = Pataguas ya cubierta. Si Vitacura está sin stock, todo cae en gris: es el comportamiento correcto.

---

## Tab Análisis (desarrollado jun-2026, en pruebas — aún no está en producción)

Pestaña completa de análisis de ventas mensuales. Estado actual en `reporte-stock-PRUEBAS`:

**Qué hace:**
- Chips de selección de mes (últimos 6 meses).
- **Alerta temprana**: compara el ritmo actual (primeros N días) contra los mismos N días de meses anteriores comparables (excluye meses con festividades en ese período). Si va -30% o más → banner rojo. Si -10% a -30% → amarillo. Si bien → verde.
- Diagnóstico: unidades al día, proyección al cierre (ponderada por patrón de día de semana), contexto del mes (vacaciones, feriados, festividades).
- Métricas: unidades vendidas, días con quiebre, días especiales, semanas activas.
- Comparativa de meses (barras).
- Distribución por día de la semana (barras, Vie/Sáb resaltados).
- Calendario visual con tooltip al hover: feriados (azul), festividades (verde), vacaciones (amarillo).
- Tabla de productos ordenable.

**Datos clave en ANA_DATA:**
- `por_mes[mes].por_dia` — lista de 7 valores (% por día de semana Lun-Dom).
- `por_mes[mes].por_dia_num` — dict `{dia_del_mes: unidades}` (1-31). Usar este para comparación de ritmo (NO `por_dia`).
- `por_mes[mes].vacaciones` — lista de `{dia, nombre}` (ya NO es lista de enteros).
- Calendarios MINEDUC 2026: verano 1-ene al 3-mar, invierno 22-jun al 3-jul. No hay receso de otoño en 2026 en régimen semestral.

**Lo que falta definir (pendiente de conversación):**
- Qué gráficos mantener, cuáles quitar o agregar.
- Análisis por día del mes (patrón de cobro fin de mes + fin de semana) — idea explorada pero se necesitan más meses para confirmar patrón.

---

## Forma de trabajo

- Comunicación **siempre en español**.
- El usuario no es programador: explicar el porqué en lenguaje simple antes de ejecutar.
- **Nunca pushear a GitHub sin aprobación explícita** ("listo súbelo" / "ok subamos"). Cada push activa Actions que envía emails.
- Cambios riesgosos → probar en `reporte-stock-PRUEBAS` primero.
- Validar cálculos mostrando impacto en todos los productos antes de dar por buenos.
- Al iniciar sesión en PC 2: `git pull` en `repo-temp` para tener lo último de producción.

---

## Estado al 18-jun-2026

**En producción (repo-temp / gh-pages):**
- Diseño móvil + reintentos de API + hora de Chile (commit 2440c39, 11-jun).
- Filtro por tienda en tabla de movimientos (commit b74a70b, 11-jun).
- Buscador ignora tildes/eñes + botón ✕ (12-jun).
- **Tab Análisis completo** (18-jun): alerta temprana, diagnóstico, métricas, comparativa mensual con unidades por barra, mapa de calor de ventas por día del mes, tabla de productos con proyección al cierre del mes y tendencia corregida para mes incompleto.

**Cambios clave en tab Análisis (18-jun):**
- Ranking eliminado; navegación: Resumen, Productos, Guías, Análisis.
- Tabla de productos: columna "Proyec. mes" (extrapolación lineal del ritmo actual, siempre ≥ ventas ya registradas). Sin columna "vs Prom." — era confusa.
- Tendencia ("Sube/Baja/Estable") corregida: para el mes en curso compara la proyección, no el parcial, vs el mes anterior.
- Mapa de calor: reemplaza los gráficos de barras "por semana del mes" y "por día de semana". Verde más oscuro = más ventas ese día. Días especiales marcados con punto de color.
- Buscador restaurado: ignora tildes/eñes, botón ✕ para borrar.
- `actualizar_local.bat` + `token.txt` en .gitignore (script de actualización local).

**Correcciones post-lanzamiento (18-jun, mismo día):**
- Mapa de calor: el número del día quedaba casi invisible (8px, opacidad 0.65). Corregido a 10px, opacidad 0.8, siempre visible junto a las unidades.
- Comparativa mensual: el número sobre la barra activa (mes en curso) se superponía o cortaba por usar `position:absolute` calculado según la altura de la barra. Patrón corregido: el número va en su propia fila de flujo normal arriba de un "track" de altura fija (`.month-track`, 90px), la barra crece dentro de ese track — nunca se superpone sin importar la altura.
- Se exploraron alternativas (línea con área, barras horizontales) pero el usuario prefirió mantener barras verticales, solo con el bug de posicionamiento corregido.

**Sincronización multi-PC:** skills (`analisis-stock-restaurante`, `frontend-design`, `ui-ux-pro-max`) y el MCP de Stitch (`.mcp.json`) ahora viven en `G:\Mi unidad\repo-temp\.claude\` y `.mcp.json` — protegidos en `.gitignore` porque el repo es público. Se sincronizan entre PCs vía Google Drive sin tocar GitHub. (Se detectó y corrigió un incidente: una copia completa de `~/.claude/` con credenciales había quedado expuesta en Drive — eliminada.)

**Migración:** carpeta de trabajo movida de OneDrive a Google Drive (`G:\Mi unidad\`) para acceso desde ambos PCs.

---

## Rediseño pendiente (rama `rediseno`)

- Rama `rediseno` creada el 18-jun-2026 desde el commit estable `08b16d2`.
- El rediseño se desarrolla en `reporte-stock-PRUEBAS` como siempre.
- Cuando esté listo: copiar `generar_dashboard.py` a repo-temp, hacer checkout a `rediseno`, commitear y pushear.
- Para fusionar a producción: PR de `rediseno` → `main` (o merge directo si todo está bien).
- `main` NO se toca durante el rediseño — GitHub Actions sigue publicando la versión estable.

---

## Proyecto futuro

Dashboard de stock para la **pizzería delivery del hermano** del usuario,
reutilizando la metodología de este proyecto (skill `analisis-stock-restaurante`).
