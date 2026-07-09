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
- **Reglas permanentes del usuario** (versión completa en
  `C:\Users\<usuario>\.claude\CLAUDE.md` de cada PC; copia maestra en
  `G:\Mi unidad\Claude code\CLAUDE-global-copiar-a-cada-PC.md`):
  1. Contradecirlo con razones si su idea es mala, ANTES de ejecutar.
  2. Investigar métodos/herramientas que ya existen hechos por expertos antes
     de inventar algo a medida (nombrar el método estándar de la industria).
  3. Opciones antes de construir: 2-3 alternativas de simple a compleja, con
     pros y contras — él elige.
- **Regla de soluciones simples (aprendida a costo alto, jul-2026):** antes de
  proponer cualquier solución técnica, elegir la MÁS SIMPLE que cumpla: sin
  instalar programas, que funcione desde el teléfono, gratis, y de preferencia
  dentro del ecosistema Google (Sheets, Apps Script) donde ya viven los datos.
  Contexto: para "que varias personas carguen recetas desde el teléfono" una
  sesión anterior propuso una extensión de Chrome y luego Tailscale (instalar
  en PC y teléfono) — se gastaron horas y tokens antes de llegar a Apps Script,
  que era la respuesta obvia. Extensiones, VPNs y servidores locales son
  ÚLTIMO recurso, no primera propuesta.
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

## Estado al 22-jun-2026

**Tabla "Contribución por producto" (tab Análisis) — vuelta a definir varias veces, estado final:**
- Columna **"Proyec. histórica"** = mismo valor que `lote_sugerido` de la pestaña Productos ("Consumo estimado 30 días"), NO un cálculo nuevo. Decisión explícita del usuario tras varias iteraciones (se probó extrapolación lineal del mes y distribución proporcional de la proyección general — ambas se descartaron).
- Debajo del número, si hay ≥2 meses de historial guardado, aparece un indicador "+X% vs YYYY-MM" cuando el cambio es ≥10% — usa el nuevo archivo `historial_proyecciones.json`.
- Columna **"vs mes ant."** (antes "Tendencia"): compara la proyección de este mes vs el mes anterior. Para evitar ruido en productos de bajo volumen, si ambos valores son menores a 10 unidades, se fuerza "Estable" en vez de Sube/Baja (variable `MIN_VOL_TEND` en `calcular_analisis()`).

**Nuevo: `historial_proyecciones.json`** — snapshot mensual (NO diario) de `lote_sugerido` por SKU, para poder comparar cómo cambia la proyección histórica con el tiempo. Se guarda una sola vez por mes (la primera corrida de ese mes). Archivo pequeño (~1 KB/mes) — **debe estar en git, NO en .gitignore**, para que persista entre corridas de GitHub Actions.

**Filtro de tienda restaurado** en tabla de Movimientos (Productos tab) — se había perdido en algún punto entre ediciones, igual que pasó antes con el buscador sin tildes. Botones Todas/Vitacura/Pataguas, función `filtrarMovs()`.

**Validación de proyecciones con datos reales (22-jun):**
- Verificado que las vacaciones bajan el ritmo de venta a casi la mitad (enero-febrero vacaciones de verano: 11-15 un./día vs marzo normal: 22.5 un./día). Las vacaciones de invierno (22-jun al 3-jul) cubren el resto de junio — cualquier proyección de cierre de mes debe descontar este efecto, no asumir que el ritmo se mantiene.
- El Día del Padre (21-jun) NO generó un pico de ventas ese día específico (28 un., normal para domingo). Hubo picos inusuales el 17-18 jun (mié-jue) sin causa clara identificada.

**Pendiente / ideas no implementadas:**
- Pestaña nueva **"Rendimiento de cocinero"**: el usuario quiere un ícono discreto en el header (no un botón de nav normal) que pida una clave simple antes de mostrar la vista. Importante: como el repo es público, esto NO es seguridad real, solo evita que alguien la encuentre navegando casualmente. Falta definir qué métricas mostrar exactamente (¿unidades producidas por cocinero?, ¿quiebres asociados?). No implementado todavía.

---

## Estado al 3-jul-2026 (commit 00d740f)

**Pestaña Análisis — mes nuevo y claridad (todo en producción):**
- El mes en curso aparece desde el día 1 (se eliminó el filtro que lo ocultaba hasta el 40% del mes).
- Comparativa mensual: solo últimos 6 meses; los chips mantienen hasta 12.
- Primeros 7 días del mes: banner gris "⏳ Mes recién comenzando" (sin veredicto ni proyección — con 2 días proyectaba ~700 un., puro ruido) y columna "vs mes ant." en "—" (`sin_datos`). Desde el día 7 vuelve el análisis normal.
- **Tarjeta "Ritmo de producción"** (solo visible en el mes en curso): crecimiento compuesto (CMGR) de los últimos 4 meses completos excluyendo enero-febrero (vacaciones), más referencia de cierre (promedio de últimos 3 meses). Cuando exista el mismo mes del año anterior en el historial cambia automáticamente a comparación año contra año (YoY) — ocurrirá desde enero 2027. Datos jul-2026: Mar 698 → Abr 667 → May 824 → Jun 756 (+2,7% mensual) → referencia julio ~749.
- **Columna "Proyec. histórica"**: en meses pasados muestra la foto de `historial_proyecciones.json` de ese mes (para validar proyectado vs vendido real); en el mes en curso muestra la proyección vigente + etiqueta "ajuste: ±X% vs jun" cuando cambió ≥10%. El detalle a fondo ("en jun se estimaban 31 und, ahora 35 (+13%) — el producto se aceleró") está en Recomendaciones de la pestaña Productos.
- **Productos con venta 0** aparecen al final de la tabla (0 en rojo) si vendieron en alguno de los 3 meses previos — un producto activo que no vendió nada es señal, no dato faltante.
- Leyenda plegable "ℹ️ ¿Cómo leer esta tabla?" sobre la tabla de contribución.
- **"Vitacura sin stock"** (detalle de producto): consolidado de días por mes a simple vista + fechas exactas en desplegable. Ojo: las fechas de `periodos_sin_stock_vit` vienen en formato dd/mm/yyyy — el JS las convierte con `fIso()`; el día "fin" es cuando volvió el stock y no se cuenta.

**Regla de diseño clave (guardada en memoria):** el dashboard lo usan terceros que no pueden preguntar — todo cálculo debe explicarse solo en la UI (mostrar valores de origen, no solo el resultado; leyendas en lenguaje simple).

**Dato importante del negocio:** el dashboard cubre SOLO productos de producción propia (56 SKUs en NOMBRES). En junio hay ~20 SKUs fuera de catálogo (PP, PN, BURMA, D, LF...) con 337 un. vendidas — son reventa/otros, excluidos a propósito. Total junio catálogo: 756 un. (no 1.093, que incluye lo de fuera).

**`CLAUDE.md`** creado en repo-temp (va en git), PRUEBAS y REDISENO (locales): contexto automático para sesiones nuevas en cualquier carpeta.

**Pendientes:**
- Extender ventana de velocidad cuando días con stock < 20 en 3 meses (caso RPP con lote inflado a 33) — confirmado, no implementado.
- Alerta de "ritmo lento" al final del mes debe considerar festividades próximas en los días restantes (limitación metodológica discutida, sin implementar).
- Pestaña "Rendimiento de cocinero" con ícono discreto + clave simple — solo idea.
- Pasada completa de leyendas explicativas al resto del dashboard (alerta temprana, calendario, guías).

---

## Estado al 9-jul-2026 — Pestaña Producción (SUBIDO a main, commit de26d52)

> Subido el 9-jul con aprobación del usuario ("carga la nueva versión al github").
> Incluye además: render perezoso del detalle en la vista Productos (cambiar el
> filtro de cocinero congelaba el teléfono) y fix de `verMovsCompletos` con filtro
> activo. Se fusionó con 3 vías contra el commit 00d740f (Análisis 3-jul) — ambas
> ramas de trabajo quedaron combinadas; REDISENO y repo-temp tienen el mismo
> `generar_dashboard.py`. PENDIENTE: el usuario debe crear el secret
> `GOOGLE_SERVICE_ACCOUNT_JSON` en GitHub para que las recetas aparezcan en
> producción (sin él, la pestaña Producción sale con recetas vacías).

La pestaña **Recetas se transformó en "Producción"**: planificador semanal donde
el usuario decide cuánto producir (definido en entrevista el 8-9 jul; maqueta
aprobada en `maqueta_produccion.html`). Implementado en `generar_dashboard.py`
de `reporte-stock-REDISENO` (ids internos siguen siendo `vista-recetas` /
`nav-recetas` / `mostrarRecetas` para minimizar cambios).

**Qué hace:**
- Cards por producto ordenadas por urgencia (orden de DATA), con stock
  Vit/Pat, cobertura, cocinero y **sugerencia semanal** (`vel_total*7`) como
  referencia — la cantidad la decide el usuario con stepper −/+.
- Filtros: buscador, chips por urgencia/A-Z y por cocinero (Carolina, Adriana,
  César, Jesús — de `COCINEROS`).
- Detalle por producto: receta por 1 unidad + **cálculo inverso** ("tengo 5 kg
  de atún → alcanza para 10 und → Usar esta cantidad").
- Barra verde fija inferior con resumen del plan → hoja (bottom sheet) con
  materia prima consolidada o por producto. Consolidación normaliza g→kg y
  ml→L; ingredientes sin cantidad o productos sin receta se listan como ⚠
  avisos, nunca se suman en silencio.
- **Compartir por menú nativo** (`navigator.share`, respaldo wa.me sin
  número): también se cambió el modal de ingredientes de Guías (se eliminó el
  campo de número de WhatsApp).
- `leer_recetas()` ahora marca ingredientes "se prepara" leyendo la pestaña
  `Elaborados` de la planilla (hoy vacía — al llenarla aparecen las etiquetas
  solas). El plan NO se guarda todavía (queda como paso siguiente: Apps Script).

**Contexto de la decisión:** el dolor real es que el cocinero dicta ingredientes
de memoria y las compras salen incompletas varias veces por semana (caso
tártaro de atún: se compró cebolla+cilantro, faltó jengibre). El usuario (jefe
de cocina) arma el plan semanal; su jefe compra — recibe la lista por WhatsApp.

---

## Sistema de recetas (jun-2026, vive en `reporte-stock-REDISENO`)

Sistema paralelo al dashboard para digitalizar las recetas de los productos.
Los archivos están en `G:\Mi unidad\reporte-stock-REDISENO\` (no en git).

**Dónde viven los datos:**
- Google Sheet **"La Cocina - Recetas"** (en la raíz de `G:\Mi unidad`), ID
  `1bzRCgI5Cs0mIvjnFvfJZiZ6K3DaNU8D69jvo0vhU49g`.
- Pestañas: `Recetas` (Producto | SKU | Ingrediente | Cantidad | Unidad, con
  dropdowns), `Ingredientes` (lista maestra, evita duplicados tipo "queso" vs
  "queso mantecoso"), `Productos` (auxiliar Producto→SKU) y `Elaborados`.
- Acceso vía cuenta de servicio (`la-cocina-498520-*.json`, gspread).

**Scripts:**
- `crear_planilla_recetas.py` — arma/repara la estructura de la Sheet desde cero.
  Seguro de re-ejecutar (no borra datos cargados).
- `importar_recetas_excel.py` — importó la matriz de ingredientes desde
  `Copia de Costos SEP 2023.xlsx` (hoja INSUMOS). Unidad `kg` por defecto en
  todo (decisión del usuario, 30-jun-2026), se corrige fila a fila después.
  Incluye mapeos manuales de nombres viejos→SKU y exclusiones confirmadas
  (ej. Pastel de Choclo Mediano quedó fuera).
- `editor_recetas.py` + `abrir_editor_recetas.bat` — editor local en el
  navegador (localhost) para agregar/editar/borrar ingredientes manteniendo
  las filas de cada producto juntas. Es el ÚNICO lugar donde se manejan datos
  sensibles (precio/proveedor/merma, elaborados).
- `apps_script_Code.gs` + `apps_script_Index.html` — **editor online** (Google
  Apps Script como aplicación web): un link para cargar recetas desde el
  teléfono/WhatsApp sin PC encendida. NO expone precios ni proveedores; usa un
  candado para escrituras simultáneas. Instrucciones de publicación en
  `INSTRUCCIONES_EDITOR_ONLINE.md`.

**Por confirmar:** si el editor online ya fue implementado en script.google.com
(y cuál es el link).

### Hoja de ruta del sistema de recetas (definida 3-jul-2026)

Visión completa del usuario, por etapas — cada una útil por sí sola:

1. **Cargar todas las recetas** (en curso). Base de todo lo demás.
2. **Explosión de materiales**: producción sugerida (lote_sugerido, ya existe
   en el dashboard) × recetas = cuánta materia prima se necesita.
3. **Formatos de proveedor + costos + merma**: por ingrediente, en qué formato
   vende el proveedor (ej. barra de queso 3,5 kg), precio y % merma. Permite:
   convertir necesidad en unidades de compra reales, generar **texto de pedido
   por proveedor** (listo para WhatsApp) y calcular **costo real por producto**.
   Datos sensibles solo en el editor local (decisión ya tomada).
4. **Registro de facturas → inventario de materias primas + gasto por
   proveedor**. Contexto: facturas llegan EN PAPEL (las recibe quien recibe la
   mercadería; no hay personal administrativo). Plan acordado:
   - **Fase 1 — manual**: formulario en el editor local donde el usuario carga
     las líneas de la factura. Valida la lógica de inventario antes de
     automatizar.
   - **Fase 2 — foto + IA**: quien recibe saca foto de la factura → carpeta de
     Drive/WhatsApp → la IA la lee → el usuario **confirma/corrige antes de
     guardar** (paso de confirmación obligatorio, nunca guardar directo).
   - Con inventario de ingredientes + recetas → "¿qué puedo producir con lo
     que tengo?" (la funcionalidad más valiosa: no parar producción).
   - Advertencia asumida: el inventario inferido se desvía con el tiempo →
     requiere **conteo físico periódico** que lo re-ancle (mismo concepto que
     el stock de Bsale).

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
