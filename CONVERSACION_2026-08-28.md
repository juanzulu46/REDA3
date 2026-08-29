# Conversación de trabajo — 28 agosto 2026

Registro de la sesión con Claude Code: pedido, decisiones, cambios implementados y cómo funcionan. Sesión corta, **solo visual y solo en `asesor.html`**. Commit `2abeff0` en `master` (empujado a GitHub como respaldo; el canal oficial sigue siendo pegar en Apps Script).

---

## 1. Pedido del usuario

> "En el detallado de los negocios en arriendo no debería aparecer todos los meses. Simplemente es un cambio visual. Si es colegaje, agrega el detalle de que es colegaje, o sea el total menos el colegaje. Y si tiene lo de 2 % también ponerle eso en el detalle."

## 2. Preguntas de Claude y decisiones del usuario

| Pregunta | Decisión |
|---|---|
| ¿Qué mostrar en lugar de la lista mes a mes de "Cobros registrados"? (resumen con montos / quitar la sección / solo rango de fechas) | **Solo el rango de fechas**, sin montos ni novedades. |
| ¿Dónde va el desglose de colegaje y la nota del bono? (tarjeta visible + panel / solo panel expandible) | **En la tarjeta visible y también en el panel "Ver detalle completo"**. |

## 3. Hallazgos de la investigación

- El único sitio que listaba mes por mes era la sección **"Cobros registrados"** del panel expandible (`negExpPanelHTML`, `asesor.html` ~2903): pinta una fila por cada fila de `CobrosArriendo`, y `generarCobrosProyectados` (`apps_script.js:301-321`) crea una por mes del contrato (12 en un contrato normal). Se ve igual en Mis Negocios y en Gestión porque comparten `negExpPanelHTML`.
- Las tarjetas de arriendo ya mostraban badges COLEGAJE / BONO CAPTACIÓN (commit `e7f3521` del 27-ago), pero el desglose numérico solo estaba en el panel expandible y con etiquetas "Comisión bruta / Parte colega / REDA3 neta", no en formato "total − colegaje = neta". Como ese commit **aún no se ha desplegado en Apps Script**, el usuario en producción no veía nada de esto.
- Los datos necesarios ya viajan al navegador: `modalidad`, `pct_colega`, `comision_bruta`, `comision_oficina` (neta), `bono_captacion_pct` en cada arriendo/venta de `mis_negocios` / `todos_negocios`. No hizo falta tocar el backend.
- Helpers ya existentes reutilizados: `esColegajeNeg`, `factorColegaNeg`, `fmtP`, `negBadgesExtra`.

## 4. Qué se cambió y cómo funciona

### 4.1 Periodo de cobros (arriendos, panel expandible)
- La sección "Cobros registrados" pasa a **"Periodo de cobros"** con una sola línea: primer mes → último mes (`Ene 2026 → Dic 2026`). Si hay un solo cobro, solo ese mes. Ordena por `año×12+mes`; los cobros sin fecha se ignoran (si ninguno tiene fecha, "Sin fecha").
- No se muestran montos, estados (COBRADO/NO_COBRADO) ni observaciones. Los datos siguen en la hoja `CobrosArriendo`; solo cambió la vista.

### 4.2 Colegaje: total − colegaje = neta (tarjeta y panel)
- Nuevo helper `colegajeFilasHTML(neg, esArr)` junto a `negBadgesExtra`. Devuelve la fila normal "Comisión [mensual] oficina" si el negocio no es colegaje; si lo es, tres filas para la grilla `neg-detail`:
  1. **Comisión total [mensual] (bruta)**: `comision_bruta` (o `neta ÷ factor` si la bruta no está guardada).
  2. **Colegaje X % [· tipo colega]**: `− (bruta − neta)` en rojo. X = 30 % arriendos; 50 % Inmobiliaria; 33,33 % Persona natural.
  3. **REDA3 neta [mensual]**: `comision_oficina`.
- Se usa en las **cuatro tarjetas**: arriendo y venta, en Mis Negocios y en Gestión (reemplaza la fila "Comisión mensual oficina" / "Comisión mensual" / "Comisión oficina"; las etiquetas quedaron unificadas).
- "Comisión total contrato" de arriendos se etiqueta "(neta)" cuando es colegaje.
- El panel expandible (sección Colegaje) usa el mismo orden y etiquetas, conservando la nota "(base de comisiones, bonificación y cobros)".
- Un negocio normal se ve exactamente igual que antes.

### 4.3 Bono de captación (+2 % / +1 %) en la tarjeta
- Nuevos helpers `bonoPPNeg(neg)` (0, 1 o 2 puntos desde `bono_captacion_pct`) y `bonoFilaHTML(neg)`.
- Tarjetas de arriendo (Mis Negocios y Gestión): fila **"Bono captación: +2 % al porcentaje de comisión de cada asesor"** después de "% oficina", solo si aplica.
- Mis Negocios: "Mi comisión: $X *(incluye bono captación +2 %)*".
- Panel expandible: texto unificado a "+2 % adicional al porcentaje de comisión de cada asesor (inmueble captado y arrendado en el mismo mes o el siguiente)".

### 4.4 Infraestructura
- `negNum(v)`: parser numérico compartido de los helpers de tarjeta (acepta `"1.234,56"`), mismo patrón que `nvM` / `numVal` / `nv` locales.
- Sin CSS nuevo: la grilla `neg-detail` de dos columnas admite las filas extra.

## 5. Valores validados con tests en Node

Se extrajo el `<script>` de `asesor.html`, se pasó `node --check` (sintaxis OK) y se ejecutaron los helpers con datos simulados:

| Caso | Resultado |
|---|---|
| Arriendo colegaje bruta 120.000, `pct_colega` 0,3, neta 84.000 | "Comisión total mensual (bruta): $120.000 · Colegaje 30 %: − $36.000 · REDA3 neta mensual: $84.000" |
| Arriendo normal (`comision_oficina` como texto "120.000,00") | Solo "Comisión mensual oficina: $120.000"; sin filas de colegaje ni bono |
| Venta colegaje Inmobiliaria 9.000.000 / 4.500.000 | "Colegaje 50 % · Inmobiliaria: − $4.500.000 · REDA3 neta: $4.500.000" |
| Venta colegaje Persona natural 9.000.000 / 6.000.000 | "Colegaje 33.33 % · Persona natural: − $3.000.000 · REDA3 neta: $6.000.000" |
| `bono_captacion_pct` 0,02 | Fila "Bono captación: +2 %"; con 0 no aparece nada |
| Cobros en Mar, Ene, Dic, Feb, May 2026 y Ene 2027 (desordenados) | "Periodo de cobros: Ene 2026 → Ene 2027", sin ninguna fila mensual |
| Un solo cobro (Ago 2026) | "Periodo de cobros: Ago 2026" |

## 6. Despliegue (pendiente — manual)

1. Pegar `asesor.html` en el archivo `asesor` de Apps Script → Implementar → Administrar implementaciones → ✏️ → **Nueva versión**.
2. Como el colegaje/bono del 27-ago (`e7f3521`) tampoco se ha desplegado: pegar `apps_script.js` en `Code.gs` → guardar → Ejecutar `prepararColumnasNuevas` (ver `NOTAS_SESION_2026-08-27.md` §Despliegue). Sin esto las tarjetas no tendrán datos de colegaje ni de bono que mostrar.
3. Pedir a los asesores refrescar (Ctrl+F5).

## 7. Pendientes

- Despliegue anterior (punto 6).
- Eliminar la venta duplicada VNT-086 desde Gestión y revisar otras ventas primarias repetidas de agosto (pendiente desde el 27-ago).
- Revisar en el navegador real un arriendo de 12 meses colegaje y uno con bono para confirmar la vista (los tests fueron en Node con datos simulados).
