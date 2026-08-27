# Notas de sesión — 27 agosto 2026

## Temas: colegaje · edición a prueba de duplicados · cancelados ocultos · bono de captación

### 0. Diagnóstico de la duplicación al editar (LEER PRIMERO)

La directora reporta que al editar una venta primaria se duplica. El fix de ese bug está en el repo desde el 6-ago (`e10fa91`, `9f0e44b`) y **en el código de HEAD no existe ninguna ruta que duplique**: `editar_venta` solo hace `actualizarFila`; `agregarFila` sobre Ventas existe únicamente en `registrar_venta`. `NOTAS_SESION_2026-08-06.md` dejó como **pendiente #1 pegar ambos archivos en Apps Script**, y el usuario cree que no se hizo. Hipótesis principal: **producción sigue con el código anterior al fix**.

Verificación en 1 minuto, en el editor de Apps Script:
- `asesor` (HTML) debe contener `El modo edición se activa ANTES de prefillar` y (tras este despliegue) `origen_prefill`.
- `Code.gs` debe contener `posible_duplicado` y (tras este despliegue) `rechazarRegistroDesdeEdicion_`.
- Implementar → Administrar implementaciones: la versión activa debe ser la publicada hoy.

### 1. Regla de gerencia para TODOS los negocios (arriendos y ventas)

**Editar = sobreescribe** (misma fila, nunca crea) · **Eliminar = borra del todo** (negocio + pagos/cobros + comisiones + partes) · **Cancelar = simplemente no aparece**.

Blindaje añadido (independiente del fix de agosto):
- Cliente: `formPrefillId` (id con el que se prefilló el formulario; lo fijan `prefillArriendo/prefillVenta`, lo limpian `resetArriendo/resetVenta`). `edicionCoherente()` aborta el guardado si hay prefill pero el modo edición no coincide. Se envía `datos.origen_prefill` siempre.
- Servidor: `rechazarRegistroDesdeEdicion_` en `registrar_arriendo`/`registrar_venta` rechaza (sin opción de confirmar) cualquier registro que traiga `id_venta`/`id_arriendo` o un `origen_prefill` que exista en la hoja.
- Tras un error al guardar el botón conserva "Guardar cambios (ID)"; el `confirm()` de posible duplicado avisa "si estabas EDITANDO…"; `apiPost` ya no trata una respuesta no-JSON como éxito; `console.log` en `editarNegocio` y en `guardarArriendo/guardarVenta` (abrir consola F12 para diagnosticar).
- `editar_arriendo`: el chequeo de duplicado excluye CANCELADOS; si no se regeneran los cobros devuelve `advertencia` (antes se silenciaba).

Cancelados ocultos:
- `mis_negocios` (servidor) ya no devuelve arriendos CANCELADO ni ventas CANCELADA (ni sus pagos/cobros).
- Gestión: se reemplazó el select "Estado venta" por la casilla **"Ver cancelados"** (aplica a arriendos y ventas). Con la casilla activa, los cancelados solo muestran **"Eliminar del todo"**.
- Dashboard: `depurarCancelados()` limpia el payload UNA vez en `_dashFetchData` (negocios cancelados, pagos/cobros ANULADO, comisiones ANULADA). **Antes los KPIs sumaban negocios cancelados** — bug corregido.
- Bonificaciones y PDFs ya excluían cancelados; sin cambios.

### 2. Bono de captación (+2 pp / +1 pp) — solo arriendos, ambas puntas

Si el inmueble se capta y se arrienda el **mismo mes**: captador y cerrador suman **+2 puntos** a su % (15% → 17%, 17,5% → 19,5%). Si se arrienda el **mes siguiente**: +1 punto. Después: nada. Cálculo por (año×12+mes) del arriendo menos el de la captación.

La fecha de captación **no se digita**: la sella la **acción comercial de tipo "Captación…"** ("Captacion en venta", "Captacion en arriendo", "Captacion compartida"), que ahora exige **inmueble** y **observación** (cliente y servidor). `registrar_accion` → `recalcularCaptacionInmueble_` escribe en `Inmuebles`: `fecha_captacion`, `id_asesor_captador`, `id_accion_captacion` (manda la captación más reciente). Eliminar la acción recalcula (vuelve a la anterior o vacía).

Columnas nuevas: `Acciones.id_inmueble`; `Inmuebles.fecha_captacion/id_asesor_captador/id_accion_captacion`; `Arriendos.bono_captacion_pct` (0.02 / 0.01 / 0, lo deriva el servidor al registrar/editar). Se muestra en el resumen del formulario, badge "BONO CAPTACIÓN +2" en tarjetas, panel de detalle y PDF de cuenta de cobro.

### 3. Colegaje (negocio compartido con otra inmobiliaria/agente)

| Caso | Colega | REDA3 | Base comisión asesores |
|---|---|---|---|
| Arriendo | 30% | 70% | (canon + admón) × factorArr × 0,70 (8% → 0,8 × 0,7 = 0,56) |
| Venta, colega **Inmobiliaria** | 50% del % total | 50% | comisión REDA3 neta (3% → 1,5%) |
| Venta, colega **Persona natural** | 1/3 del % total | 2/3 | comisión REDA3 neta (3% → 2%) |

- El asesor escoge el **% total** pactado; el sistema reparte.
- **`comision_oficina` guarda la parte NETA de REDA3** → bonificaciones (`neta × 50% × participación`, fórmula sin cambios), plan de pagos, cobros mensuales, PDFs y Dashboard quedan sobre la neta sin tocar su código. Se guardan además `comision_bruta`, `pct_colega` (lo deriva el servidor), `modalidad`, `tipo_colega`, `nombre_colega`.
- En colegaje normalmente solo se diligencia la punta que hizo REDA3; si se llenan ambas aparece un `confirm()`.
- Power BI (`Visual_CommandCenter.dax`) suma `comision_oficina` → los KPIs de ingreso pasan a neta automáticamente; `comision_bruta` disponible para vista bruta.
- Código: `calcBaseNegocio(tipo)` en `asesor.html` reemplaza las 3 copias de la fórmula (`recalcComisiones`, `recolectarComisiones`, `calcRefPuntaParaGuardar`). Servidor: `normalizarColegaje_`, `aplicarColegaje_`, `factorColegaDe_`, `heredarColegajeSiFalta_` (cliente viejo editando un colegaje no lo convierte en normal).
- Fix colateral: `prefillPagosForm` y `editarPago` reconstruían el monto de cada cuota como `valor_cobrado / % oficina`; ahora `valor_cobrado / comisión × valor base` (con neta daba montos errados).
- `agregarFila` ahora escribe por los encabezados físicos de la hoja (antes era posicional por `COLUMNAS`).

### Despliegue (pendiente — manual, ver NOTAS_SESION_2026-06-05.md)

1. Pegar `apps_script.js` en `Code.gs` → guardar → menú Ejecutar → `prepararColumnasNuevas` (crea las columnas nuevas y limpia el caché; las funciones con `_` final no aparecen en el desplegable). Revisar en Ejecuciones/Registro que diga "OK".
2. Pegar `asesor.html` en el archivo `asesor` → Implementar → Administrar implementaciones → ✏️ → **Nueva versión** (no nueva implementación).
3. Pedir a los asesores refrescar la página (Ctrl+F5).
4. Después: eliminar la venta duplicada VNT-086 desde Gestión y revisar otras ventas primarias repetidas de agosto.

### Casos de prueba (valores esperados)

| Caso | Esperado |
|---|---|
| Editar venta primaria (2 hitos) y guardar | misma fila VNT-xxx, toast "Venta actualizada", sin duplicado |
| Forzar `registrar_venta` con `origen_prefill: 'VNT-xxx'` | error del servidor, sin fila nueva |
| Cancelar venta/arriendo | desaparece de Mis Negocios, Dashboard y Gestión; con "Ver cancelados" aparece solo con "Eliminar del todo" |
| Captación INM-X (ago-2026) + arriendo ago-2026 canon 1,2M/10%/12 m freelance | ambos asesores al 17% → **204.000**; `bono_captacion_pct` 0.02; sep → 192.000; oct → 180.000 |
| Arriendo colegaje 1.000.000 + 200.000, 10%, 12 m, freelance captador 100% | bruta 120.000/mes, colega 36.000, REDA3 **84.000**; base 840.000; asesor **126.000**; CobrosArriendo 12 × 84.000 |
| Venta colegaje Inmobiliaria 300M × 3%, cerrador freelance | neta **4.500.000**, asesor **900.000**, `comision_por_punta` 2.250.000 |
| Venta colegaje Persona natural, 2 cuotas 150M/150M | neta **6.000.000**, asesor **1.200.000**, cuotas 3.000.000 c/u; editar cuota muestra monto 150.000.000 |
| Regresión normal (arriendo 1,2M/10%/12 m; venta 300M/3%) | 120.000 / 180.000; 9.000.000 / 1.800.000 (idénticos a hoy) |

Las fórmulas se validaron con tests en Node (`calcBaseNegocio`, `calcComisionAsesor`, helpers del servidor): todos los casos anteriores dan el valor esperado.
