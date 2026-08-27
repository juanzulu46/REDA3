# Notas de sesión — 6 agosto 2026

## Tema: edición completa para asesores en Mis Negocios + fix de negocios duplicados

### Qué se pidió

1. Que los asesores tengan en "Mis Negocios" los mismos botones de gestión que la gerencia tiene en "Gestión" (Editar, Cancelar y — pedido sobre la marcha — también Eliminar).
2. Urgente: al editar un negocio se estaba **duplicando** (reporte de la directora comercial, pantalla "Venta registrada VNT-086" con "Comisión: $ NaN").
3. Garantía de que no se vuelva a repetir un negocio.

### Causa raíz de la duplicación (bug desde 38d061c, hitos primario)

`recalcComisiones` en el detalle "Plan de pagos" leía `.pago-valor` en TODAS las filas, pero en mercado **Primario** las filas son hitos con `.pago-pct` → TypeError al final del prefill de edición → `editandoNeg` nunca se asignaba → el formulario quedaba lleno pero en modo **registro** → al guardar se creaba una venta nueva (VNT-086). Solo afectaba ventas primarias. El NaN de la pantalla era un bug hermano (mismo origen: leer `valor_pago` en hitos que solo tienen `pct_pago`).

### Cambios implementados (commits)

- `9485a5a` — **Mis Negocios: Editar y Cancelar para asesores** + fix del panel de éxito pegado al editar:
  - Flujo de edición parametrizado por fuente (`fuenteNegocio`: misNegociosData vs gestionData; `comisionesDeNegocio` usa `comisiones_negocio` con todas las puntas — usar `comisiones` de mis_negocios habría borrado las puntas de los compañeros al editar).
  - Backend: helper `puedeGestionarNegocio_` (gestor O asesor que figura en Comisiones del negocio, incluye ANULADAS) aplicado a `editar_arriendo`, `editar_venta`, `cancelar_venta`; `cancelar_arriendo` refactorizado al mismo helper.
  - El prefill solo siembra asesor + participación; los valores se recalculan (`recolectarComisiones`) — el enmascaramiento de comisiones ajenas no afecta el cálculo.
  - Año original del negocio viaja en `editandoNeg.anio`; `finalizarEdicion` vuelve a la vista de origen.
  - Botón "Editar" de cuota individual oculto para asesores (solo gerencia).
  - Incluía además trabajo previo del 03-08 sin commitear (comisiones_negocio, decorarDetalleNegocios_, doGet sin candado, detalle en bonificaciones_asesores).
- `e10fa91` — **Fix duplicación + Eliminar para asesores**:
  - `recalcComisiones` distingue Primario/Secundario en el plan de pagos, con guardas de nulos (también en `resumenPagos`, `recolectarPagos`, `prefillPagosForm`).
  - `editarNegocio` blindado: modo edición se activa ANTES del prefill; si el prefill falla, la edición se aborta completa (estructuralmente imposible duplicar por esta vía).
  - `eliminar_negocio` acepta gestor o participante; botón Eliminar en Mis Negocios (misma confirmación escrita ELIMINAR y degradación a cancelar si hay bonificaciones cobradas); `eliminarNegocio` refresca la vista de origen.
- `9f0e44b` — **Barrera anti-duplicados en `registrar_venta`**: el servidor rechaza registrar una venta si ya existe una activa del mismo inmueble en el mismo mes/año, salvo confirmación explícita (`confirmar_duplicado`) — necesaria porque en Primario un proyecto vende varias unidades el mismo mes. Arriendos ya tenían bloqueo duro.

### Cómo se guardan los pagos (aclarado a petición)

Misma hoja Pagos para ambos mercados (`construirPagosVenta_`):
- **Secundario**: asesor escribe montos del inmueble por cuota (deben sumar el valor base); se guarda `valor_cobrado = (monto/valor base) × comisión oficina`, `pct_pago` vacío.
- **Primario**: asesor escribe % de la comisión por hito (suman 100); se guarda `pct_pago` (fracción) y `valor_cobrado = % × comisión oficina`.
- Bonificación por caja lee `valor_cobrado` con `mes_pago/año_pago` — igual para ambos.
- Al editar, el formulario reconstruye montos/% desde lo guardado y el plan se regenera completo al guardar.

### Pendientes

1. **Pegar `apps_script.js` y `asesor.html` en el editor de Apps Script** (los 3 commits requieren ambos archivos; sin el backend los botones de asesores fallan).
2. **Eliminar el duplicado VNT-086** (venta primaria $395.600.000, 2 hitos) y revisar si hay otras ventas primarias repetidas (mismo inmueble+valor, distinto ID; la de ID más reciente es el duplicado).
3. Siguen vigentes los pendientes del 03-08: liquidar julio y eliminar `AJUSTE_CONTINUIDAD_2026_07`.
