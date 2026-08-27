# Conversación de trabajo — 27 agosto 2026

Registro de la sesión con Claude Code: pedidos, decisiones de negocio, cambios implementados y cómo funcionan. Complementa `NOTAS_SESION_2026-08-27.md` (que tiene los casos de prueba y el despliegue).

---

## 1. Pedidos del usuario (en orden)

1. **Colegaje**: modalidad de negocio compartido con otra inmobiliaria/agente que cambia la base de comisiones.
   - Arriendo: (canon + admón) × 70% es la base para las comisiones de los asesores.
   - Venta: si el colega es inmobiliaria (persona jurídica) se paga 1,5% y 1,5%, y todas las comisiones se calculan sobre esa mitad; si es persona natural se le paga 1% y REDA3 se queda con 2% de los 3%.
2. **Duplicación al editar**: "estoy dándole editar a una venta primaria y cuando termina no se sobreescribe, se duplican; pensé que ya lo habíamos solucionado. Si edita se reescribe, si elimina se elimina del todo y cancelar simplemente no aparece. Para absolutamente toda la lógica de los inmuebles."
3. **Bono de captación**: "si un asesor capta un inmueble y se alquila ese mes se suma 2% a su comisión; si se alquila el segundo mes 1%; a partir de terminar el segundo ya queda igual."
4. Tras el despliegue: no aparecían `asegurarColumnasColegaje_` ni `asegurarColumnasCaptacion_` en el editor; qué hacer con `limpiarCacheCatalogos`.
5. Explicación item por item de lo cambiado.

## 2. Preguntas de Claude y decisiones del usuario

| Pregunta | Decisión |
|---|---|
| ¿El recorte del colegaje aplica también a la comisión de oficina guardada (bonificaciones, plan de pagos, cobros, PDFs, dashboard)? | **Sí, todo sobre la parte de REDA3** (`comision_oficina` = neta; bruta guardada aparte). |
| En venta colegaje, ¿el % total es siempre 3% o se escoge? | **Se escoge el % total y se aplica la proporción** (inmobiliaria 50%, persona natural 1/3). |
| ¿Cómo se diligencian las puntas en colegaje? | **Solo la punta que hizo REDA3**; se advierte (no se bloquea) si se llenan ambas. |
| En arriendo colegaje con % de oficina distinto de 10%, ¿cómo se combina con el 70%? | **Se multiplican**: base = (canon+admón) × (%×10) × 70%. |
| ¿Cómo aporta un colegaje a la bonificación mensual? | **Fórmula actual sin cambios**: neta × 50% × participación. |
| ¿Se pegaron los archivos en Apps Script tras el 6 de agosto? | "No estoy seguro / creo que no" → hipótesis principal del duplicado: producción con código viejo. |
| ¿Cómo se suma el +2%? | **Puntos sobre el % del asesor** (15% → 17%), al captador **y también al cerrador**. Solo arriendos. |
| ¿De dónde sale la fecha de captación? | **De la acción comercial**: al registrar una captación empieza a contar; el campo se pone automático; la observación es obligatoria. |
| ¿Dónde se ocultan los cancelados? | **En todas partes; Gestión con casilla "Ver cancelados"**. |

## 3. Hallazgos de la investigación

- En el código de HEAD (`9f0e44b`) **no existe ninguna ruta que duplique al editar**: `editar_venta` solo hace `actualizarFila`; `agregarFila` sobre Ventas existe solo en `registrar_venta`. `NOTAS_SESION_2026-08-06.md` dejó como pendiente #1 pegar los archivos en Apps Script.
- La comisión del asesor se calculaba en el navegador con **tres copias** de la misma fórmula (`recalcComisiones`, `recolectarComisiones`, `calcRefPuntaParaGuardar`).
- No existía fecha de captación ni asesor captador a nivel de inmueble; las acciones "Captación…" no tenían inmueble asociado.
- El Dashboard **sumaba negocios cancelados y pagos anulados** en los KPIs (bug no detectado).
- `prefillPagosForm` y `editarPago` reconstruían montos de cuotas como `valor_cobrado ÷ % oficina`, lo que fallaría con comisión neta.
- `agregarFila` escribía por posición (`COLUMNAS`), no por encabezados del Sheet.
- Apps Script oculta del menú Ejecutar las funciones terminadas en `_`.

## 4. Qué se cambió y cómo funciona (explicación item por item)

### 4.1 Edición a prueba de duplicados (Editar = sobreescribe)

- **`formPrefillId` + `edicionCoherente()`** (asesor.html): segunda memoria independiente de `editandoNeg`. `prefillVenta`/`prefillArriendo` anotan el id con el que se llenó el formulario; solo lo borran `resetVenta`/`resetArriendo`. Al guardar, si hay prefill pero el modo edición no coincide, **no se envía nada** y sale "Se perdió el modo edición de VNT-xxx…".
- **`rechazarRegistroDesdeEdicion_`** (apps_script.js): el formulario manda siempre `datos.origen_prefill`; `registrar_venta`/`registrar_arriendo` rechazan **sin opción de confirmar** cualquier registro que traiga `id_venta`/`id_arriendo` o un `origen_prefill` que exista en la hoja.
- Tras un error el botón conserva "✓ Guardar cambios (ID)"; el `confirm()` de posible duplicado avisa "si estabas EDITANDO…"; `apiPost` ya no toma una respuesta no-JSON como éxito; `console.log` en `editarNegocio` y `guardarVenta/guardarArriendo` para diagnosticar con F12.
- `editar_arriendo`: el chequeo de duplicado ignora cancelados; si no se regeneran los cobros devuelve `advertencia` (toast rojo) en vez de silenciarlo.

### 4.2 Cancelados (Cancelar = no aparece)

- `mis_negocios` (servidor) ya no devuelve cancelados ni sus pagos/cobros; Mis Negocios filtra de nuevo en cliente (`negCancelado`).
- Gestión: casilla **"Ver cancelados"** (arriendos y ventas) reemplaza al select "Estado venta"; los cancelados visibles solo tienen **"Eliminar del todo"**.
- Dashboard: `depurarCancelados(data)` limpia una sola vez al descargar (`_dashFetchData`): negocios cancelados, pagos/cobros ANULADO, comisiones ANULADA.
- Eliminar no cambió: `eliminar_negocio` borra negocio + pagos/cobros + comisiones + partes.

### 4.3 Bono de captación (+2 pp / +1 pp, solo arriendos, ambas puntas)

- Acción comercial tipo "Captación…": selector **Inmueble captado** y **Observación** obligatorios (`cambioTipoAccion`, validado también en `registrar_accion`). Se guarda `Acciones.id_inmueble`.
- `recalcularCaptacionInmueble_`: toma la captación más reciente del inmueble y escribe `Inmuebles.fecha_captacion / id_asesor_captador / id_accion_captacion`. Eliminar la acción recalcula.
- `bonoCaptacionDe_` (servidor) / `bonoCaptacionPP` (cliente): `diff = (año×12+mes arriendo) − (año×12+mes captación)`; 0 → +0,02; 1 → +0,01; otro → 0. `calcComisionAsesor` suma `bonoPP` al % en arriendos (15% → 17%, 17,5% → 19,5%). El servidor guarda `Arriendos.bono_captacion_pct`.
- Se ve en el resumen del formulario, badge azul "BONO CAPTACIÓN +2", panel de detalle y PDF de cuenta de cobro.

### 4.4 Colegaje

- Formularios: **Modalidad** (Normal/Colegaje) en arriendo y venta; en venta **Tipo de colega** obligatorio; nombre del colega opcional. El % de oficina que se escoge es el total pactado.
- **`calcBaseNegocio(tipo)`** unifica las tres copias de la fórmula: arriendo `base = (canon+admón) × factorArr × factorColega`; venta `base = comisión neta`.
- Servidor: `normalizarColegaje_` (valida, exige tipo en ventas, **deriva `pct_colega`**), `aplicarColegaje_` (`comision_bruta` = total, `comision_oficina` = neta), `factorColegaDe_`, `heredarColegajeSiFalta_` (cliente viejo no convierte un colegaje en normal). En modalidad Normal las columnas se guardan vacías para que al editar colegaje → normal se limpien.
- Al ser `comision_oficina` la neta, bonificaciones (fórmula sin cambios), plan de pagos, cobros, PDFs, Dashboard y Power BI quedan sobre la parte de REDA3 sin tocar su código.
- Se ve en el resumen (bruta / colega / REDA3 neta), badge ámbar "COLEGAJE", sección en el detalle, panel de referencia al editar, filas en el PDF de cuenta de cobro y `[colegaje]` en el PDF de bonificación.
- Fix colateral: montos de cuotas = `valor_cobrado ÷ comisión × valor base` (antes ÷ % oficina).

### 4.5 Infraestructura

- Columnas nuevas: Arriendos `bono_captacion_pct, modalidad, tipo_colega, nombre_colega, pct_colega, comision_bruta`; Ventas las 5 de colegaje; Inmuebles `fecha_captacion, id_asesor_captador, id_accion_captacion`; Acciones `id_inmueble`.
- `asegurarColumnas_` genérica e idempotente; **`prepararColumnasNuevas`** (pública, sin `_`) las crea, limpia caché y reporta "OK" o "FALTAN columnas".
- `agregarFila` escribe por encabezados físicos del Sheet.
- Negocios antiguos con columnas vacías = Normal y sin bono; sin migración.

## 5. Valores validados con tests en Node

| Caso | Resultado |
|---|---|
| Arriendo normal 1,2M / 10% / 12 m, freelance | 180.000 (sin cambios) |
| Mismo arriendo, inmueble captado el mismo mes | 204.000 (+2 pp); mes siguiente 192.000; después 180.000 |
| Arriendo colegaje 1.000.000 + 200.000 / 10% / 12 m | bruta 120.000, REDA3 84.000, base 840.000, asesor 126.000 |
| Arriendo colegaje 8% / 6 m | base 672.000, asesor 50.400, neta 67.200 |
| Venta normal 300M × 3% | 9.000.000 / asesor 1.800.000 (sin cambios) |
| Venta colegaje Inmobiliaria | neta 4.500.000 / asesor 900.000 |
| Venta colegaje Persona natural | neta 6.000.000 / asesor 1.200.000 |

## 6. Despliegue realizado

- Commits: `e7f3521` (cambios) y `c91a5b0` (`prepararColumnasNuevas`), en GitHub.
- El usuario pegó `apps_script.js` y `asesor.html` en el editor de Apps Script, ejecutó `prepararColumnasNuevas` y publicó **nueva versión** de la implementación existente.
- `limpiarCacheCatalogos` ya no hace falta aparte (la llama `prepararColumnasNuevas`); solo sirve sola si se editan datos a mano en el Sheet y se quieren ver de inmediato.

## 7. Pendientes

1. Eliminar la venta duplicada VNT-086 desde Gestión y revisar ventas primarias repetidas de agosto.
2. Probar los casos de `NOTAS_SESION_2026-08-27.md` en producción (editar venta primaria, cancelar, captación + arriendo, colegaje).
3. Pendientes anteriores: liquidar julio 2026 y eliminar `AJUSTE_CONTINUIDAD_2026_07`.
