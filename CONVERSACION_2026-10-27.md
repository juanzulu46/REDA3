# Conversación de trabajo — 27 octubre 2026

Registro de la sesión con Claude Code: pedido, diagnóstico, decisiones, cambios y verificación.
Cambios en `apps_script.js` (servidor) y `asesor.html` (cliente). Deploy manual: pegar ambos en el
proyecto de producción de Apps Script y publicar nueva versión; después ejecutar los atajos del
checklist de abajo desde el editor.

---

## 1. Pedido

> "Me llamaron por un tema: Germán quería editar un negocio VNT-076 (Maracay Casa 10 Lagos). En donde
> aparecen los negocios aparecen unas fechas, y cuando él le da editar aparecen otras fechas. El
> propietario pagó una parte de la última cuota y por eso Germán quería entrar a editar. Creo que
> Germán editó una que ya se había pagado. Revisa muy bien todo porque todos los meses me llaman a
> decirme algo sobre la edición."

## 2. Diagnóstico

### 2.1 Dos fuentes de verdad para la fecha de una cuota
- Las listas (Mis Negocios, Gestión, panel "Cuotas guardadas", modal de cuota) pintan **`mes_pago` /
  `año_pago`** ("Ago 2026"); nunca el día.
- El editor (`prefillPagosForm`) y el modal `editarPago` llenan el `<input type="date">` desde
  **`fecha_pago`**.
- El servidor escribe ambos a la vez (`construirPagosVenta_`, `actualizar_pago`): a la hora de guardar
  siempre son coherentes. Si difieren es porque la hoja se dañó después.

### 2.2 La hoja Pagos tenía `fecha_pago` corrida dos filas
En las filas 99-114 de Pagos (PAG-153, 154, 158, 159, 160, 172, 173, 174, 175) `fecha_pago` tenía el
valor del hito **dos filas más arriba**; las demás columnas (id, venta, mes/año, valor, observación,
marca de comisión pagada) estaban bien. Por eso VNT-076 mostraba *May/Ago/Dic 2026* en la lista y
*2026-08-01 / 2027-11-30 / 2026-05-04* en el editor.

**Cuándo y cómo (historial de versiones de la hoja):**
- 7-oct 17:25: todo alineado.
- 14-oct 17:42: primera ejecución de `borrarPorReescritura_` (introducida ese día para reemplazar el
  `deleteRows` que fallaba) al editar VNT-051, **con el filtro básico de Pagos activo** (solo mostraba
  VNT-051; filas 99-116 ocultas). Las filas visibles (97-98) quedaron completas; en las ocultas las
  columnas A,B,D..L subieron dos filas como debía ser, pero la columna C (`fecha_pago`) conservó el
  valor viejo. El mecanismo exacto del lado de Google no se confirmó; la correlación con las filas
  ocultas por el filtro sí.
- Pagos, Ventas y Comisiones seguían con filtros básicos activos.

**Fechas verdaderas recuperadas del historial (7-oct 17:25 y 14-oct 17:44):**

| id_pago | venta   | fecha_pago hoy | correcta   | mes/año guardado |
|---------|---------|----------------|------------|------------------|
| PAG-153 | VNT-086 | 2026-02-20     | 2026-08-01 | 8/2026 |
| PAG-154 | VNT-086 | 2027-07-16     | 2027-11-30 | 11/2027 |
| PAG-158 | VNT-076 | 2026-08-01     | 2026-05-04 | 5/2026 (comisión pagada) |
| PAG-159 | VNT-076 | 2027-11-30     | 2026-08-28 | 8/2026 |
| PAG-160 | VNT-076 | 2026-05-04     | 2026-12-28 | 12/2026 |
| PAG-172 | VNT-074 | 2025-05-30     | 2026-04-24 | 4/2026 (comisión pagada) |
| PAG-173 | VNT-074 | 2027-11-30     | 2026-10-10 | 10/2026 |
| PAG-174 | VNT-093 | 2026-04-24     | 2026-10-01 | 10/2026 |
| PAG-175 | VNT-093 | 2026-10-10     | 2026-11-30 | 11/2026 |

VNT-075 y VNT-092 también quedaron corridas el 14-oct, pero se re-editaron el 16-oct y sus filas
nuevas (PAG-178..184) están bien.

### 2.3 ¿Germán editó una cuota ya pagada?
No llegó a guardar nada: los pagos de VNT-076 siguen siendo PAG-158..160 con los valores del 14-oct
(un guardado habría creado ids ≥185) y las versiones del 26-oct y 27-oct no tocan Pagos. Hito 1
(5/2026) tiene la comisión pagada; hito 2 (8/2026) está en mes liquidado para ASE-003 y lo protege
`validarEdicionVentaLiquidada_`; hito 3 (12/2026, el del pago parcial) es editable.

**¿Quién hizo la edición del 14-oct que corrió la columna?** El propio usuario, durante la sesión en
que se arregló VNT-051: 17:22 código nuevo pegado, 17:28-17:36 intentos fallidos, 17:42/17:44 el
guardado que funcionó ("VNT-051 ya guardó"), con la hoja Pagos filtrada a VNT-051 (puesta para mirar
las filas 97-98 del error). El historial atribuye todo a la cuenta del propietario porque la app corre
con ella. No es un error de uso: el código debía detectar o quitar el filtro.

### 2.4 Por qué "todos los meses llaman por la edición"
1. El formulario no enviaba `id_pago`; el servidor borraba todos los pagos y los reinsertaba copiando
   las marcas de "comisión pagada" **por posición**. Partir la última cuota en dos (el caso del pago
   parcial) desplazaba o perdía marcas; `estado` no se conservaba.
2. El editor no mostraba qué hitos tenían la comisión pagada, no bloqueaba nada y recalculaba el valor
   del último hito.
3. `borrarPorReescritura_` no verificaba lo que escribía (solo contaba filas del id) ni miraba filtros.
4. Dashboard y tres chequeos del servidor parseaban `new Date('YYYY-MM-DD')` (UTC → día anterior en
   Colombia: una cuota del día 1 caía al mes anterior); `aFechaInput` leía 'dd/mm/yyyy' al estilo US.

## 3. Decisiones
| Tema | Decisión |
|---|---|
| Qué escribir al reparar | Texto `'YYYY-MM-DD'` (igual que la app). |
| Filtro activo al reescribir | Se **quita** automáticamente y se avisa; además se verifica celda a celda lo escrito. Lanzar rompería también la restauración de `conRespaldo_`. |
| Fecha vs mes/año en el editor | **Manda mes/año**: si discrepan, el editor propone el día 1 del mes guardado y avisa; guardar deja ambos coherentes. |
| Emparejar hitos al editar | Por `id_pago` si el formulario lo manda; por posición solo con un cliente viejo en caché. |
| Ids de hitos | Un hito emparejado **conserva su id_pago**; los nuevos reciben `siguienteId`. |
| Hito con comisión pagada | No se puede quitar ni cambiar de mes/valor (servidor y cliente); sí la observación y el día dentro del mes. |

## 4. Cambios

### 4.1 `apps_script.js`
- **Reparación (menú Ejecutar):** `verFechasPagoCorridas()` (informe: celdas Date/texto, filtro, filas
  cuyo mes de `fecha_pago` ≠ `mes_pago/año_pago`, plan sobre los 9 ids) y
  `aplicarFechasPagoCorridas()` (escribe solo la celda `fecha_pago` de los ids de
  `FECHAS_PAGO_RECUPERADAS_20261014` que sigan corridos; omite si el mes guardado no coincide o si ya
  fue corregida a mano; rehúsa con filtro activo; idempotente). `verFiltrosActivos()` /
  `quitarFiltrosHojas()`.
- **`borrarPorReescritura_`:** quita el filtro básico antes de escribir (`quitarFiltroSiHay_`, con
  aviso al usuario vía `AVISOS_OP_`/`conAvisos_`), escribe el tramo en una sola llamada y **relee y
  compara celda a celda** (`normCelda_`); si algo difiere lanza (dentro de `conRespaldo_` eso
  restaura).
- **`editar_venta`:** `construirPagosVenta_` arrastra `id_pago_prev`; `emparejarPagosPrev_` (modo id /
  posición) y `validarHitosPagados_` (hito pagado no se elimina ni cambia de mes/valor); el `preparar`
  conserva `id_pago`, `estado` y marcas del hito emparejado; advertencia explícita cuando cambia el
  número de cuotas. `actualizar_pago` aplica la misma regla.
- **Fechas:** `anioMesDe_` devuelve también `dia`; nueva `fechaLocalDe_` (medianoche local) reemplaza
  `new Date(p.fecha_pago)` en `pagoYaIngresado_`, `pagosDelMes` (liquidación), `p.efectuado`
  (mis_negocios) y la cuenta de cobro.
- `periodoPago_` y `nombreHitoPago_` a nivel de módulo (antes locales de
  `validarEdicionVentaLiquidada_`).

### 4.2 `asesor.html`
- `addPagoRow(opts)`: muestra "guardado: Ago 2026" (mes_pago/año_pago) y, si la comisión está pagada,
  "✓ comisión pagada" con fecha y monto/% de solo lectura y sin botón ×. Guarda `id_pago`, mes/año,
  estado y marca en `dataset`.
- `prefillPagosForm`: pasa esos datos a cada fila; el ajuste de cuadre va al **último hito no
  bloqueado**; la fecha sale de `fechaInputDeHito` (si `fecha_pago` no cae en el mes guardado,
  propone el día 1 y marca la fila en ámbar con aviso).
- `recolectarPagos`: envía `id_pago`. `cambioMercadoVenta` conserva id/mes/marca al reconstruir.
- `editarPago` (modal): "Guardado como Ago 2026", misma regla del día 1, hito pagado con valor/% de
  solo lectura y validación del mes antes de enviar.
- `aFechaInput`: rama explícita para 'dd/mm/yyyy' y getters locales (sin `toISOString`).
  `fechaLocalDeTexto` y `_pagoAnioMes` reemplazan `new Date(p.fecha_pago)` en el dashboard (pipeline,
  cobrado, gráfica 12 meses, embudo); el mes de la cuota es `mes_pago/año_pago`. La fecha por defecto
  de la acción comercial también es local.

## 5. Verificación
- `node --check apps_script.js` y del `<script>` de asesor.html: sin errores.
- Arnés en Node (hojas simuladas, no versionado): 19 pruebas unitarias (fechas, `normCelda_`,
  reescritura con filtro, reescritura con el bug simulado → la verificación lanza nombrando
  `fecha_pago`, emparejamiento por id/posición, hitos pagados, reparación `ver`/`aplicar`
  idempotente, filtros), 5 de helpers del cliente y 7 de extremo a extremo sobre `doPost`:
  - partir la última cuota de VNT-076 conserva PAG-158..160 con la marca de PAG-158 y crea PAG-161;
  - quitar o cambiar de mes/valor el hito pagado → rechazo sin tocar la hoja; observación/día → ok;
  - cliente viejo sin `id_pago` → ok por posición;
  - filtro en Pagos → se quita, guarda y avisa;
  - `actualizar_pago` aplica la misma regla.

## 6. Deploy y checklist en producción (en este orden)
1. Pegar `apps_script.js` en `Código.gs` y `asesor.html` en `asesor`; Implementar → Administrar
   implementaciones → ✏️ de la implementación en uso → Nueva versión.
2. En el editor, menú Ejecutar: `verFiltrosActivos()` → `quitarFiltrosHojas()` → `verFiltrosActivos()`
   (ninguna).
3. `verFechasPagoCorridas()`: deben salir exactamente las 9 filas de la tabla. Si aparecen otras,
   parar y analizar (el `aplicar` solo toca las 9).
4. `aplicarFechasPagoCorridas()` → "escritas 9"; repetir `ver…` → 0 inconsistentes.
5. Ctrl+F5 en la app. VNT-076: lista May/Ago/Dic 2026; Editar muestra 2026-05-04 / 2026-08-28 /
   2026-12-28 con "guardado: …" y el hito 1 bloqueado con ✓. Cancelar sin guardar.
6. Registrar el pago parcial real de VNT-076: en Editar, fila 3 = fecha y monto del abono, "+ cuota"
   con el saldo en 2026-12-28, Guardar. Debe conservar PAG-158 pagado y mostrar "PAGADO 1/4".
7. Editar un arriendo con cobros (debe seguir regenerando CobrosArriendo).

## 7. Al ejecutar en producción
- `verFiltrosActivos()` agotó los 6 minutos: contaba filas ocultas con `isRowHiddenByFilter` fila por
  fila (miles de llamadas en CobrosArriendo/Comisiones/Partes). Se quitó ese conteo (también se acotó
  en `diagnosticarBorrado` a hojas de ≤300 filas).
- `verFechasPagoCorridas()` listó **24** inconsistencias: las 9 del 14-oct y **15 hitos "Migrado desde
  Pipeline"** (PAG-032, 034, 044, 046-050, 053, 054, 059, 061-064) que ya estaban así en los respaldos
  del 31-jul y 26-ago: `fecha_pago` = día 1 del mes de la venta (relleno de la migración) y
  `mes_pago/año_pago` = mes esperado de pago. Como el mes que manda es `mes_pago/año_pago`, se
  agregaron `verFechasPagoAlMesGuardado()` / `aplicarFechasPagoAlMesGuardado()`, que ponen
  `fecha_pago` = día 1 del mes guardado en esas filas (excluyen las 9 del mapa, las sin fecha y las
  coherentes). También se vio que la hoja guarda `fecha_pago` como **Date** (115 Date · 0 texto):
  Google convierte el texto 'YYYY-MM-DD' al escribir; `normCelda_` y `fechaTexto_` ya lo contemplan.

## 8. Pendiente / siguiente paso sugerido
- Ejecutar `verFechasPagoCorridas()` una vez al mes como control.
- Más adelante: actualizar Pagos **en sitio** por `id_pago` (sin desplazar filas) en vez de
  borrar-y-reinsertar; con los ids estables ya queda el camino abierto.

## 9. Resultado en producción (27-oct, noche)
Checklist ejecutado por el usuario con el código final (commit 6377c20):
- `verFiltrosActivos` / `quitarFiltrosHojas`: filtros fuera.
- `aplicarFechasPagoCorridas`: escritas 9.
- `verFechasPagoAlMesGuardado` → 15 filas "Migrado desde Pipeline"; `aplicarFechasPagoAlMesGuardado`: escritas 15.
- `verFechasPagoCorridas`: 0 inconsistencias.
- Nueva versión publicada; el usuario confirmó **"funcionó"**: VNT-076 muestra las mismas fechas en
  lista y editor y Germán puede registrar el pago parcial.
