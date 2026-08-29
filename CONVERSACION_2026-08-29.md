# Conversación de trabajo — 29 agosto 2026

Registro de la sesión con Claude Code: pedidos, diagnóstico, decisiones, cambios y verificación
en producción. Commits en `master` (empujados a GitHub como respaldo; el canal oficial sigue
siendo pegar en Apps Script):

| Commit | Qué |
|---|---|
| `d6b80c0` | Fix del duplicado de asesores al editar + botón "Pagado" en Gestión |
| `01c59ed` | Comisión pagada **por hito** en ventas + filtro por **mes de pago** |
| `5b853dd` | `marcarPagadosHasta()`: marca masiva de comisiones anteriores a un corte |
| `2403816` | Atajos sin argumentos para el menú Ejecutar del editor de GAS |

---

## 1. Pedidos del usuario (en orden)

1. > "Está ocurriendo algo que te dije que ajustaras. Cuando edito un negocio y cambio algunas cosas, y le doy guardar se duplican los asesores que pertenecen a ese negocio, corrige eso"
2. > "Que la gerencia, el perfil de Germán Zuluaga, tenga un botón en la parte de gestión que diga Pagado, y marque la tarjeta en un verde sutil, con un filtro de pagado y no pagado"
3. > "En venta se pagan las comisiones por cada hito que pase… el gerente especifica cuál de los hitos ya se pagó" + "si un negocio alguien lo creó en junio, pero tiene una cuota en agosto y filtro agosto, no aparece la parte del negocio de agosto"
4. > "Necesito poner todos los negocios antes de julio de este año, tanto en ventas como en arriendo, pagados… si un inmueble tiene 3 hitos y dos de esos fueron antes de julio ponerlos pagados y el que está para el futuro sin marcar"

---

## 2. El duplicado de asesores: se cayó la hipótesis de tres sesiones

Desde el 6 de agosto veníamos asumiendo que el duplicado se debía a que **el fix nunca se pegó
en Apps Script**. Se verificó y **era falso**:

| Qué se verificó | Resultado |
|---|---|
| Proyecto de producción | "App Asesores" (modif. 28-ago) → hoja `data_set_final` (`1_jMuomNA4c…`). Hay **otro** proyecto con el mismo nombre (10-abr) apuntando a una hoja vieja: no tocarlo |
| Implementación activa | **Versión 83** (28-ago 12:42); todas las ejecuciones de los últimos 7 días corren ahí, sin errores |
| Código desplegado vs repo | `Código.gs` y `asesor` **idénticos byte a byte** a HEAD (comparados por hash por bloques de 100 líneas desde el editor) |

Es decir: **el duplicado lo producía el código vigente**.

### Causa raíz

`editar_arriendo` / `editar_venta` hacían *borrar y volver a insertar*, y
`borrarFilasPorColumna_` **fallaba en silencio**: devolvía `0` si no encontraba la columna
o si `deleteRow` no surtía efecto. **Nadie verificaba el resultado**, así que el código seguía
insertando las filas nuevas encima de las viejas.

Evidencia recogida en la hoja viva (Comisiones, 797 filas en ese momento):

- **VNT-076**: pareja vieja en f706/707 **+ pareja nueva en f797/798** (la intermedia, f749/750, sí se había borrado).
- **VNT-043** (f487/488 + f765/766), **VNT-051** (f503/504 + f763/764), **VNT-046** (ASE-003 Cerrador en f494 y f516).
- En la **misma edición** de VNT-076 los **pagos sí se borraron bien** (quedaron solo PAG-158/159/160) → el fallo era **intermitente**, no sistemático.
- También había comisiones huérfanas (VNT-052, VNT-060, VNT-065) y pagos huérfanos: el mismo
  borrado silencioso fallando en `eliminar_negocio`.

Los `id_negocio` estaban limpios (verificado byte a byte en el xlsx), así que la solución no fue
"arreglar el borrado" sino **dejar de confiar en él**.

> Hallazgo lateral: la hoja **Comisiones tiene un filtro activo** ("Se muestran 57 de 795 filas").
> No causa el bug, pero esconde el 93 % de las filas y hace que `gviz` y cualquier revisión
> manual engañen. **Sigue pendiente quitarlo.**

---

## 3. Decisiones del usuario

| Pregunta | Decisión |
|---|---|
| ¿Qué significa "Pagado"? | **Comisiones del negocio ya pagadas** a los asesores (no "la oficina recibió el dinero") |
| ¿Quién puede marcar? | **Solo el rol `gerente`**; la directora ve la marca y el filtro, pero no los botones |
| ¿El filtro Año/Mes cambia de significado? | **Selector "Filtrar por"**, con **Mes de pago** por defecto; se conserva "Mes del negocio" |
| ¿Dónde aparece un arriendo con el criterio "mes de pago"? | **Solo en el mes del arriendo** (su comisión se paga una vez, no en los 12 meses de cobros) |
| ¿Botón para marcar la venta completa? | **Sí**, "Marcar todos los hitos", además del botón por hito |

---

## 4. Qué se cambió

### 4.1 Blindaje contra el duplicado (`apps_script.js`)

- `borrarFilasPorColumna_`: compara encabezados y valores **normalizados** (trim), borra por
  **bloques contiguos** con `deleteRows` + `flush`, **re-lee para verificar**, reintenta una vez
  y **lanza excepción** si quedan filas. Nunca más un `0` en silencio.
- `reemplazarFilasDeNegocio_(hoja, col, valor, filas, columnas, preparar)`: borra → inserta →
  **verifica que queden exactamente las filas nuevas**. Es el único camino permitido para
  "editar = sobreescribe". Se usa en comisiones y pagos; `verificarFilasDeNegocio_` para partes.
- `filasComisiones_`: colapsa repetidos por `(id_asesor, punta)`, también en
  `registrar_arriendo` / `registrar_venta` → el duplicado es **estructuralmente imposible**.
- `actualizarFila` resuelve la columna con `indiceColumna_` (trim) y compara ids normalizados.
- `Logger.log` + campo `verificacion` en la respuesta → si vuelve a pasar queda rastro en *Ejecuciones*.
- `repararDuplicados()`: informe (o limpieza con `true`) de comisiones y pagos duplicados y de
  filas huérfanas de negocios inexistentes.
- Cliente: `prefillComisionesForm` de-duplica por `(asesor, punta)` y avisa por toast.

### 4.2 Marca "Pagado" (arriendos por negocio, ventas por hito)

- **Arriendos**: columnas `pagado` / `fecha_pagado` / `pagado_por`; una sola marca.
- **Ventas**: la marca fina vive en **Pagos** (`comision_pagada` / `fecha_comision_pagada` /
  `comision_pagada_por`). `Ventas.pagado` es solo el **resumen**: 'SI' cuando todos los hitos
  vivos están pagados (`recalcularPagadoVenta_`); los ANULADOS no cuentan.
- Acción `marcar_pagado` con `esGerente_` (excluye a la directora). Acepta `id_pago` (un hito)
  o sin él (todos los hitos / el arriendo).
- `editar_venta` regenera los pagos pero **conserva las marcas emparejando por posición**, y
  devuelve advertencia si cambió el número de cuotas.
- UI: botón por cuota dentro de "Cuotas / hitos", "Marcar todos los hitos" en la tarjeta,
  badge **PAGADO 1/3** ámbar + verde tenue para el estado parcial, verde pleno solo si están
  todos, y filtro **Todos / Pagados / Parciales / Sin pagar**.

### 4.3 Filtro por mes de pago

- Selector **"Filtrar por: Mes de pago (cuotas) / Mes del negocio"**, por defecto mes de pago.
  Una venta de junio con cuota en agosto **ya aparece al filtrar agosto**.
- La tarjeta muestra la franja azul **"En Ago 2026: Cuota 2 de 3 — $X · pendiente/pagada"** y el
  resumen agrega el stat **"Cuotas del periodo"**.
- Los arriendos siguen filtrándose por su propio mes.

### 4.4 Carga masiva de lo histórico

`marcarPagadosHasta(aplicar, anio, mes)`: marca todo lo **anterior** al corte — arriendos por su
mes, hitos por su `mes_pago` — ignorando cancelados, ANULADOS, sin fecha y lo ya marcado, y
recalculando el resumen de cada venta. Es **idempotente**. Escribe por bloques
(`escribirColumnasSiCumple_`, `recalcularPagadoTodasLasVentas_`) para no tardar minutos.

Como el desplegable de **Ejecutar** no admite argumentos, se agregaron cuatro atajos:
`verReparacionDuplicados`, `aplicarReparacionDuplicados`, `verMarcaPagadosHastaJulio2026`,
`aplicarMarcaPagadosHastaJulio2026`.

---

## 5. Pruebas

**72 tests en Node** sobre el código real (se extraen las funciones y se ejecutan contra hojas
simuladas), todos en verde:

- Borrado/reemplazo: duplicados viejos, encabezado con espacios, valores con espacios, columna
  inexistente, **borrado saboteado** (el bug real → lanza excepción y **no** inserta), negocio sin
  comisiones, `agregarFila` con encabezados en otro orden.
- Pago por hito: parcial vs completo, badge, permisos (directora sin botones), ANULADO que no
  cuenta, venta sin hitos, quitar un hito desmarca la venta.
- Filtro: venta de junio con cuota en agosto aparece en junio, agosto y diciembre, **no** en septiembre.
- Marca masiva: bordes del corte (junio sí, julio no), cancelados, sin fecha, idempotencia y el
  caso "2 de 3 hitos".

---

## 6. Verificación en producción (tras ejecutar todo)

| Verificación | Resultado |
|---|---|
| Hitos de venta marcados | **65**, todos anteriores a julio |
| Hitos marcados de julio en adelante | **0** ✅ |
| Hitos futuros sin marcar | 30 |
| Ventas completas | 55 de 79 · 0 canceladas marcadas |
| Arriendos marcados | 220 · ninguno anterior a julio quedó sin marcar |
| Marca | fecha `2026-08-29`, por `ASE-055` |
| Pagos | 116 → **109** filas (se limpiaron duplicados/huérfanos) |
| Implementación activa | **Versión 85** (17:47), posterior al último cambio de `asesor.html` (17:35) |

---

## 7. Pendientes

1. **Verificar que no quedaron comisiones duplicadas**: Comisiones bajó de 797 a 795 filas, menos
   de lo esperado (~13). Puede deberse a ediciones hechas entre medias, pero conviene correr
   `verReparacionDuplicados` (solo informa) y, si reporta duplicados, `aplicarReparacionDuplicados`.
2. **Revisar 4 arriendos de agosto marcados como pagados** (ARR-231, ARR-237, ARR-238, ARR-239):
   la función no pudo marcarlos por ser posteriores al corte, así que salieron de pruebas manuales.
   Si no van, quitarlos con "↩ Quitar pagado".
3. **Quitar el filtro de la hoja Comisiones** (hoy muestra 57 de 795 filas).
4. Publicar una versión nueva cuando se quiera dejar el desplegado idéntico al editor (lo que falta
   son solo funciones de mantenimiento; la app web ya está completa en la Versión 85).
5. Vienen de antes: liquidar julio 2026 y eliminar `AJUSTE_CONTINUIDAD_2026_07`.
