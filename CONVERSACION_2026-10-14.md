# Conversación de trabajo — 14 octubre 2026

Registro de la sesión con Claude Code: pedidos, diagnóstico, decisiones, cambios y verificación.
Empezó como "no puedo editar los hitos pendientes de una venta" y acabó destapando que la
edición de negocios podía dejar una venta **sin asesores**. Todo el código está en `apps_script.js`
(canal oficial: pegar en el proyecto de producción de Apps Script).

**Hilo conductor:** al probar el primer arreglo sobre VNT-051, el guardado falló con el viejo error
de borrado (*"No se pudieron borrar todas las filas de Pagos…"*), y al mirar la hoja el usuario vio
que la venta había perdido sus comisiones. Las tres cosas estaban conectadas.

---

## 1. Hitos pendientes bloqueados en ventas con un mes ya liquidado

### Pedido
> "Después de que se cobra una comisión, por ejemplo de venta con más de 1 hito, los demás hitos
> así no se hayan cobrado no se pueden editar. Corrige eso."

### Diagnóstico
`editar_venta` llamaba `negocioBloqueadoPorCobro_`, que rechazaba **toda** la edición si
**cualquier** hito caía en un mes donde algún asesor del negocio ya tenía la bonificación cobrada
(`Bonificaciones_mes.cobrada_en`). Liquidar el mes del hito 1 congelaba también los hitos 2, 3…

Lo que de verdad hay que proteger es lo que cada hito liquidado aportó a cada asesor:
`valor_comision(asesor) × valor_cobrado(hito) / comision_oficina`, en el `mes_pago` del hito.

### Cambio
- Nueva `validarEdicionVentaLiquidada_` (+ helper `periodosCobradosAsesores_`). Sustituye el
  bloqueo total en `editar_venta`. Si la venta tiene hitos en meses liquidados, exige:
  1. misma comisión de oficina (congela valor base, % y colegaje);
  2. misma comisión por asesor (no se agregan/quitan asesores ni cambian participaciones);
  3. cada hito liquidado sigue existiendo con el mismo mes/año y `valor_cobrado` (emparejado por
     valor, no por posición);
  4. ningún hito nuevo o movido cae en un mes liquidado.
- Mensajes concretos, p. ej. *"El hito de jul-2026 ($300) ya está en una bonificación cobrada: no
  se puede cambiar su fecha ni su %/valor…"*.
- Las comisiones previas se cuentan **sin repetir** por (asesor, punta), igual que
  `filasComisiones_`: una hoja con comisiones duplicadas (VNT-051 las tenía) no debe leerse como
  "la comisión cambió".
- Sin cambios en `editar_arriendo` ni `eliminar_negocio`: un arriendo vive en un solo mes y borrar
  sigue exigiendo que nada esté cobrado.

---

## 2. El borrado de filas volvió a fallar (ahora en Pagos)

### Reporte del usuario
Al editar VNT-051: *"No se pudieron borrar todas las filas de "Pagos" con id_venta="VNT-051":
quedan 2 (filas 97, 98)…"*.

### Diagnóstico
Es el mismo fallo intermitente de `deleteRows` perseguido desde el 29-ago (antes en Comisiones),
nunca explicado del todo.

### Cambio
- `borrarFilasPorColumna_` ya **no usa `deleteRows`**: llama a la nueva `borrarPorReescritura_`,
  que lee el tramo desde la primera fila del negocio hasta el final, lo reescribe con `setValues`
  sin esas filas (escribe también en filas ocultas por filtro) y limpia con `clearContent` las
  filas sobrantes del final. Aborta si el tramo tiene fórmulas.
- Se eliminó `confirmarBloqueBorrable_` (ya no hay borrado por posición).

---

## 3. VNT-051 quedó sin asesores

### Reporte del usuario
> "Veo que se borró el negocio de la hoja Comisiones, entonces no quedó relacionado a ningún
> asesor? Qué pasó y puede pasar con otro negocio?"

### Diagnóstico
El esquema del 31-ago ("borrar todo primero, escribir después") **no era atómico entre hojas**.
En VNT-051: Partes → borradas, Comisiones → borradas, Pagos → falló → excepción. "No se insertó
nada nuevo" era cierto, pero lo ya borrado no volvía. Puede haber pasado con cualquier negocio
cuya edición falló con *"No se pudieron borrar…"* desde el 31-ago:
- falló en Comisiones → perdió las partes;
- falló en Pagos/CobrosArriendo → perdió partes y comisiones.

### Cambio
- `respaldarNegocio_` / `restaurarNegocio_` / `conRespaldo_`: `editar_venta` y `editar_arriendo`
  envuelven todos sus borrados y escrituras en `conRespaldo_`. Antes de tocar nada se respaldan
  la fila cabecera (restaurada en su misma posición), partes, comisiones y pagos/cobros; si algo
  lanza, se restaura todo y el error termina con *"El negocio se restauró tal como estaba; no se
  perdió nada."* (o avisa si la restauración misma falló).
- Nueva `verNegociosIncompletos()` (solo lectura, en el menú Ejecutar): lista ventas y arriendos
  vigentes sin comisiones, sin partes o (ventas) sin pagos.

### Recuperación
Recomendación: recuperar las filas perdidas de Comisiones y Partes desde **Archivo → Historial de
versiones** de la hoja, no re-editando el negocio (con hitos liquidados, la validación del punto 1
rechazaría volver a agregar los asesores). El usuario pegó el código nuevo y reportó que **VNT-051
ya guardó**.

### Pendiente
- Ejecutar `verNegociosIncompletos()` para ver si otros negocios quedaron dañados y recuperarlos
  del historial de versiones.

---

## Verificación
- `node --check apps_script.js` sin errores.
- Pruebas en Node con hojas simuladas:
  - validación de hitos: permite editar hitos pendientes; rechaza cambiar el hito liquidado,
    mover un hito a un mes liquidado, cambiar el valor base y cambiar comisiones de asesores;
  - borrado por reescritura: quita las 2 filas de VNT-051 y deja intactas las de otras ventas;
  - respaldo: simulando un fallo a mitad, Comisiones y Pagos vuelven a tener las filas originales.
- En producción: VNT-051 guardó correctamente tras el despliegue.
