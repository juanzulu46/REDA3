# Notas de sesión — 31 julio / 3 agosto 2026

## Tema: corrección definitiva de bonificaciones + 6 ajustes de visibilidad

### Reglas de negocio DEFINITIVAS (validadas con gerencia)

1. **Arriendos** (sin cambios desde antes): suman a la bonificación en el **mes del negocio** como `comisión_oficina × meses_contrato × 50% por punta × participación`. La comisión de oficina se calcula sobre **canon + administración** (verificado: los 223 arriendos históricos están bien).
2. **Ventas — CAJA PURA**: cada pago suma a la bonificación del **mes en que ingresa** a la oficina, sin importar el mes del negocio. Si una venta se paga en 3 cuotas, cada cuota aporta a su propio mes. Tope: lo ingresado en un mes nunca supera la comisión pactada (protege de pagos duplicados, ej. VNT-060).
   - Ejemplo canónico VNT-076 (Maracay, $4.000M — valor confirmado como real): mayo $22,2M → mayo; junio $0; julio $31,5M → julio; diciembre $66,2M → diciembre.
   - Consecuencia aceptada: julio de Luisa Ledesma = $8.640.000 → BRONCE **$506.615** (su liquidación manual de $90.654 estaba incompleta: le faltaban las cuotas de Foresta y Nexo que ingresaron en julio).
3. **Ventas mercado Primario**: la constructora paga la comisión por hitos definidos como **% de la comisión REDA3** (primera parte / punto de equilibrio / escritura), cada hito con fecha estimada. Los % deben sumar 100%. Se guarda en la columna `pct_pago` de la hoja Pagos (se crea sola).
4. **% variable**: 4% primer mes en una categoría, 5% por continuidad (fuente: mes anterior cerrado en BonificacionesMes). Empleado ÷1.3.
5. **Fecha del PDF de cuenta de cobro de bonificación**: 1° del mes siguiente al mes liquidado.

### Ajuste manual TEMPORAL (julio 2026)

`AJUSTE_CONTINUIDAD_2026_07 = ['ASE-024','ASE-032','ASE-013']` (Mariana Rosero, Sandra Zapata, Jeisson Naranjo) → liquidan julio con **5%** por continuidad desde junio (junio nunca se liquidó en el sistema; BonificacionesMes solo tiene mayo).
**⚠️ ELIMINAR esta constante de apps_script.js después de liquidar julio.**

### Valores esperados julio 2026 (verificados contra data_set_final_descarga.xlsx)

| Asesor | Categoría | % | Total |
|---|---|---|---|
| Luisa Ledesma (Empleado) | BRONCE ($8.640.000) | 4% | $506.615 |
| Mariana Rosero (Freelance) | PIEDRA ($4.008.000) | 5% | $356.900 |
| Sandra Zapata (Freelance) | PIEDRA ½ ($1.140.000) | 5% | $135.250 |
| Jeisson Naranjo (Empleado) | PISO ($5.933.750, 0 acciones registradas) | 5% | $228.221 |

### Cambios implementados (commits)

- `a9a6c4f` — Ventas por mes del negocio + ajuste 5% julio + fórmula unificada (`basesBonificacion_`) — *regla luego reemplazada por caja*.
- `6fc4b4a` — Fecha del PDF de bonificación = 1° del mes siguiente.
- `d713641` — Ventas por mes del negocio solo con lo ingresado — *regla luego reemplazada por caja*.
- `38d061c` — **Definitivo**: (1) caja pura en bonificaciones + PDF por ingreso del mes; (2) primario por hitos % (`pct_pago`, form dual, `construirPagosVenta_`, `actualizar_pago` con %); (3) Mis Negocios muestra ventas en los meses de sus cuotas con badge; (4) tabla "Negocios que componen tu bonificación" en Mis Bonificaciones (`detalle` en la respuesta); (5) panel de referencia al editar negocio + modal de edición de cuota con contexto; (6) card "Bonificaciones por asesor" en Gestión (action `bonificaciones_asesores`, gestor-only, liquidado o preliminar).

### Otras verificaciones de la sesión

- Comisión de arriendos SÍ incluye administración en código y datos; ARR-220 (Casa Centro) y ARR-221 (Casa Maraya) tienen admón registrada en $0 — verificar si es real.
- El ×10 en comisión de asesor de arriendos: base = canon total × (% oficina × 10) → si la oficina cobra 8% en vez de 10%, el asesor comisiona sobre el 80% (comparte el descuento). Reproduce el Excel histórico.
- VNT-076 Maracay: valor $4.000 millones confirmado como real por gerencia.

### Pendientes

1. **Pegar `apps_script.js` y `asesor.html` en el editor de Apps Script** (producción) si aún no se hizo.
2. **Liquidar julio** desde Gestión y luego **eliminar `AJUSTE_CONTINUIDAD_2026_07`**.
3. Jeisson no tiene acciones comerciales de julio registradas → queda en PISO sin fijo; si la oficina lo tiene en categoría, registrar sus acciones antes de liquidar.
4. Ventas primario antiguas (Foresta, Nexo…) conservan cuotas por monto; editarlas una vez desde Gestión si se quieren pasar a hitos %.
