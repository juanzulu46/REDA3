# Notas de sesión — 5 de junio de 2026

## Cambios realizados (commit `e15ee7d`, ya en GitHub y en Apps Script)

### 1. Cobros mensuales de arriendo ocultos en el frontend
- **Mis Negocios** y **Gestión**: se eliminó la sección "Cobros mensuales" de las tarjetas de arriendo. Solo se muestra `Canon: $X` (el primer canon).
- En Gestión también se eliminó el contador "cobrado / proyectado" del encabezado y el botón **Editar** de cada cobro (la función `editarCobro()` se eliminó por quedar inalcanzable).
- **El backend NO se tocó**: los cobros se siguen generando con `generarCobrosProyectados()` y guardando en la hoja `CobrosArriendo`. Siguen disponibles para informes/Power BI.
- ⚠️ Para corregir un cobro (estado COBRADO/NO_COBRADO, valor, observación) ahora se edita **directo en la hoja `CobrosArriendo`** del Google Sheet.

### 2. Desplegables: lo más reciente arriba
- Los catálogos **inmuebles, clientes, oficinas, orígenes y zonas** se invierten al cargarlos (`cargarCatalogos()` en `asesor.html`): el último registro de la hoja aparece de primero en todos los selects.
- Al crear un inmueble o cliente nuevo desde el formulario, se inserta al inicio de la lista (`unshift`).
- Sin cambios: meses, años, tipos de inmueble y asesores (orden calendario/fijo).

## Decisión de arquitectura: canal oficial = Apps Script directo

- **Los asesores usan el link directo de Apps Script** (guardar como marcador):
  `https://script.google.com/macros/s/AKfycbwvk6g9qhm_fqLbyjMspkvTf4MitW0gd0K-kvAU0KSmpYngxquq0XWV5mWS8EU8ATwk/exec`
- **Render queda como secundario** (no es producción). Motivos: dos copias de `asesor.html` causaban drift de versiones; Render free se duerme (~30-60 s de espera al despertar); la página no depende de nada del proxy (`server.js` solo agregaba un campo `efectuado` que el frontend no usa).
- `server.js` sigue siendo útil para pruebas locales: `node server.js` → http://localhost:8080

## Flujo de publicación de ahora en adelante

1. Editar `asesor.html` / `apps_script.js` en el repo local.
2. Pegar el archivo actualizado en el editor de Apps Script.
3. **Actualizar la implementación EXISTENTE** (Implementar → Administrar implementaciones → ✏️ → Versión: Nueva). **No crear una implementación nueva** — cambiaría la URL y rompería los marcadores de los asesores.
4. `git commit` + `git push` a GitHub (juanzulu46/REDA3) como respaldo.
