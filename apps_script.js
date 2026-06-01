// =====================================================================
// REDA3 — Google Apps Script (Backend + Frontend)
// =====================================================================
// DESPLIEGUE COMO WEB APP (HTML servido desde el mismo Apps Script):
// 1. Abre tu Google Sheet "Negocios" → Extensiones → Apps Script.
// 2. En Code.gs: borra el contenido y pega este archivo completo.
// 3. Crear archivo HTML: + (Agregar archivo) → HTML → nombre "asesor"
//    (sin la extensión .html, Apps Script la añade sola).
//    Pega el contenido de asesor.html dentro y guarda.
// 4. Guarda todo (Ctrl+S).
// 5. Implementar → Nueva implementación → tipo "Aplicación web":
//    - Ejecutar como: Yo
//    - Quién tiene acceso: Cualquier persona (o "con cuenta de Google" si quieres
//      forzar login Google además del password de la app).
// 6. Copia la URL del Web App: ahí entran los asesores (no necesita server.js).
//
// MODO DESARROLLO LOCAL (opcional, para probar sin redeploy):
// - npm start con server.js sirve asesor.html en http://localhost:8080
// - El frontend detecta automáticamente si está en GAS o local y usa el
//   transporte adecuado (google.script.run vs fetch al proxy /api).
// =====================================================================

// ===== CONFIGURACIÓN CUENTA DE COBRO =====
// *** IMPORTANTE: reemplaza estos valores antes de usar "Imprimir cuenta de cobro" ***
// 1. Sube "Cuenta Cobro.docx" a Google Drive
// 2. Ábrelo con Google Docs (Abrir con → Google Docs) — se crea una copia editable
// 3. Edita el template reemplazando los datos fijos por estos placeholders EXACTOS:
//      {{ciudad_emision}}, {{fecha_texto}}
//      {{empresa_razon_social}}, {{empresa_nit}}
//      {{asesor_nombre}}, {{asesor_cedula}}, {{asesor_ciudad_cc}}
//      {{valor_numero}}, {{valor_letras}}
//      {{concepto}}
//      {{asesor_direccion}}
//      {{banco}}, {{tipo_cuenta}}, {{numero_cuenta}}
// 4. Copia el ID del Google Doc (de la URL: /document/d/ID_AQUI/edit) y pégalo abajo
const EMPRESA_RAZON_SOCIAL = 'BIENES S.A.S';        // TODO: confirmar razón social
const EMPRESA_NIT          = '900.144.609-8';       // TODO: confirmar NIT
const CIUDAD_EMISION       = 'Pereira';
const GERENTE_EMAIL        = 'german.zuluaga@a3inmobiliarios.com';
const TEMPLATE_CUENTA_COBRO_ID = '1AuwcoDnsX_NX3k-jFe6xSUNQfWtgLT4Oy3sTynGAx5c';

// ===== CONFIGURACIÓN =====
// Nombres de las hojas en tu Google Sheet (deben existir)
const HOJAS = {
  asesores: 'Asesores',
  inmuebles: 'Inmuebles',
  clientes: 'Clientes',
  arriendos: 'Arriendos',
  ventas: 'Ventas',
  pagos: 'Pagos',
  comisiones: 'Comisiones',
  partes: 'Partes',
  oficina: 'Oficina',
  origen: 'Origen',
  zona: 'Zona',
  acciones: 'Acciones',
  tipos_accion: 'TipoAccion',
  bonificaciones: 'Bonificaciones',
  parametros: 'Parametros',
  cobros_arriendo: 'CobrosArriendo',
  bonificaciones_mes: 'BonificacionesMes',
  ppto: 'Ppto'
};

// Columnas de cada hoja (en orden exacto)
const COLUMNAS = {
  asesores: ['id_asesor', 'nombre', 'vinculacion', 'estado',
             'cedula', 'ciudad_cc', 'direccion', 'banco', 'tipo_cuenta', 'numero_cuenta', 'email',
             'password', 'rol'],
  inmuebles: ['id_inmueble', 'codigo_plataforma', 'nombre', 'ciudad', 'zona', 'tipo', 'residencial_comercial', 'estado'],
  clientes: ['id_cliente', 'nombre', 'telefono', 'email', 'tipo_persona', 'tipo_documento', 'numero_documento', 'activo'],
  arriendos: ['id_arriendo', 'año', 'mes', 'mercado', 'id_inmueble',
              'valor_canon', 'administracion', 'pct_comision_oficina', 'comision_oficina',
              'oficina_captacion', 'origen_captacion', 'oficina_cierre', 'origen_cierre',
              'referido_captador', 'numero_captador_r', 'valor_ref_captador',
              'referido_cerrador', 'numero_cerrador_r', 'valor_ref_cerrador',
              'meses_contrato', 'estado_arriendo'],
  ventas: ['id_venta', 'año', 'mes', 'mercado', 'id_inmueble',
           'valor_base_comision', 'pct_comision_oficina', 'comision_oficina',
           'comision_por_punta',
           'oficina_captacion', 'origen_captacion', 'oficina_cierre', 'origen_cierre',
           'referido_captador', 'numero_captador_r', 'valor_ref_captador',
           'referido_cerrador', 'numero_cerrador_r', 'valor_ref_cerrador',
           'estado_venta'],
  pagos: ['id_pago', 'id_venta', 'fecha_pago', 'año_pago', 'mes_pago', 'valor_cobrado', 'observacion', 'estado'],
  comisiones: ['id_asesor', 'id_negocio', 'valor_comision', 'punta', 'participacion', 'estado'],
  partes: ['id_parte', 'id_negocio', 'tipo_negocio', 'rol', 'id_cliente', 'participacion_pct'],
  oficina: ['id_oficina', 'nombre'],
  origen: ['id_origen', 'nombre', 'circulo'],
  zona: ['id_zona', 'comuna', 'ciudad'],
  acciones: ['id_accion', 'id_asesor', 'fecha', 'mes', 'tipo', 'descripcion'],
  tipos_accion: ['id_tipo', 'nombre', 'activo'],
  cobros_arriendo: ['id_cobro', 'id_arriendo', 'año_cobro', 'mes_cobro', 'fecha_pago', 'valor_cobrado', 'estado', 'observacion'],
  bonificaciones_mes: ['id_bonmes', 'id_asesor', 'año', 'mes', 'fecha', 'categoria', 'comision_generada', 'acciones_mes', 'fijo', 'pct_variable', 'variable', 'total', 'continuidad', 'calculado_en', 'cobrada_en']
};

// ===== UTILIDADES =====

// Función de un solo uso para forzar la autorización de Drive, Documents y Gmail.
// Ejecútala manualmente desde el editor de Apps Script (▶ Ejecutar → autorizarPermisos)
// la primera vez. SIN try/catch para que el primer error de permisos dispare el popup
// de autorización. Después de aceptar permisos, vuelve a ejecutarla y debería terminar OK.
function autorizarPermisos() {
  // 1) Drive: leer carpeta raíz para forzar scope drive
  Logger.log('Probando DriveApp...');
  var rootName = DriveApp.getRootFolder().getName();
  Logger.log('DriveApp OK - root: ' + rootName);

  // 2) Drive + Documents: abrir el template para forzar scope documents
  Logger.log('Probando template del Doc...');
  var f = DriveApp.getFileById(TEMPLATE_CUENTA_COBRO_ID);
  Logger.log('Template encontrado: ' + f.getName());
  var doc = DocumentApp.openById(TEMPLATE_CUENTA_COBRO_ID);
  Logger.log('DocumentApp OK - doc: ' + doc.getName());

  // 2b) Forzar scope drive COMPLETO (no solo readonly): hacer una copia y borrarla
  // Esto es necesario porque el flujo real de cuenta de cobro usa makeCopy().
  Logger.log('Probando makeCopy (scope drive completo)...');
  var copiaTemp = f.makeCopy('__test_autorizacion_borrar__');
  copiaTemp.setTrashed(true);
  Logger.log('makeCopy OK - copia temporal creada y borrada');

  // 3) Mail: consultar cuota para forzar scope send_mail
  Logger.log('Probando MailApp...');
  var quota = MailApp.getRemainingDailyQuota();
  Logger.log('MailApp OK - cuota restante: ' + quota);

  Logger.log('=== Autorización completada exitosamente ===');
}

function getSheet(nombre) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
}

// Parseo numérico tolerante a formato español ("1.234,56") y valores vacíos.
function numVal(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var s = String(v).trim();
  if (s === '') return 0;
  if (s.indexOf(',') > -1) s = s.replace(/\./g, '').replace(',', '.');
  var n = Number(s);
  return isFinite(n) ? n : 0;
}

// Normaliza un nombre para comparar duplicados: minúsculas, sin tildes, sin signos, espacios simples
function normalizarNombre_(s) {
  if (s === null || s === undefined) return '';
  return String(s).toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// Lee todos los datos de una hoja y devuelve array de objetos
function leerHoja(nombreHoja) {
  const sheet = getSheet(nombreHoja);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  }).filter(obj => {
    // Filtrar filas vacías
    return Object.values(obj).some(v => v !== '' && v !== null && v !== undefined);
  });
}

// ===== CACHE DE HOJAS (Fase 1 performance) =====
// Cachea el resultado de leerHoja() en CacheService por TTL.
// Las hojas grandes (>95KB serializadas) saltean el cache y se leen siempre frescas.
// Para invalidar cuando se escribe, llamar invalidarCacheHojas([...]).
const TTL_HOJA = {
  // Catálogos casi estáticos (cambian rara vez)
  'Oficina': 1800, 'Origen': 1800, 'Zona': 1800,
  'TipoAccion': 1800, 'Bonificaciones': 1800, 'Parametros': 1800,
  // Maestros que pueden cambiar varias veces al día
  'Asesores': 300, 'Inmuebles': 120, 'Clientes': 120,
  // Datos transaccionales (cambian al guardar negocios)
  'Arriendos': 30, 'Ventas': 30, 'Pagos': 30, 'Comisiones': 30,
  'Partes': 30, 'CobrosArriendo': 30, 'Acciones': 60,
  'BonificacionesMes': 60, 'Ppto': 1800
};
function leerHojaCache(nombreHoja, ttlOverride) {
  var ttl = ttlOverride || TTL_HOJA[nombreHoja] || 60;
  var key = 'hoja_' + nombreHoja;
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(key);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) { /* fall through */ }
    }
    var data = leerHoja(nombreHoja);
    try {
      var json = JSON.stringify(data);
      if (json.length < 95 * 1024) cache.put(key, json, ttl);
    } catch (e) { /* hoja demasiado grande o no serializable */ }
    return data;
  } catch (e) {
    // Si el cache no está disponible, leer fresco
    return leerHoja(nombreHoja);
  }
}
function invalidarCacheHojas(listaNombres) {
  if (!listaNombres || !listaNombres.length) return;
  try {
    var cache = CacheService.getScriptCache();
    cache.removeAll(listaNombres.map(function(n){ return 'hoja_' + n; }));
  } catch (e) { /* ignore */ }
}

// Genera el siguiente ID secuencial (ASE-054, INM-219, etc.)
function siguienteId(nombreHoja, prefijo) {
  const datos = leerHoja(nombreHoja);
  let max = 0;
  datos.forEach(d => {
    const id = d[Object.keys(d)[0]] || '';
    const num = parseInt(String(id).split('-')[1]);
    if (!isNaN(num) && num > max) max = num;
  });
  return prefijo + '-' + String(max + 1).padStart(3, '0');
}

// Devuelve los meses del contrato del arriendo. Para registros legacy sin el campo,
// infiere por la regla del % comisión (≤10% → 12 meses administración, >10% → 1 mes colocación).
function mesesContratoDe(arriendo) {
  var n = parseInt(arriendo.meses_contrato, 10);
  if (n && n > 0) return n;
  var pct = numVal(arriendo.pct_comision_oficina);
  return (pct > 0 && pct <= 0.10) ? 12 : 1;
}

// Genera N filas en CobrosArriendo desde el mes de firma del arriendo hacia adelante
// (N = meses_contrato). Cada fila nace COBRADO con fecha_pago = día 1 del mes_cobro.
// Gerencia puede inhabilitar (NO_COBRADO) o cancelar desde su perfil si algún mes no se cobra.
function generarCobrosProyectados(arriendo) {
  var meses = mesesContratoDe(arriendo);
  var comMensual = numVal(arriendo.comision_oficina);
  var anoBase = parseInt(arriendo['año'], 10) || new Date().getFullYear();
  var mesBase = parseInt(arriendo.mes, 10) || (new Date().getMonth() + 1); // 1..12
  for (var i = 0; i < meses; i++) {
    var mTotal = mesBase + i;
    var anoCobro = anoBase + Math.floor((mTotal - 1) / 12);
    var mesCobro = ((mTotal - 1) % 12) + 1;
    agregarFila(HOJAS.cobros_arriendo, COLUMNAS.cobros_arriendo, {
      id_cobro: siguienteId(HOJAS.cobros_arriendo, 'COB'),
      id_arriendo: arriendo.id_arriendo,
      'año_cobro': anoCobro,
      mes_cobro: mesCobro,
      fecha_pago: new Date(anoCobro, mesCobro - 1, 1),
      valor_cobrado: comMensual,
      estado: 'COBRADO',
      observacion: ''
    });
  }
}

// Crea hoja CobrosArriendo si no existe, agrega columna meses_contrato a Arriendos
// si falta y actualiza el umbral de PIEDRA en Bonificaciones.
// Se ejecuta una vez vía endpoint setup_cobros_arriendo.
function setupCobrosArriendo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resultado = { sheet_creada: false, columna_agregada: false, piedra_actualizada: false };

  // 1) Crear hoja CobrosArriendo
  var hojaCobros = ss.getSheetByName(HOJAS.cobros_arriendo);
  if (!hojaCobros) {
    hojaCobros = ss.insertSheet(HOJAS.cobros_arriendo);
    hojaCobros.appendRow(COLUMNAS.cobros_arriendo);
    hojaCobros.getRange(1, 1, 1, COLUMNAS.cobros_arriendo.length).setFontWeight('bold');
    resultado.sheet_creada = true;
  } else {
    // Hoja ya existe: agregar columna fecha_pago si falta (entre mes_cobro y valor_cobrado)
    var headersCob = hojaCobros.getRange(1, 1, 1, hojaCobros.getLastColumn()).getValues()[0];
    if (headersCob.indexOf('fecha_pago') === -1) {
      var idxValor = headersCob.indexOf('valor_cobrado');
      if (idxValor === -1) {
        hojaCobros.getRange(1, headersCob.length + 1).setValue('fecha_pago').setFontWeight('bold');
      } else {
        hojaCobros.insertColumnBefore(idxValor + 1);
        hojaCobros.getRange(1, idxValor + 1).setValue('fecha_pago').setFontWeight('bold');
      }
      resultado.fecha_pago_agregada = true;
    }
  }

  // 2) Agregar columna meses_contrato a Arriendos
  var hojaArr = ss.getSheetByName(HOJAS.arriendos);
  var headersArr = hojaArr.getRange(1, 1, 1, hojaArr.getLastColumn()).getValues()[0];
  if (headersArr.indexOf('meses_contrato') === -1) {
    hojaArr.getRange(1, headersArr.length + 1).setValue('meses_contrato').setFontWeight('bold');
    resultado.columna_agregada = true;
  }

  // 3) Actualizar umbral PIEDRA en Bonificaciones (3.260.417 según Comisiones.xlsx)
  actualizarFila(HOJAS.bonificaciones, 'categoria', 'PIEDRA', { min_comision_oficina: 3260417 });
  resultado.piedra_actualizada = true;

  return resultado;
}

// Asegura que la hoja BonificacionesMes tenga la columna cobrada_en al final.
// Idempotente: si ya existe, no hace nada. Útil para hojas creadas antes del cambio.
function asegurarColumnaCobradaEn() {
  var sheet = getSheet(HOJAS.bonificaciones_mes);
  if (!sheet) return false;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('cobrada_en') === -1) {
    sheet.getRange(1, headers.length + 1).setValue('cobrada_en').setFontWeight('bold');
    return true;
  }
  return false;
}

// Asegura que la hoja Pagos tenga la columna estado al final.
// Idempotente. Permite marcar pagos como ANULADO al cancelar la venta asociada.
function asegurarColumnaEstadoPagos() {
  var sheet = getSheet(HOJAS.pagos);
  if (!sheet) return false;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('estado') === -1) {
    sheet.getRange(1, headers.length + 1).setValue('estado').setFontWeight('bold');
    return true;
  }
  return false;
}

// Asegura que la hoja Arriendos tenga la columna estado_arriendo al final.
// Idempotente. Permite marcar arriendos como CANCELADO.
function asegurarColumnaEstadoArriendo() {
  var sheet = getSheet(HOJAS.arriendos);
  if (!sheet) return false;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('estado_arriendo') === -1) {
    sheet.getRange(1, headers.length + 1).setValue('estado_arriendo').setFontWeight('bold');
    return true;
  }
  return false;
}

// Inserta o actualiza la fila COBRE en la hoja Bonificaciones.
// COBRE: ≥ $5.540.000 + ≥ 5 acciones → fijo $250.000 + variable (4%/5%).
// Orden: entre BRONCE y PIEDRA en la cascada.
// Ejecutar 1 vez desde el editor de Apps Script (▶ Ejecutar → setupCobre).
function setupCobre() {
  var sheet = getSheet(HOJAS.bonificaciones);
  if (!sheet) throw new Error('Hoja "' + HOJAS.bonificaciones + '" no encontrada');

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var catIdx = headers.indexOf('categoria');
  var ordenIdx = headers.indexOf('orden');
  if (catIdx === -1) throw new Error('Columna "categoria" no encontrada en Bonificaciones');

  var datosCobre = {
    categoria: 'COBRE',
    min_comision_oficina: 5540000,
    min_acciones: 5,
    fijo: 250000,
    fijo_medio: 0,
    pct_variable_inicial: 0.04,
    pct_variable_continuidad: 0.05
  };

  // Buscar BRONCE/PIEDRA para calcular un orden intermedio, y COBRE para detectar update
  var ordenBronce = null, ordenPiedra = null, filaCobre = -1;
  for (var i = 1; i < data.length; i++) {
    var cat = String(data[i][catIdx] || '').toUpperCase();
    var orden = ordenIdx !== -1 ? Number(data[i][ordenIdx]) : null;
    if (cat === 'BRONCE') ordenBronce = orden;
    if (cat === 'PIEDRA') ordenPiedra = orden;
    if (cat === 'COBRE') filaCobre = i;
  }

  if (ordenIdx !== -1) {
    if (ordenBronce !== null && !isNaN(ordenBronce) && ordenPiedra !== null && !isNaN(ordenPiedra)) {
      datosCobre.orden = (ordenBronce + ordenPiedra) / 2;
    } else if (ordenBronce !== null && !isNaN(ordenBronce)) {
      datosCobre.orden = ordenBronce + 0.5;
    }
  }

  var resultado = { creada: false, actualizada: false, orden: datosCobre.orden };

  if (filaCobre !== -1) {
    Object.keys(datosCobre).forEach(function(key) {
      var ci = headers.indexOf(key);
      if (ci !== -1) sheet.getRange(filaCobre + 1, ci + 1).setValue(datosCobre[key]);
    });
    resultado.actualizada = true;
  } else {
    var fila = headers.map(function(h) { return datosCobre[h] !== undefined ? datosCobre[h] : ''; });
    sheet.appendRow(fila);
    resultado.creada = true;
  }

  return resultado;
}

// ===== PARTES (N clientes por rol por negocio) =====
// Valida un arreglo de partes y las escribe en la hoja Partes.
// partes = [{rol, id_cliente, participacion_pct}, ...]
// rolesRequeridos = ['arrendador','arrendatario'] | ['vendedor','comprador']
// Retorna null si ok, string de error si falla.
// Valida coherencia de referidos: si hay valor_ref > 0, debe haber nombre;
// los valores nunca pueden ser negativos. Retorna error o null.
function validarReferidos(datos) {
  var nombreCap = String(datos.referido_captador || '').trim();
  var valorCap  = numVal(datos.valor_ref_captador);
  if (valorCap < 0) return 'El valor del referido de captación no puede ser negativo';
  if (valorCap > 0 && !nombreCap) return 'Hay valor de referido en captación pero falta el nombre';

  var nombreCer = String(datos.referido_cerrador || '').trim();
  var valorCer  = numVal(datos.valor_ref_cerrador);
  if (valorCer < 0) return 'El valor del referido de cierre no puede ser negativo';
  if (valorCer > 0 && !nombreCer) return 'Hay valor de referido en cierre pero falta el nombre';
  return null;
}

// Valida el array de comisiones_asesores:
// - cada id_asesor existe en Asesores
// - punta es Captador o Cerrador
// - suma de participaciones por punta = 100% (tolerancia ±0.5%)
// - ningún asesor duplicado en la misma punta
// Retorna string con error o null si está OK.
function validarComisionesAsesores(comisiones, asesoresRef) {
  if (!comisiones || !Array.isArray(comisiones) || comisiones.length === 0) return null;
  var puntasValidas = ['Captador','Cerrador'];
  var porPunta = { Captador: [], Cerrador: [] };
  for (var i = 0; i < comisiones.length; i++) {
    var c = comisiones[i];
    if (puntasValidas.indexOf(c.punta) === -1) {
      return 'Punta inválida "' + c.punta + '" en comisiones';
    }
    if (!asesoresRef.some(function(a){ return a.id_asesor === c.id_asesor; })) {
      return 'Asesor "' + c.id_asesor + '" no existe';
    }
    porPunta[c.punta].push(c);
  }
  for (var p = 0; p < puntasValidas.length; p++) {
    var arr = porPunta[puntasValidas[p]];
    if (arr.length === 0) continue;
    var vistos = {};
    var suma = 0;
    for (var j = 0; j < arr.length; j++) {
      if (vistos[arr[j].id_asesor]) {
        return 'Asesor "' + arr[j].id_asesor + '" duplicado en punta "' + puntasValidas[p] + '"';
      }
      vistos[arr[j].id_asesor] = true;
      suma += numVal(arr[j].participacion);
    }
    if (Math.abs(suma - 100) > 0.5) {
      return 'La suma de participación en punta "' + puntasValidas[p] + '" debe ser 100% (actual: ' + suma.toFixed(2) + '%)';
    }
  }
  return null;
}

function validarYGuardarPartes(idNegocio, tipoNegocio, rolesRequeridos, partes, clientesRef, escribir) {
  if (!partes || !Array.isArray(partes) || partes.length === 0) {
    return 'Debe haber al menos un cliente por rol (' + rolesRequeridos.join(', ') + ')';
  }
  // Chequear que haya al menos 1 por rol requerido
  for (var i = 0; i < rolesRequeridos.length; i++) {
    var rol = rolesRequeridos[i];
    if (!partes.some(function(p){ return p.rol === rol; })) {
      return 'Falta al menos un cliente con rol "' + rol + '"';
    }
  }
  // Integridad referencial
  for (var j = 0; j < partes.length; j++) {
    var p = partes[j];
    if (rolesRequeridos.indexOf(p.rol) === -1) {
      return 'Rol inválido "' + p.rol + '". Permitidos: ' + rolesRequeridos.join(', ');
    }
    if (!clientesRef.some(function(c){ return String(c.id_cliente) === String(p.id_cliente); })) {
      return 'Cliente "' + p.id_cliente + '" no existe';
    }
  }
  // Suma de participacion_pct por rol = 1.0 (decimal 0-1) con tolerancia ±0.01
  for (var k = 0; k < rolesRequeridos.length; k++) {
    var r = rolesRequeridos[k];
    var suma = partes.filter(function(pp){ return pp.rol === r; })
      .reduce(function(acc, pp){ return acc + numVal(pp.participacion_pct); }, 0);
    if (Math.abs(suma - 1) > 0.005) {
      return 'La suma de participación para rol "' + r + '" debe ser 100% (actual: ' + (suma*100).toFixed(2) + '%)';
    }
  }
  // Duplicados: un mismo cliente no puede aparecer 2 veces en el mismo rol
  for (var m = 0; m < rolesRequeridos.length; m++) {
    var rr = rolesRequeridos[m];
    var vistos = {};
    var parDup = partes.filter(function(pp){ return pp.rol === rr; });
    for (var n = 0; n < parDup.length; n++) {
      var id = String(parDup[n].id_cliente);
      if (vistos[id]) return 'Cliente "' + id + '" está duplicado en rol "' + rr + '"';
      vistos[id] = true;
    }
  }
  // Conflicto entre roles: el mismo cliente no puede ser ambos lados del negocio
  // (arrendador y arrendatario, o vendedor y comprador)
  if (rolesRequeridos.length === 2) {
    var rolA = rolesRequeridos[0], rolB = rolesRequeridos[1];
    var clientesA = partes.filter(function(pp){ return pp.rol === rolA; }).map(function(pp){ return String(pp.id_cliente); });
    var clientesB = partes.filter(function(pp){ return pp.rol === rolB; }).map(function(pp){ return String(pp.id_cliente); });
    for (var q = 0; q < clientesA.length; q++) {
      if (clientesB.indexOf(clientesA[q]) !== -1) {
        return 'El cliente "' + clientesA[q] + '" no puede ser "' + rolA + '" y "' + rolB + '" al mismo tiempo';
      }
    }
  }
  // Escribir (salvo que se pida solo validar con escribir === false)
  if (escribir !== false) {
    partes.forEach(function(p) {
      var idParte = siguienteId(HOJAS.partes, 'PRT');
      agregarFila(HOJAS.partes, COLUMNAS.partes, {
        id_parte: idParte,
        id_negocio: idNegocio,
        tipo_negocio: tipoNegocio,
        rol: p.rol,
        id_cliente: p.id_cliente,
        participacion_pct: numVal(p.participacion_pct)
      });
    });
  }
  return null;
}

// Borra todas las partes de un negocio (para editar)
function borrarPartesDeNegocio(idNegocio) {
  var sheet = getSheet(HOJAS.partes);
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;
  var headers = data[0];
  var idxNeg = headers.indexOf('id_negocio');
  if (idxNeg === -1) return 0;
  var borradas = 0;
  // Recorrer de abajo hacia arriba para no desfasar índices
  for (var r = data.length - 1; r >= 1; r--) {
    if (String(data[r][idxNeg]) === String(idNegocio)) {
      sheet.deleteRow(r + 1);
      borradas++;
    }
  }
  return borradas;
}

// Borra todas las filas de una hoja cuyo valor en colNombre coincide con valor.
// Genérico (usado para editar/eliminar negocios: comisiones, partes, pagos, cobros).
function borrarFilasPorColumna_(nombreHoja, colNombre, valor) {
  var sheet = getSheet(nombreHoja);
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;
  var idx = data[0].indexOf(colNombre);
  if (idx === -1) return 0;
  var n = 0;
  for (var r = data.length - 1; r >= 1; r--) {
    if (String(data[r][idx]) === String(valor)) { sheet.deleteRow(r + 1); n++; }
  }
  return n;
}

// Periodos {ano, mes} en que la comisión de un negocio impacta una bonificación:
//   arriendo → el mes del arriendo
//   venta    → el mes_pago de cada pago (o el mes de la venta si no tiene pagos)
function periodosFinancierosNegocio_(idNegocio, tipoNegocio) {
  var periodos = [];
  if (tipoNegocio === 'arriendo') {
    var arr = leerHoja(HOJAS.arriendos).find(function(a){ return String(a.id_arriendo) === String(idNegocio); });
    if (arr) periodos.push({ ano: parseInt(arr['año'], 10) || 0, mes: parseInt(arr.mes, 10) || 0 });
  } else {
    var vnt = leerHoja(HOJAS.ventas).find(function(v){ return String(v.id_venta) === String(idNegocio); });
    var pagosV = leerHoja(HOJAS.pagos).filter(function(p){ return String(p.id_venta) === String(idNegocio); });
    if (pagosV.length) {
      pagosV.forEach(function(p){
        periodos.push({ ano: parseInt(p['año_pago'] || p['ano_pago'], 10) || 0, mes: parseInt(p.mes_pago, 10) || 0 });
      });
    } else if (vnt) {
      periodos.push({ ano: parseInt(vnt['año'], 10) || 0, mes: parseInt(vnt.mes, 10) || 0 });
    }
  }
  return periodos;
}

// True si algún asesor del negocio YA cobró su bonificación en un periodo que el negocio
// impacta. En ese caso, editar o borrar el negocio desincronizaría una cuenta de cobro ya
// emitida → no se permite borrado real; se debe cancelar (soft) dejando rastro de auditoría.
function negocioBloqueadoPorCobro_(idNegocio, tipoNegocio) {
  var comis = leerHoja(HOJAS.comisiones).filter(function(c){ return String(c.id_negocio) === String(idNegocio); });
  if (!comis.length) return false;
  var asesoresNeg = {};
  comis.forEach(function(c){ asesoresNeg[c.id_asesor] = true; });
  var periodos = periodosFinancierosNegocio_(idNegocio, tipoNegocio);
  if (!periodos.length) return false;
  var bonos = leerHoja(HOJAS.bonificaciones_mes).filter(function(b){ return b.cobrada_en; });
  return bonos.some(function(b){
    if (!asesoresNeg[b.id_asesor]) return false;
    var ba = parseInt(b['año'], 10) || 0, bm = parseInt(b.mes, 10) || 0;
    return periodos.some(function(p){ return p.ano === ba && p.mes === bm; });
  });
}

// Agrega una fila a una hoja
function agregarFila(nombreHoja, columnas, datos) {
  const sheet = getSheet(nombreHoja);
  if (!sheet) throw new Error('Hoja "' + nombreHoja + '" no encontrada');
  const fila = columnas.map(col => datos[col] !== undefined ? datos[col] : '');
  sheet.appendRow(fila);
}

// Actualiza una fila existente buscando por ID
function actualizarFila(nombreHoja, colId, idBuscado, datos) {
  var sheet = getSheet(nombreHoja);
  if (!sheet) throw new Error('Hoja "' + nombreHoja + '" no encontrada');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIdx = headers.indexOf(colId);
  if (colIdx === -1) throw new Error('Columna "' + colId + '" no encontrada');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colIdx]) === String(idBuscado)) {
      Object.keys(datos).forEach(function(key) {
        var ci = headers.indexOf(key);
        if (ci !== -1) sheet.getRange(i + 1, ci + 1).setValue(datos[key]);
      });
      return true;
    }
  }
  throw new Error(colId + ' "' + idBuscado + '" no encontrado');
}

// Calcula la categoría de bonificación de un asesor en un mes específico.
// Usa los datos pre-cargados (datos.arriendos, ventas, comisiones, acciones, bonificaciones)
// para evitar releer las hojas.
// Retorna: { categoria, esMedio, escalon, comisionGeneradaOficina, totalRecibido, numAcciones }
function calcularCategoriaMes(idAsesor, mes, datos) {
  var arriendos = datos.arriendos;
  var ventas = datos.ventas;
  var pagos = datos.pagos;
  var comisiones = datos.comisiones;
  var acciones = datos.acciones;
  var bonificaciones = datos.bonificaciones;

  // Comisiones del asesor (excluye ANULADAS por cancelación de venta)
  var misCom = comisiones.filter(function(c) {
    if (c.id_asesor !== idAsesor) return false;
    if (String(c.estado || '').toUpperCase() === 'ANULADA') return false;
    return true;
  });
  var negociosIds = misCom.map(function(c) { return c.id_negocio; });

  // Función auxiliar para sumar participación del asesor en un negocio
  function sumParticipacion(idNegocio) {
    return misCom
      .filter(function(c) { return c.id_negocio === idNegocio; })
      .reduce(function(acc, c) {
        var p = c.participacion === '' || c.participacion === null || c.participacion === undefined
          ? 1 : (Number(c.participacion) || 0);
        return acc + p;
      }, 0);
  }

  var comisionGeneradaOficina = 0;
  var totalRecibido = 0;
  var hoy = new Date();

  // --- Arriendos: filtran por mes del arriendo, anualizados según meses_contrato ---
  arriendos.forEach(function(a) {
    if (parseInt(a.mes, 10) !== mes) return;
    if (negociosIds.indexOf(a.id_arriendo) === -1) return;
    var meses = mesesContratoDe(a);
    comisionGeneradaOficina += (Number(a.comision_oficina) || 0) * meses * 0.5 * sumParticipacion(a.id_arriendo);
  });

  // --- Ventas: filtran por mes_pago de cada pago efectuado (fecha_pago <= hoy) ---
  var pagosDelMes = pagos.filter(function(p) {
    if (parseInt(p.mes_pago, 10) !== mes) return false;
    if (!p.fecha_pago || p.fecha_pago === '') return true;
    // Comparar usando año_pago y mes_pago si la fecha no parsea bien
    var fp = new Date(p.fecha_pago);
    if (isNaN(fp.getTime())) {
      // Fallback: si tiene año_pago y mes_pago, comparar solo año-mes
      var ap = Number(p.año_pago) || 0;
      var mp = Number(p.mes_pago) || 0;
      if (ap && mp) {
        var hoyAM = hoy.getFullYear() * 100 + (hoy.getMonth() + 1);
        return (ap * 100 + mp) <= hoyAM;
      }
      return true;
    }
    return fp <= hoy;
  });

  pagosDelMes.forEach(function(p) {
    var venta = ventas.find(function(v) { return v.id_venta === p.id_venta; });
    if (!venta) return;
    if (String(venta.estado_venta).toUpperCase() === 'CANCELADA') return;
    if (negociosIds.indexOf(venta.id_venta) === -1) return;
    // El pago representa dinero que entró a la oficina
    comisionGeneradaOficina += (Number(p.valor_cobrado) || 0) * 0.5 * sumParticipacion(venta.id_venta);
  });

  // Total recibido: arriendos del mes + ventas proporcional a pagos del mes
  misCom.forEach(function(c) {
    var arriendo = arriendos.find(function(a) { return a.id_arriendo === c.id_negocio; });
    if (arriendo && parseInt(arriendo.mes, 10) === mes) {
      totalRecibido += Number(c.valor_comision) || 0;
    }
  });

  pagosDelMes.forEach(function(p) {
    var venta = ventas.find(function(v) { return v.id_venta === p.id_venta; });
    if (!venta) return;
    if (String(venta.estado_venta).toUpperCase() === 'CANCELADA') return;
    var comOficina = Number(venta.comision_oficina) || 0;
    if (comOficina === 0) return;
    var fraccion = (Number(p.valor_cobrado) || 0) / comOficina;
    misCom.filter(function(c) { return c.id_negocio === venta.id_venta; })
      .forEach(function(c) {
        totalRecibido += (Number(c.valor_comision) || 0) * fraccion;
      });
  });

  // Acciones comerciales del mes
  var numAcciones = acciones.filter(function(ac) {
    return ac.id_asesor === idAsesor && parseInt(ac.mes, 10) === mes;
  }).length;

  // Filtrar escalones vigentes para el mes consultado y ordenar por 'orden' asc
  // (asume año actual; vigente_desde/vigente_hasta son strings YYYY-MM-DD)
  var escalones = bonificaciones.slice().sort(function(a, b) {
    return (Number(a.orden) || 0) - (Number(b.orden) || 0);
  });

  // Evaluación en cascada: el primer escalón cuyas condiciones cumple el asesor.
  // Nota: PIEDRA tiene un caso especial — si cumple acciones pero no el umbral monetario,
  // y generó comisión > 0, cae en PIEDRA con fijo medio (PIEDRA 1/2). Esto se evalúa
  // dentro del mismo paso de PIEDRA para que tenga prioridad sobre ARENA en la cascada.
  var escalonAsignado = null;
  var esMedio = false;
  for (var i = 0; i < escalones.length; i++) {
    var e = escalones[i];
    var minCom = Number(e.min_comision_oficina) || 0;
    var minAcc = Number(e.min_acciones) || 0;

    // Caso normal: cumple ambos umbrales
    if (comisionGeneradaOficina >= minCom && numAcciones >= minAcc) {
      escalonAsignado = e;
      esMedio = false;
      break;
    }

    // Caso especial PIEDRA medio: cumple acciones y generó comisión > 0 pero no llegó al umbral
    if (String(e.categoria).toUpperCase() === 'PIEDRA' &&
        Number(e.fijo_medio) > 0 &&
        numAcciones >= minAcc &&
        comisionGeneradaOficina > 0 &&
        comisionGeneradaOficina < minCom) {
      escalonAsignado = e;
      esMedio = true;
      break;
    }
  }

  // Si aún no cae en ninguno: ARENA (cumple acciones mínimas pero no comisión),
  // PISO 4% (no cumple acciones pero sí generó comisión a la oficina) o ARENA MOVEDIZA
  var categoria;
  var esPiso4 = false;
  if (escalonAsignado) {
    categoria = String(escalonAsignado.categoria).toUpperCase();
  } else {
    // Buscar fila ARENA en la tabla para usar su min_acciones
    var arenaRow = escalones.find(function(e) { return String(e.categoria).toUpperCase() === 'ARENA'; });
    var arenaMin = arenaRow ? (Number(arenaRow.min_acciones) || 5) : 5;
    if (numAcciones >= arenaMin) {
      categoria = 'ARENA';
      escalonAsignado = arenaRow || null;
    } else if (comisionGeneradaOficina > 0) {
      // Cerró venta o arriendo pero no llegó al mínimo de acciones
      // → 4% de la comisión generada a la oficina (no se persiste escalón)
      categoria = 'PISO 4%';
      escalonAsignado = null;
      esPiso4 = true;
    } else {
      categoria = 'ARENA MOVEDIZA';
      escalonAsignado = null;
    }
  }

  return {
    categoria: categoria,
    esMedio: esMedio,
    esPiso4: esPiso4,
    escalon: escalonAsignado,
    comisionGeneradaOficina: comisionGeneradaOficina,
    totalRecibido: totalRecibido,
    numAcciones: numAcciones
  };
}

// ===== LIQUIDACIÓN MENSUAL DE BONIFICACIONES =====
// Calcula la bonificación de todos los asesores activos para un año/mes dado
// y la persiste en la hoja BonificacionesMes (sobreescribe el periodo si ya existía).
// Retorna { ok, filas_escritas, asesores_procesados, errores }
function liquidarMes(anio, mes) {
  anio = parseInt(anio, 10);
  mes = parseInt(mes, 10);
  if (!anio || anio < 2020 || anio > 2100) throw new Error('Año inválido');
  if (!mes || mes < 1 || mes > 12) throw new Error('Mes inválido (1-12)');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJAS.bonificaciones_mes);
  if (!hoja) throw new Error('Hoja ' + HOJAS.bonificaciones_mes + ' no existe');

  // Pre-cargar datos (una sola vez, evita N lecturas)
  var datos = {
    arriendos: leerHoja(HOJAS.arriendos),
    ventas: leerHoja(HOJAS.ventas),
    pagos: leerHoja(HOJAS.pagos),
    comisiones: leerHoja(HOJAS.comisiones),
    acciones: leerHoja(HOJAS.acciones),
    bonificaciones: leerHoja(HOJAS.bonificaciones)
  };
  var asesores = leerHoja(HOJAS.asesores).filter(function(a){
    return String(a.estado || '').toLowerCase() === 'activo';
  });

  // Borrar filas existentes del periodo (sobreescritura)
  asegurarColumnaCobradaEn();
  var existentes = leerHoja(HOJAS.bonificaciones_mes);

  // Bloquear re-liquidación si alguna fila del periodo ya fue cobrada
  // (los PDFs ya enviados quedarían desincronizados con la nueva liquidación)
  var yaCobradas = existentes.filter(function(row){
    return parseInt(row['año'], 10) === anio
      && parseInt(row.mes, 10) === mes
      && row.cobrada_en;
  });
  if (yaCobradas.length > 0) {
    var idsCobrados = yaCobradas.map(function(r){ return r.id_asesor; }).join(', ');
    throw new Error('No se puede re-liquidar ' + mes + '/' + anio + ': ya hay bonificaciones cobradas por: ' + idsCobrados);
  }

  var filasABorrar = [];
  existentes.forEach(function(row, idx) {
    if (parseInt(row['año'], 10) === anio && parseInt(row.mes, 10) === mes) {
      filasABorrar.push(idx + 2); // +2: 1 por header, 1 por índice base 0
    }
  });
  // Borrar de abajo hacia arriba para no corromper los índices
  filasABorrar.sort(function(a,b){ return b - a; });
  filasABorrar.forEach(function(nroFila){ hoja.deleteRow(nroFila); });

  var filasEscritas = 0;
  var errores = [];
  var ahora = new Date();

  asesores.forEach(function(asesor) {
    try {
      var actual = calcularCategoriaMes(asesor.id_asesor, mes, datos);

      // Determinar continuidad con mes anterior
      var esContinuidad = false;
      if (mes > 1 && actual.escalon) {
        var anterior = calcularCategoriaMes(asesor.id_asesor, mes - 1, datos);
        if (anterior.categoria === actual.categoria) esContinuidad = true;
      }

      var pctVariable = 0, fijoBase = 0, variableBase = 0;
      if (actual.escalon) {
        pctVariable = esContinuidad
          ? (Number(actual.escalon.pct_variable_continuidad) || 0)
          : (Number(actual.escalon.pct_variable_inicial) || 0);
        fijoBase = actual.esMedio
          ? (Number(actual.escalon.fijo_medio) || 0)
          : (Number(actual.escalon.fijo) || 0);
        variableBase = actual.comisionGeneradaOficina * pctVariable;
      } else if (actual.esPiso4) {
        // PISO 4%: pagar 4% de la comisión generada a la oficina, sin fijo
        pctVariable = 0.04;
        fijoBase = 0;
        variableBase = actual.comisionGeneradaOficina * pctVariable;
      }

      // Factor vinculación: empleado /1.3, freelance ×1
      var vinculacion = String(asesor.vinculacion || '').toLowerCase();
      var factorVinc = vinculacion === 'empleado' ? (1 / 1.3) : 1;
      var fijo = fijoBase * factorVinc;
      var variable = variableBase * factorVinc;

      var categoriaLabel = actual.categoria + (actual.esMedio ? ' (1/2)' : '');
      var continuidadLabel = actual.escalon
        ? (esContinuidad ? 'CONTINUA' : 'INICIAL')
        : 'N/A';

      agregarFila(HOJAS.bonificaciones_mes, COLUMNAS.bonificaciones_mes, {
        id_bonmes: siguienteId(HOJAS.bonificaciones_mes, 'BNM'),
        id_asesor: asesor.id_asesor,
        'año': anio,
        mes: mes,
        fecha: new Date(anio, mes - 1, 1),
        categoria: categoriaLabel,
        comision_generada: actual.comisionGeneradaOficina,
        acciones_mes: actual.numAcciones,
        fijo: fijo,
        pct_variable: pctVariable,
        variable: variable,
        total: fijo + variable,
        continuidad: continuidadLabel,
        calculado_en: ahora
      });
      filasEscritas++;
    } catch (err) {
      errores.push({ id_asesor: asesor.id_asesor, error: err.message });
    }
  });

  return {
    ok: true,
    filas_escritas: filasEscritas,
    asesores_procesados: asesores.length,
    errores: errores,
    anio: anio,
    mes: mes
  };
}

// Respuesta JSON con CORS
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== ENDPOINTS =====

// Endpoints que NO requieren sesión validada (login + catálogos básicos).
// Cualquier action fuera de esta lista debe traer id_asesor + password válidos.
var ENDPOINTS_PUBLICOS = ['login', 'catalogos', 'login_password'];

// Puentes para google.script.run (frontend servido desde la propia Web App).
// El iframe sandbox de Apps Script bloquea fetch directo, así que el cliente
// llama a estas funciones globales y aquí se reusa la lógica de doPost/doGet.
function api(body) {
  var output = doPost({ postData: { contents: JSON.stringify(body || {}) } });
  return JSON.parse(output.getContent());
}
function apiGet(params) {
  var output = doGet({ parameter: params || {} });
  // Si no viene action, doGet devuelve HtmlOutput; aquí asumimos params.action existe.
  return JSON.parse(output.getContent());
}

// Valida que el id_asesor + password coincidan con la hoja Asesores.
// Retorna null si OK, mensaje de error si falla.
function validarSesion(idAsesor, password) {
  if (!idAsesor || !password) return 'Sesión inválida: faltan credenciales';
  var asesor = leerHoja(HOJAS.asesores).find(function(a){ return a.id_asesor === idAsesor; });
  if (!asesor) return 'Sesión inválida: asesor no existe';
  if (!asesor.password || String(asesor.password) !== String(password)) {
    return 'Sesión inválida: credenciales incorrectas';
  }
  return null;
}

// Roles con poderes de gestión: gerencia y dirección comercial.
// La directora comercial (rol "directora") tiene los mismos permisos que el gerente.
function esGestor_(asesor) {
  if (!asesor) return false;
  var rol = String(asesor.rol || '').toLowerCase();
  return rol === 'gerente' || rol === 'directora';
}

// Correos de la(s) directora(s) comercial(es): asesores con rol "directora" que tengan email.
// Se usa para que las cuentas de cobro también lleguen a la dirección comercial.
function emailsDirectora_() {
  return leerHoja(HOJAS.asesores)
    .filter(function(a){ return String(a.rol || '').toLowerCase() === 'directora' && a.email; })
    .map(function(a){ return String(a.email).trim(); })
    .filter(Boolean);
}

// Construye la cadena CC de una cuenta de cobro: asesor + directora(s), sin el gerente
// (que va como destinatario principal). Limpia correos mal formados y evita comas
// internas que romperían el join. Devuelve '' si no hay copias válidas.
function ccCobro_(asesorEmail) {
  var lista = [];
  var emailRe = /^[^\s,]+@[^\s,]+\.[^\s,]+$/;
  function add(em){
    em = String(em || '').trim();
    if (em && em !== GERENTE_EMAIL && emailRe.test(em) && lista.indexOf(em) === -1) lista.push(em);
  }
  add(asesorEmail);
  emailsDirectora_().forEach(add);
  return lista.join(',');
}

function doGet(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Manejo seguro de parámetros (a veces e o e.parameter pueden venir vacíos)
    var params = {};
    if (e && e.parameter) {
      params = e.parameter;
    } else if (e && e.queryString) {
      // Parsear manualmente si parameter viene vacío
      e.queryString.split('&').forEach(function(pair) {
        var kv = pair.split('=');
        params[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1] || '');
      });
    }
    var action = params.action || '';

    // --- SERVIR LA PÁGINA: si no viene action, devolver el HTML del frontend ---
    // Permite que la web app entregue tanto el HTML como las llamadas API desde la misma URL.
    // Requiere un archivo HTML en el proyecto Apps Script llamado "asesor" (sin extensión).
    if (!action) {
      return HtmlService.createHtmlOutputFromFile('asesor')
        .setTitle('REDA3 — Portal de Asesores')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // Validar sesión para cualquier endpoint que no sea público
    if (ENDPOINTS_PUBLICOS.indexOf(action) === -1) {
      var errSesGet = validarSesion(params.id_asesor, params.password);
      if (errSesGet) return jsonResponse({ ok: false, error: errSesGet });
    }

    return dispatchGet(params);

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message, stack: err.stack });
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

// Despacha las acciones de LECTURA (las mismas que entran por doGet). Se extrae para
// reusarla desde doPost cuando el frontend envía las lecturas por POST (credenciales en
// el cuerpo, no en la URL → no quedan en logs). NO toma lock: lo gestiona quien la llama.
function dispatchGet(params) {
  var e = null; // compat: el bloque de "acción no reconocida" referencia 'e'
  var action = params.action || '';

    // --- LOGIN: devuelve SOLO lista para el dropdown (sin passwords ni datos sensibles) ---
    if (action === 'login') {
      var asesores = leerHojaCache(HOJAS.asesores);
      var listado = asesores.map(function(a){
        return {
          id_asesor: a.id_asesor,
          nombre: a.nombre,
          estado: a.estado,
          vinculacion: a.vinculacion
        };
      });
      return jsonResponse({ ok: true, asesores: listado });
    }

    // --- CATALOGOS: devuelve todos los catálogos para los selects ---
    if (action === 'catalogos') {
      // Sólo clientes activos (activo != FALSE) para los selects del formulario
      var clientesCat = leerHojaCache(HOJAS.clientes).filter(function(c){
        var a = c.activo;
        if (a === false || String(a).toUpperCase() === 'FALSE') return false;
        return true;
      });
      return jsonResponse({
        ok: true,
        inmuebles: leerHojaCache(HOJAS.inmuebles),
        clientes: clientesCat,
        oficinas: leerHojaCache(HOJAS.oficina),
        origenes: leerHojaCache(HOJAS.origen),
        zonas: leerHojaCache(HOJAS.zona),
        tipos_accion: leerHojaCache(HOJAS.tipos_accion)
      });
    }

    // --- MIS NEGOCIOS: todo lo del asesor ---
    if (action === 'mis_negocios') {
      var idAsesor = params.id_asesor || '';
      if (!idAsesor) return jsonResponse({ ok: false, error: 'Falta id_asesor' });

      var arriendos = leerHojaCache(HOJAS.arriendos);
      var ventas = leerHojaCache(HOJAS.ventas);
      var pagos = leerHojaCache(HOJAS.pagos);
      var comisiones = leerHojaCache(HOJAS.comisiones);
      var inmuebles = leerHojaCache(HOJAS.inmuebles);
      var clientes = leerHojaCache(HOJAS.clientes);
      var partes = leerHojaCache(HOJAS.partes);

      var misComisiones = comisiones.filter(function(c) { return c.id_asesor === idAsesor; });
      var misNegocioIds = misComisiones.map(function(c) { return c.id_negocio; });

      var misArriendos = arriendos.filter(function(a) { return misNegocioIds.indexOf(a.id_arriendo) !== -1; });
      var misVentas = ventas.filter(function(v) { return misNegocioIds.indexOf(v.id_venta) !== -1; });
      var misVentaIds = misVentas.map(function(v) { return v.id_venta; });
      var misPagos = pagos.filter(function(p) { return misVentaIds.indexOf(p.id_venta) !== -1; });
      var misPartes = partes.filter(function(p) { return misNegocioIds.indexOf(p.id_negocio) !== -1; });
      var misArriendoIds = misArriendos.map(function(a) { return a.id_arriendo; });
      var misCobros = leerHojaCache(HOJAS.cobros_arriendo)
        .filter(function(c){ return misArriendoIds.indexOf(c.id_arriendo) !== -1; });

      // Calcular flag p.efectuado para cada pago (antes lo hacía server.js como proxy)
      var hoyMN = new Date();
      var hoyAM = hoyMN.getFullYear() * 12 + (hoyMN.getMonth() + 1);
      misPagos.forEach(function(p) {
        if (!p.fecha_pago || p.fecha_pago === '') {
          p.efectuado = true;
          return;
        }
        var fp = new Date(p.fecha_pago);
        if (isNaN(fp.getTime())) {
          var ap = Number(p['año_pago']) || Number(p['ano_pago']) || 0;
          var mp = Number(p.mes_pago) || 0;
          p.efectuado = (ap && mp) ? ((ap * 12 + mp) <= hoyAM) : true;
        } else {
          p.efectuado = fp <= hoyMN;
        }
      });

      return jsonResponse({
        ok: true,
        arriendos: misArriendos,
        ventas: misVentas,
        pagos: misPagos,
        comisiones: misComisiones,
        inmuebles: inmuebles,
        clientes: clientes,
        partes: misPartes,
        cobros_arriendo: misCobros
      });
    }

    // --- SIGUIENTE ID ---
    if (action === 'siguiente_id') {
      var tipo = params.tipo || '';
      var prefijos = { arriendos: 'ARR', ventas: 'VNT', inmuebles: 'INM', clientes: 'CLI', pagos: 'PAG', partes: 'PRT' };
      var hoja = HOJAS[tipo];
      var prefijo = prefijos[tipo];
      if (!hoja || !prefijo) return jsonResponse({ ok: false, error: 'Tipo inválido' });
      return jsonResponse({ ok: true, id: siguienteId(hoja, prefijo) });
    }

    // --- MIS BONIFICACIONES: calcula bonificación del mes para un asesor ---
    if (action === 'mis_bonificaciones') {
      var idAsesorB = params.id_asesor || '';
      var mesB = parseInt(params.mes || '0', 10);
      if (!idAsesorB || !mesB) return jsonResponse({ ok: false, error: 'Faltan parámetros (id_asesor, mes)' });

      var datosBon = {
        arriendos: leerHojaCache(HOJAS.arriendos),
        ventas: leerHojaCache(HOJAS.ventas),
        pagos: leerHojaCache(HOJAS.pagos),
        comisiones: leerHojaCache(HOJAS.comisiones),
        acciones: leerHojaCache(HOJAS.acciones),
        bonificaciones: leerHojaCache(HOJAS.bonificaciones)
      };

      // Cargar asesor para conocer la vinculación (Empleado vs Freelance)
      var asesorB = leerHojaCache(HOJAS.asesores).find(function(a){ return a.id_asesor === idAsesorB; });
      var vinculacion = asesorB ? String(asesorB.vinculacion || '').toLowerCase() : '';
      // Empleado: el total se divide por 1.3 (G41 del Excel Comisiones.xlsx)
      var factorVinc = vinculacion === 'empleado' ? (1 / 1.3) : 1;

      // Categoría del mes actual
      var actual = calcularCategoriaMes(idAsesorB, mesB, datosBon);

      // Determinar pct_variable según continuidad con el mes anterior
      // Regla: enero (mes 1) siempre arranca en pct_variable_inicial (4%)
      var esContinuidad = false;
      var catAnterior = null;
      if (mesB > 1 && actual.escalon) {
        var anterior = calcularCategoriaMes(idAsesorB, mesB - 1, datosBon);
        catAnterior = anterior.categoria;
        // PIEDRA y PIEDRA con fijo medio cuentan como la misma categoría
        if (anterior.categoria === actual.categoria) {
          esContinuidad = true;
        }
      }

      var pctVariable = 0;
      var fijoBase = 0;
      var variableBase = 0;
      if (actual.escalon) {
        pctVariable = esContinuidad
          ? Number(actual.escalon.pct_variable_continuidad) || 0
          : Number(actual.escalon.pct_variable_inicial) || 0;
        fijoBase = actual.esMedio
          ? (Number(actual.escalon.fijo_medio) || 0)
          : (Number(actual.escalon.fijo) || 0);
        variableBase = actual.comisionGeneradaOficina * pctVariable;
      } else if (actual.esPiso4) {
        // PISO 4%: 4% de la comisión generada a la oficina, sin fijo
        pctVariable = 0.04;
        fijoBase = 0;
        variableBase = actual.comisionGeneradaOficina * pctVariable;
      }

      // Aplicar factor de vinculación (empleado: ÷1.3; freelance: ×1)
      var fijo = fijoBase * factorVinc;
      var variable = variableBase * factorVinc;
      var bonificacionTotal = fijo + variable;

      return jsonResponse({
        ok: true,
        mes: mesB,
        comision_generada_oficina: actual.comisionGeneradaOficina,
        total_recibido: actual.totalRecibido,
        num_acciones: actual.numAcciones,
        categoria: actual.categoria,
        es_medio: actual.esMedio,
        es_piso4: actual.esPiso4,
        bonificacion_fija: fijo,
        bonificacion_variable: variable,
        bonificacion_total: bonificacionTotal,
        bonificacion_fija_freelance: fijoBase,
        bonificacion_variable_freelance: variableBase,
        bonificacion_total_freelance: fijoBase + variableBase,
        vinculacion: asesorB ? asesorB.vinculacion : '',
        factor_vinculacion: factorVinc,
        pct_variable: pctVariable,
        es_continuidad: esContinuidad,
        categoria_mes_anterior: catAnterior
      });
    }

    // --- MIS ACCIONES ---
    if (action === 'mis_acciones') {
      var idAsesorA = params.id_asesor || '';
      var mesA = params.mes ? parseInt(params.mes, 10) : null;
      var todasAcciones = leerHojaCache(HOJAS.acciones);
      var filt = todasAcciones.filter(function(a) {
        if (a.id_asesor !== idAsesorA) return false;
        if (mesA && parseInt(a.mes, 10) !== mesA) return false;
        return true;
      });
      return jsonResponse({ ok: true, acciones: filt });
    }

    // --- TODOS LOS NEGOCIOS (solo gerente) ---
    if (action === 'todos_negocios') {
      var idAsesorG = params.id_asesor || '';
      // Verificar que sea gerente
      var asesoresG = leerHojaCache(HOJAS.asesores);
      var asesorG = asesoresG.find(function(a) { return a.id_asesor === idAsesorG; });
      if (!esGestor_(asesorG)) {
        return jsonResponse({ ok: false, error: 'Acceso denegado' });
      }
      return jsonResponse({
        ok: true,
        arriendos: leerHojaCache(HOJAS.arriendos),
        ventas: leerHojaCache(HOJAS.ventas),
        pagos: leerHojaCache(HOJAS.pagos),
        comisiones: leerHojaCache(HOJAS.comisiones),
        inmuebles: leerHojaCache(HOJAS.inmuebles),
        clientes: leerHojaCache(HOJAS.clientes),
        partes: leerHojaCache(HOJAS.partes),
        cobros_arriendo: leerHojaCache(HOJAS.cobros_arriendo),
        ppto: leerHojaCache(HOJAS.ppto),
        acciones: leerHojaCache(HOJAS.acciones),
        oficina: leerHojaCache(HOJAS.oficina),
        origen: leerHojaCache(HOJAS.origen),
        zona: leerHojaCache(HOJAS.zona),
        bonificaciones_mes: leerHojaCache(HOJAS.bonificaciones_mes),
        bonificaciones: leerHojaCache(HOJAS.bonificaciones),
        asesores: asesoresG.map(function(a) { return { id_asesor: a.id_asesor, nombre: a.nombre, estado: a.estado, rol: a.rol }; })
      });
    }

    // --- COBROS DE UN ARRIENDO (lectura) ---
    if (action === 'mis_cobros_arriendo') {
      var idArrCob = params.id_arriendo || '';
      if (!idArrCob) return jsonResponse({ ok:false, error:'Falta id_arriendo' });
      var cobrosArr = leerHojaCache(HOJAS.cobros_arriendo)
        .filter(function(c){ return String(c.id_arriendo) === String(idArrCob); });
      return jsonResponse({ ok:true, cobros: cobrosArr });
    }

    // --- VERIFICAR DUPLICADO ---
    if (action === 'verificar_duplicado') {
      var tipoDup = params.tipo || '';
      var idInmueble = params.id_inmueble || '';
      var mesDup = params.mes || '';
      // compradores/vendedores pueden venir como CSV: "CLI-001,CLI-003"
      var compradoresCSV = params.id_comprador || '';
      var compradoresList = compradoresCSV
        ? String(compradoresCSV).split(',').map(function(s){ return s.trim(); }).filter(Boolean)
        : [];
      var hojaDup = tipoDup === 'arriendos' ? HOJAS.arriendos : HOJAS.ventas;
      var datosDup = leerHojaCache(hojaDup);
      var duplicado = null;

      // 1) Duplicado por inmueble + mes
      for (var i = 0; i < datosDup.length; i++) {
        var d = datosDup[i];
        if (String(d.id_inmueble) === String(idInmueble) && String(d.mes) === String(mesDup)) {
          duplicado = d;
          break;
        }
      }

      // 2) Duplicado por inmueble + comprador (ventas): usa hoja Partes
      if (!duplicado && tipoDup === 'ventas' && compradoresList.length > 0) {
        var partesAll = leerHojaCache(HOJAS.partes);
        var ventasConInm = datosDup.filter(function(v){
          return String(v.id_inmueble) === String(idInmueble);
        });
        for (var vi = 0; vi < ventasConInm.length && !duplicado; vi++) {
          var compsVenta = partesAll.filter(function(p){
            return String(p.id_negocio) === String(ventasConInm[vi].id_venta) && p.rol === 'comprador';
          }).map(function(p){ return String(p.id_cliente); });
          for (var ci = 0; ci < compradoresList.length; ci++) {
            if (compsVenta.indexOf(compradoresList[ci]) !== -1) { duplicado = ventasConInm[vi]; break; }
          }
        }
      }

      return jsonResponse({ ok: true, duplicado: !!duplicado, negocio: duplicado });
    }

    // Si no se reconoce, mostrar debug
    return jsonResponse({
      ok: false,
      error: 'Acción no reconocida: ' + action,
      debug: { params: params, hasE: !!e, hasParam: !!(e && e.parameter), queryString: e ? e.queryString : 'sin e' }
    });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    // Validar sesión para cualquier endpoint que no sea público
    if (ENDPOINTS_PUBLICOS.indexOf(action) === -1) {
      var errSesPost = validarSesion(body.id_asesor, body.password);
      if (errSesPost) {
        lock.releaseLock();
        return jsonResponse({ ok: false, error: errSesPost });
      }
    }

    // Lecturas enviadas por POST (credenciales en el cuerpo, NO en la URL → no quedan en
    // logs de navegador/servidor/proxy). Se delegan al mismo despachador de doGet.
    var READ_ACTIONS = ['login','catalogos','mis_negocios','siguiente_id','mis_bonificaciones',
                        'mis_acciones','todos_negocios','mis_cobros_arriendo','verificar_duplicado'];
    if (READ_ACTIONS.indexOf(action) !== -1) {
      var outRead = dispatchGet(body);
      lock.releaseLock();
      return outRead;
    }

    // --- LOGIN CON CONTRASEÑA ---
    // body: { action, id_asesor, password }
    // Devuelve el asesor (sin password) si las credenciales son válidas
    if (action === 'login_password') {
      var idLog = body.id_asesor || '';
      var pwdLog = body.password || '';
      if (!idLog || !pwdLog) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Ingrese asesor y contraseña' });
      }
      var asesoresLog = leerHoja(HOJAS.asesores);
      var found = asesoresLog.find(function(a){ return a.id_asesor === idLog; });
      if (!found) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Asesor no encontrado' });
      }
      if (!found.password || String(found.password).trim() === '') {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Este asesor aún no tiene contraseña asignada. Contacta al administrador.' });
      }
      if (String(found.password) !== String(pwdLog)) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Contraseña incorrecta' });
      }
      // Eliminar password del objeto antes de devolverlo
      var safe = {};
      Object.keys(found).forEach(function(k){ if (k !== 'password') safe[k] = found[k]; });
      lock.releaseLock();
      return jsonResponse({ ok:true, asesor: safe });
    }

    // --- REGISTRAR ARRIENDO ---
    if (action === 'registrar_arriendo') {
      const datos = body.datos;
      // Año automático
      if (!datos['año']) datos['año'] = new Date().getFullYear();

      // Validar valores no negativos
      if (numVal(datos.valor_canon) < 0 || numVal(datos.administracion) < 0 || numVal(datos.pct_comision_oficina) < 0) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Canon, administración y porcentaje de comisión no pueden ser negativos' });
      }
      if (numVal(datos.valor_canon) === 0) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'El valor del canon debe ser mayor a 0' });
      }

      // Integridad referencial del inmueble
      var inmRef = leerHoja(HOJAS.inmuebles);
      var cliRef = leerHoja(HOJAS.clientes);
      var inmArr = inmRef.find(function(i){ return String(i.id_inmueble) === String(datos.id_inmueble); });
      if (!inmArr) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Inmueble "' + datos.id_inmueble + '" no existe' });
      }
      if (String(inmArr.estado || '').toLowerCase() === 'inactivo') {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'El inmueble "' + datos.id_inmueble + '" está inactivo. Reactívelo antes de registrar un arriendo.' });
      }

      // Retrocompat: si no vienen partes[] pero sí id_arrendador/id_arrendatario legacy, convertir
      var partesArr = body.partes;
      if ((!partesArr || partesArr.length === 0) && (datos.id_arrendador || datos.id_arrendatario)) {
        partesArr = [];
        if (datos.id_arrendador) partesArr.push({ rol:'arrendador', id_cliente:datos.id_arrendador, participacion_pct:1 });
        if (datos.id_arrendatario) partesArr.push({ rol:'arrendatario', id_cliente:datos.id_arrendatario, participacion_pct:1 });
      }

      // Bloquear duplicado: mismo inmueble + mes + año
      var arrExist = leerHoja(HOJAS.arriendos);
      var dup = arrExist.find(function(a){
        return String(a.id_inmueble) === String(datos.id_inmueble)
            && String(a.mes) === String(datos.mes)
            && String(a['año']) === String(datos['año']);
      });
      if (dup) {
        lock.releaseLock();
        return jsonResponse({
          ok:false,
          error:'Ya existe un arriendo para este inmueble en ' + datos.mes + '/' + datos['año'] + ' (' + dup.id_arriendo + ')',
          duplicado: dup
        });
      }

      // Generar ID
      datos.id_arriendo = siguienteId(HOJAS.arriendos, 'ARR');
      // Calcular comisión sobre canon + administración (consistente con frontend)
      var canonTotalArr = numVal(datos.valor_canon) + numVal(datos.administracion);
      datos.comision_oficina = canonTotalArr * numVal(datos.pct_comision_oficina);
      // Meses del contrato: obligatorio. El frontend lo valida; acá se rechaza
      // si falta para evitar registros sin duración (requerido por CobrosArriendo).
      var mesesContratoNum = parseInt(datos.meses_contrato, 10);
      if (!mesesContratoNum || mesesContratoNum <= 0) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'meses_contrato es obligatorio y debe ser un entero positivo' });
      }
      datos.meses_contrato = mesesContratoNum;

      // Validar y guardar partes (arrendador + arrendatario, suma=100% por rol)
      var errPartes = validarYGuardarPartes(
        datos.id_arriendo, 'arriendo', ['arrendador','arrendatario'], partesArr, cliRef
      );
      if (errPartes) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error: errPartes });
      }

      // Validar comisiones ANTES de escribir nada
      var errComArr = validarComisionesAsesores(body.comisiones_asesores, leerHoja(HOJAS.asesores));
      if (errComArr) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error: errComArr });
      }

      // Validar coherencia de referidos
      var errRefArr = validarReferidos(datos);
      if (errRefArr) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error: errRefArr });
      }

      // Guardar arriendo
      agregarFila(HOJAS.arriendos, COLUMNAS.arriendos, datos);

      // Auto-generar cobros proyectados (uno por cada mes del contrato)
      try { generarCobrosProyectados(datos); } catch(eCob) { /* no bloquear el registro si falla */ }

      // Guardar comisiones de los asesores
      if (body.comisiones_asesores && body.comisiones_asesores.length > 0) {
        body.comisiones_asesores.forEach(com => {
          agregarFila(HOJAS.comisiones, COLUMNAS.comisiones, {
            id_asesor: com.id_asesor,
            id_negocio: datos.id_arriendo,
            valor_comision: com.valor_comision,
            punta: com.punta,
            participacion: (numVal(com.participacion) || 100) / 100,
            estado: 'ACTIVA'
          });
        });
      }

      invalidarCacheHojas([HOJAS.arriendos, HOJAS.comisiones, HOJAS.partes, HOJAS.cobros_arriendo]);
      lock.releaseLock();
      return jsonResponse({ ok: true, id: datos.id_arriendo, mensaje: 'Arriendo registrado' });
    }

    // --- REGISTRAR VENTA ---
    if (action === 'registrar_venta') {
      const datos = body.datos;
      if (!datos['año']) datos['año'] = new Date().getFullYear();

      // Validar valores no negativos
      if (numVal(datos.valor_base_comision) < 0 || numVal(datos.pct_comision_oficina) < 0) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Valor base y porcentaje de comisión no pueden ser negativos' });
      }
      if (numVal(datos.valor_base_comision) === 0) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'El valor base de la venta debe ser mayor a 0' });
      }

      // Integridad referencial del inmueble
      var inmRefV = leerHoja(HOJAS.inmuebles);
      var cliRefV = leerHoja(HOJAS.clientes);
      var inmVnt = inmRefV.find(function(i){ return String(i.id_inmueble) === String(datos.id_inmueble); });
      if (!inmVnt) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Inmueble "' + datos.id_inmueble + '" no existe' });
      }
      if (String(inmVnt.estado || '').toLowerCase() === 'inactivo') {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'El inmueble "' + datos.id_inmueble + '" está inactivo. Reactívelo antes de registrar una venta.' });
      }

      // Retrocompat: si no vienen partes[] pero sí id_vendedor/id_comprador legacy, convertir
      var partesArrV = body.partes;
      if ((!partesArrV || partesArrV.length === 0) && (datos.id_vendedor || datos.id_comprador)) {
        partesArrV = [];
        if (datos.id_vendedor) partesArrV.push({ rol:'vendedor', id_cliente:datos.id_vendedor, participacion_pct:1 });
        if (datos.id_comprador) partesArrV.push({ rol:'comprador', id_cliente:datos.id_comprador, participacion_pct:1 });
      }

      datos.id_venta = siguienteId(HOJAS.ventas, 'VNT');
      datos.comision_oficina = numVal(datos.valor_base_comision) * numVal(datos.pct_comision_oficina);
      datos.comision_por_punta = datos.comision_oficina / 2;
      datos.estado_venta = 'ACTIVA';

      // Validar y guardar partes (vendedor + comprador, suma=100% por rol)
      var errPartesV = validarYGuardarPartes(
        datos.id_venta, 'venta', ['vendedor','comprador'], partesArrV, cliRefV
      );
      if (errPartesV) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error: errPartesV });
      }

      // Validar comisiones ANTES de escribir la venta
      var errComVnt = validarComisionesAsesores(body.comisiones_asesores, leerHoja(HOJAS.asesores));
      if (errComVnt) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error: errComVnt });
      }

      // Validar coherencia de referidos
      var errRefVnt = validarReferidos(datos);
      if (errRefVnt) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error: errRefVnt });
      }

      agregarFila(HOJAS.ventas, COLUMNAS.ventas, datos);

      if (body.comisiones_asesores && body.comisiones_asesores.length > 0) {
        body.comisiones_asesores.forEach(com => {
          agregarFila(HOJAS.comisiones, COLUMNAS.comisiones, {
            id_asesor: com.id_asesor,
            id_negocio: datos.id_venta,
            valor_comision: com.valor_comision,
            punta: com.punta,
            participacion: (numVal(com.participacion) || 100) / 100,
            estado: 'ACTIVA'
          });
        });
      }

      // Registrar pagos/cuotas
      // El frontend envía valor_pago (monto del inmueble que se paga).
      // Convertimos a valor_cobrado (comisión proporcional que entra a la oficina).
      if (body.pagos && body.pagos.length > 0) {
        var valorBase = numVal(datos.valor_base_comision);
        // Validar que la suma de pagos cuadre con el valor base (tolerancia $1)
        var sumaPagosV = body.pagos.reduce(function(acc, p){ return acc + numVal(p.valor_pago); }, 0);
        if (Math.abs(sumaPagosV - valorBase) > 1) {
          lock.releaseLock();
          return jsonResponse({ ok:false, error:'La suma de pagos ($' + Math.round(sumaPagosV).toLocaleString() + ') no cuadra con el valor base ($' + Math.round(valorBase).toLocaleString() + ')' });
        }
        // Validar que ningún pago sea negativo
        for (var iP = 0; iP < body.pagos.length; iP++) {
          if (numVal(body.pagos[iP].valor_pago) < 0) {
            lock.releaseLock();
            return jsonResponse({ ok:false, error:'Los pagos no pueden ser negativos' });
          }
        }
        body.pagos.forEach(function(pago) {
          var idPago = siguienteId(HOJAS.pagos, 'PAG');
          var fechaPago = pago.fecha_pago ? new Date(pago.fecha_pago + 'T12:00:00') : null;
          var valorPago = numVal(pago.valor_pago);
          var valorComision = valorBase > 0 ? (valorPago / valorBase) * datos.comision_oficina : 0;
          agregarFila(HOJAS.pagos, COLUMNAS.pagos, {
            id_pago: idPago,
            id_venta: datos.id_venta,
            fecha_pago: pago.fecha_pago || '',
            año_pago: fechaPago ? fechaPago.getFullYear() : '',
            mes_pago: fechaPago ? (fechaPago.getMonth() + 1) : '',
            valor_cobrado: Math.round(valorComision),
            observacion: pago.observacion || ''
          });
        });
      }

      invalidarCacheHojas([HOJAS.ventas, HOJAS.pagos, HOJAS.comisiones, HOJAS.partes]);
      lock.releaseLock();
      return jsonResponse({ ok: true, id: datos.id_venta, mensaje: 'Venta registrada' });
    }

    // --- REGISTRAR INMUEBLE ---
    if (action === 'registrar_inmueble') {
      const datos = body.datos;
      var inmExist = leerHoja(HOJAS.inmuebles);

      // Bloquear duplicado por codigo_plataforma (si viene)
      var codigo = String(datos.codigo_plataforma || '').trim();
      if (codigo) {
        var dupCod = inmExist.find(function(i){
          return String(i.codigo_plataforma || '').trim().toLowerCase() === codigo.toLowerCase();
        });
        if (dupCod) {
          lock.releaseLock();
          return jsonResponse({
            ok:false,
            error:'Ya existe un inmueble con código de plataforma "' + codigo + '" (' + dupCod.id_inmueble + ' — ' + (dupCod.nombre || '') + '). Use ese en vez de crear uno nuevo.',
            duplicado: dupCod
          });
        }
      }

      // Bloquear duplicado por nombre normalizado + ciudad
      var nmNorm = normalizarNombre_(datos.nombre);
      var ciudad = String(datos.ciudad || '').trim();
      if (nmNorm) {
        var dupNm = inmExist.find(function(i){
          return normalizarNombre_(i.nombre) === nmNorm && String(i.ciudad || '').trim() === ciudad;
        });
        if (dupNm) {
          lock.releaseLock();
          return jsonResponse({
            ok:false,
            error:'Ya existe un inmueble con ese nombre en ' + ciudad + ' (' + dupNm.id_inmueble + ' — ' + (dupNm.nombre || '') + '). Use ese en vez de crear uno nuevo.',
            duplicado: dupNm
          });
        }
      }

      datos.id_inmueble = siguienteId(HOJAS.inmuebles, 'INM');
      datos.estado = 'Disponible';
      agregarFila(HOJAS.inmuebles, COLUMNAS.inmuebles, datos);
      invalidarCacheHojas([HOJAS.inmuebles]);
      lock.releaseLock();
      return jsonResponse({ ok: true, id: datos.id_inmueble, mensaje: 'Inmueble registrado' });
    }

    // --- REGISTRAR CLIENTE ---
    if (action === 'registrar_cliente') {
      const datos = body.datos;

      // Bloquear duplicado por nombre normalizado
      var nmCliNorm = normalizarNombre_(datos.nombre);
      if (nmCliNorm) {
        var cliExist = leerHoja(HOJAS.clientes);
        var dupCli = cliExist.find(function(c){
          return normalizarNombre_(c.nombre) === nmCliNorm;
        });
        if (dupCli) {
          lock.releaseLock();
          return jsonResponse({
            ok:false,
            error:'Ya existe un cliente con ese nombre (' + dupCli.id_cliente + ' — ' + (dupCli.nombre || '') + '). Use ese en vez de crear uno nuevo.',
            duplicado: dupCli
          });
        }
      }

      datos.id_cliente = siguienteId(HOJAS.clientes, 'CLI');
      agregarFila(HOJAS.clientes, COLUMNAS.clientes, datos);
      invalidarCacheHojas([HOJAS.clientes]);
      lock.releaseLock();
      return jsonResponse({ ok: true, id: datos.id_cliente, mensaje: 'Cliente registrado' });
    }

    // --- IMPRIMIR CUENTA DE COBRO ---
    // body: { action, id_asesor, id_negocio, tipo: 'arriendo'|'venta'|'venta_pago' }
    // tipo 'venta_pago': id_negocio es un id_pago, calcula comisión proporcional al pago
    if (action === 'imprimir_cuenta_cobro') {
      var idAsesor  = body.id_asesor;
      var idNegocio = body.id_negocio;
      var tipoNeg   = body.tipo; // 'arriendo' | 'venta' | 'venta_pago'
      if (!idAsesor || !idNegocio || !tipoNeg) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Faltan parámetros (id_asesor, id_negocio, tipo)' });
      }
      if (!TEMPLATE_CUENTA_COBRO_ID || TEMPLATE_CUENTA_COBRO_ID === 'PEGAR_ID_DEL_GOOGLE_DOC_AQUI') {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Falta configurar TEMPLATE_CUENTA_COBRO_ID en apps_script.js' });
      }

      // Cargar asesor
      var asesores = leerHoja(HOJAS.asesores);
      var asesor = asesores.find(function(a){ return a.id_asesor === idAsesor; });
      if (!asesor) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Asesor no encontrado' }); }

      // Cargar negocio + inmueble
      var negocio, mesNeg, anoNeg, conceptoBase;
      var inmuebles = leerHoja(HOJAS.inmuebles);
      if (tipoNeg === 'arriendo') {
        negocio = leerHoja(HOJAS.arriendos).find(function(a){ return a.id_arriendo === idNegocio; });
        if (!negocio) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Arriendo no encontrado' }); }
        mesNeg = parseInt(negocio.mes,10); anoNeg = negocio['año'];
        var inmA = inmuebles.find(function(i){ return i.id_inmueble === negocio.id_inmueble; });
        conceptoBase = 'Comisión por arriendo del inmueble ' + (inmA ? inmA.nombre : negocio.id_inmueble);
      } else if (tipoNeg === 'venta_pago') {
        // Cuenta de cobro por pago individual
        var pago = leerHoja(HOJAS.pagos).find(function(p){ return p.id_pago === idNegocio; });
        if (!pago) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Pago no encontrado' }); }
        negocio = leerHoja(HOJAS.ventas).find(function(v){ return v.id_venta === pago.id_venta; });
        if (!negocio) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Venta no encontrada' }); }
        idNegocio = negocio.id_venta; // para buscar comisiones
        mesNeg = parseInt(negocio.mes,10); anoNeg = negocio['año'];
        var inmP = inmuebles.find(function(i){ return i.id_inmueble === negocio.id_inmueble; });
        conceptoBase = 'Comisión por venta del inmueble ' + (inmP ? inmP.nombre : negocio.id_inmueble) +
          ' — Cuota ' + pago.id_pago + (pago.observacion ? ' (' + pago.observacion + ')' : '');
      } else {
        negocio = leerHoja(HOJAS.ventas).find(function(v){ return v.id_venta === idNegocio; });
        if (!negocio) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Venta no encontrada' }); }
        mesNeg = parseInt(negocio.mes,10); anoNeg = negocio['año'];
        var inmV = inmuebles.find(function(i){ return i.id_inmueble === negocio.id_inmueble; });
        conceptoBase = 'Comisión por venta del inmueble ' + (inmV ? inmV.nombre : negocio.id_inmueble);
      }

      // Bloquear cuenta de cobro si la venta está CANCELADA
      if ((tipoNeg === 'venta' || tipoNeg === 'venta_pago') &&
          String(negocio.estado_venta).toUpperCase() === 'CANCELADA') {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'No se puede generar cuenta de cobro: la venta está cancelada' });
      }

      // Sumar comisiones del asesor en este negocio (puede tener varias puntas)
      var comisiones = leerHoja(HOJAS.comisiones);
      var misCom = comisiones.filter(function(c){
        return c.id_asesor === idAsesor && c.id_negocio === idNegocio;
      });
      if (misCom.length === 0) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'No hay comisión registrada para este asesor en este negocio' });
      }
      var valorTotal = misCom.reduce(function(s,c){ return s + (Number(c.valor_comision)||0); }, 0);

      // Si es cuenta de cobro por pago, aplicar proporción
      if (tipoNeg === 'venta_pago') {
        var comOficina = Number(negocio.comision_oficina) || 0;
        var fraccion = comOficina > 0 ? (Number(pago.valor_cobrado) || 0) / comOficina : 0;
        valorTotal = valorTotal * fraccion;
      }
      valorTotal = Math.round(valorTotal);

      // Detalle del NEGOCIO COMPLETO (todos los asesores que participaron), para que
      // la dirección comercial vea la totalidad además de la tabla individual del asesor.
      // Para cuenta de cobro por pago, se prorratea con la misma fracción del pago.
      var fraccionCC = 1;
      if (tipoNeg === 'venta_pago') {
        var comOfCC = Number(negocio.comision_oficina) || 0;
        fraccionCC = comOfCC > 0 ? (Number(pago.valor_cobrado) || 0) / comOfCC : 0;
      }
      var comisionesNegocio = comisiones.filter(function(c){
        return c.id_negocio === idNegocio && String(c.estado || '').toUpperCase() !== 'ANULADA';
      });
      var detalleNegocio = comisionesNegocio.map(function(c){
        var a = asesores.find(function(x){ return x.id_asesor === c.id_asesor; });
        return {
          nombre: a ? a.nombre : c.id_asesor,
          punta: c.punta || '',
          participacion: (c.participacion === '' || c.participacion === null || c.participacion === undefined) ? 1 : numVal(c.participacion),
          valor: Math.round((Number(c.valor_comision) || 0) * fraccionCC)
        };
      });
      var totalNegocio = detalleNegocio.reduce(function(s, d){ return s + d.valor; }, 0);

      // Fecha + concepto final
      var hoy = new Date();
      var mesesNom = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
      var fechaTexto = hoy.getDate() + ' de ' + mesesNom[hoy.getMonth()] + ' de ' + hoy.getFullYear();
      var concepto = conceptoBase + (mesNeg ? ' — mes de ' + mesesNom[mesNeg-1] + ' de ' + (anoNeg||hoy.getFullYear()) : '');

      // Reemplazos
      var reemplazos = {
        'ciudad_emision':       CIUDAD_EMISION,
        'fecha_texto':          fechaTexto,
        'empresa_razon_social': EMPRESA_RAZON_SOCIAL,
        'empresa_nit':          EMPRESA_NIT,
        'asesor_nombre':        asesor.nombre || '',
        'asesor_cedula':        asesor.cedula || '',
        'asesor_ciudad_cc':     asesor.ciudad_cc || '',
        'valor_numero':         '$' + Number(valorTotal).toLocaleString('es-CO'),
        'valor_letras':         numeroALetras(valorTotal) + ' pesos M/cte.',
        'concepto':             concepto,
        'asesor_direccion':     asesor.direccion || '',
        'banco':                asesor.banco || '',
        'tipo_cuenta':          asesor.tipo_cuenta || '',
        'numero_cuenta':        asesor.numero_cuenta || ''
      };

      // Copiar template, reemplazar, exportar PDF y enviar. La copia temporal se borra
      // SIEMPRE (finally), aunque falle el PDF o el correo, para no dejar basura en Drive.
      var templateFile = DriveApp.getFileById(TEMPLATE_CUENTA_COBRO_ID);
      var nombreCopia = 'Cuenta de cobro ' + (asesor.nombre||'') + ' - ' + idNegocio;
      var copia = templateFile.makeCopy(nombreCopia);
      try {
        var doc = DocumentApp.openById(copia.getId());
        var docBody = doc.getBody();
        Object.keys(reemplazos).forEach(function(k){
          docBody.replaceText('\\{\\{' + k + '\\}\\}', String(reemplazos[k]));
        });

        // Anexar al documento el detalle del negocio completo: tabla individual por asesor
        // + total del negocio (todos los participantes). Visible para la dirección comercial.
        var fmtCopCC = function(n){ return '$' + Number(Math.round(n)).toLocaleString('es-CO'); };
        docBody.appendParagraph('').setSpacingBefore(12);
        var hdNeg = docBody.appendParagraph('DETALLE DEL NEGOCIO ' + idNegocio);
        hdNeg.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        hdNeg.editAsText().setBold(true);
        var rowsNeg = [['Asesor', 'Punta', 'Participación', 'Comisión']];
        detalleNegocio.forEach(function(d){
          rowsNeg.push([
            d.nombre,
            d.punta,
            Math.round(d.participacion * 100) + '%',
            fmtCopCC(d.valor)
          ]);
        });
        rowsNeg.push(['TOTAL DEL NEGOCIO', '', '', fmtCopCC(totalNegocio)]);
        var tblNeg = docBody.appendTable(rowsNeg);
        tblNeg.getRow(0).editAsText().setBold(true);
        tblNeg.getRow(rowsNeg.length - 1).editAsText().setBold(true);

        doc.saveAndClose();

        var pdfBlob = copia.getAs('application/pdf').setName(nombreCopia + '.pdf');

        // Enviar correo
        var asunto = 'Cuenta de cobro — ' + (asesor.nombre||'') + ' — ' + idNegocio;
        var cuerpo = 'Adjunto cuenta de cobro generada automáticamente por el portal REDA3.\n\n' +
                     'Asesor: ' + (asesor.nombre||'') + '\n' +
                     'Negocio: ' + idNegocio + ' (' + tipoNeg + ')\n' +
                     'Concepto: ' + concepto + '\n' +
                     'Valor del asesor: ' + reemplazos.valor_numero + '\n' +
                     'Total del negocio: $' + Number(totalNegocio).toLocaleString('es-CO') + '\n';
        // Destinatarios: gerente (principal) + dirección comercial + asesor (copia)
        var opts = { attachments: [pdfBlob] };
        var ccStr = ccCobro_(asesor.email);
        if (ccStr) opts.cc = ccStr;
        MailApp.sendEmail(GERENTE_EMAIL, asunto, cuerpo, opts);
      } catch (eEnvioCC) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'No se pudo generar o enviar la cuenta de cobro: ' + eEnvioCC.message + '. Revisa los correos e inténtalo de nuevo.' });
      } finally {
        // Borrar copia temporal del Doc pase lo que pase (el PDF ya fue adjuntado al correo)
        try { copia.setTrashed(true); } catch(_) {}
      }

      lock.releaseLock();
      return jsonResponse({
        ok:true,
        mensaje:'Cuenta de cobro enviada al gerente y a la dirección comercial' + (asesor.email ? ' (con copia a ' + asesor.email + ')' : ''),
        valor: valorTotal
      });
    }

    // --- COBRAR BONIFICACIÓN: PDF cuenta de cobro + Excel liquidación ---
    if (action === 'cobrar_bonificacion') {
      var idAsesorBon = body.id_asesor;
      var mesBon = parseInt(body.mes, 10);
      if (!idAsesorBon || !mesBon) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Faltan parámetros (id_asesor, mes)' }); }

      // Validación: sólo meses cerrados (estricto: mes < mes actual)
      var hoyB = new Date();
      var anoActualB = hoyB.getFullYear();
      var mesActualB = hoyB.getMonth() + 1;
      if (mesBon >= mesActualB) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Sólo puede cobrar la bonificación de meses ya cerrados (desde el día 1 del mes siguiente)' });
      }

      // Cargar asesor
      var asesoresBon = leerHoja(HOJAS.asesores);
      var asesorBon = asesoresBon.find(function(a){ return a.id_asesor === idAsesorBon; });
      if (!asesorBon) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Asesor no encontrado' }); }

      // Calcular bonificación (mismo cálculo que mis_bonificaciones)
      var datosBonC = {
        arriendos: leerHoja(HOJAS.arriendos),
        ventas: leerHoja(HOJAS.ventas),
        pagos: leerHoja(HOJAS.pagos),
        comisiones: leerHoja(HOJAS.comisiones),
        acciones: leerHoja(HOJAS.acciones),
        bonificaciones: leerHoja(HOJAS.bonificaciones)
      };
      var actualB = calcularCategoriaMes(idAsesorBon, mesBon, datosBonC);
      var esContB = false;
      if (mesBon > 1 && actualB.escalon) {
        var antB = calcularCategoriaMes(idAsesorBon, mesBon - 1, datosBonC);
        if (antB.categoria === actualB.categoria) esContB = true;
      }
      var pctVarB = 0, fijoBaseB = 0, varBaseB = 0;
      if (actualB.escalon) {
        pctVarB = esContB ? Number(actualB.escalon.pct_variable_continuidad) || 0 : Number(actualB.escalon.pct_variable_inicial) || 0;
        fijoBaseB = actualB.esMedio ? (Number(actualB.escalon.fijo_medio) || 0) : (Number(actualB.escalon.fijo) || 0);
        varBaseB = actualB.comisionGeneradaOficina * pctVarB;
      } else if (actualB.esPiso4) {
        // PISO 4%: 4% de la comisión generada, sin fijo
        pctVarB = 0.04;
        fijoBaseB = 0;
        varBaseB = actualB.comisionGeneradaOficina * pctVarB;
      }
      var vincB = String(asesorBon.vinculacion || '').toLowerCase();
      var factorB = vincB === 'empleado' ? (1 / 1.3) : 1;
      var fijoB = fijoBaseB * factorB;
      var variableB = varBaseB * factorB;
      var totalB = Math.round(fijoB + variableB);

      if (totalB <= 0) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'La bonificación de ' + mesBon + ' es $0. No hay nada que cobrar.' });
      }

      // Validar que la bonificación no haya sido cobrada antes (anti doble cobro)
      asegurarColumnaCobradaEn();
      var liqsExistentesB = leerHoja(HOJAS.bonificaciones_mes);
      var liqExistenteB = liqsExistentesB.find(function(b){
        return b.id_asesor === idAsesorBon
          && parseInt(b['año'], 10) === anoActualB
          && parseInt(b.mes, 10) === mesBon;
      });
      if (liqExistenteB && liqExistenteB.cobrada_en) {
        var fechaCob = liqExistenteB.cobrada_en;
        var fechaStr = (fechaCob instanceof Date)
          ? fechaCob.toLocaleDateString('es-CO')
          : String(fechaCob);
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Ya cobraste la bonificación de ' + mesBon + '/' + anoActualB + ' el ' + fechaStr });
      }

      // Datos para el Excel
      var inmueblesB = leerHoja(HOJAS.inmuebles);
      var nomInmB = function(id){ var i = inmueblesB.find(function(x){return x.id_inmueble===id;}); return i ? i.nombre : id; };
      var misComB = datosBonC.comisiones.filter(function(c){
        return c.id_asesor === idAsesorBon && String(c.estado||'').toUpperCase() !== 'ANULADA';
      });
      var misNegIdsB = misComB.map(function(c){ return c.id_negocio; });
      var sumPart = function(idNeg) {
        return misComB.filter(function(c){ return c.id_negocio === idNeg; })
          .reduce(function(s,c){ return s + (c.participacion === '' ? 1 : numVal(c.participacion)); }, 0);
      };

      // Cierres de arriendo del mes
      var cierresArrB = datosBonC.arriendos.filter(function(a){
        return parseInt(a.mes,10) === mesBon && misNegIdsB.indexOf(a.id_arriendo) !== -1;
      });
      // Cierres de venta: por mes_pago
      var pagosMesBonC = datosBonC.pagos.filter(function(p){ return parseInt(p.mes_pago,10) === mesBon; });
      var ventasIdsB = pagosMesBonC.map(function(p){ return p.id_venta; });
      var cierresVntB = datosBonC.ventas.filter(function(v){
        return ventasIdsB.indexOf(v.id_venta) !== -1
          && misNegIdsB.indexOf(v.id_venta) !== -1
          && String(v.estado_venta||'').toUpperCase() !== 'CANCELADA';
      });

      // PDF cuenta de cobro con tabla embebida
      var mesesEs = ['','enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
      var nombreMesB = mesesEs[mesBon];
      var fechaTextoB = hoyB.getDate() + ' de ' + mesesEs[hoyB.getMonth()+1] + ' de ' + anoActualB;
      var conceptoB = 'Bonificación correspondiente al mes de ' + nombreMesB + ' de ' + anoActualB;
      var reemplazosB = {
        'ciudad_emision':       CIUDAD_EMISION,
        'fecha_texto':          fechaTextoB,
        'empresa_razon_social': EMPRESA_RAZON_SOCIAL,
        'empresa_nit':          EMPRESA_NIT,
        'asesor_nombre':        asesorBon.nombre || '',
        'asesor_cedula':        asesorBon.cedula || '',
        'asesor_ciudad_cc':     asesorBon.ciudad_cc || '',
        'valor_numero':         '$' + Number(totalB).toLocaleString('es-CO'),
        'valor_letras':         numeroALetras(totalB) + ' pesos M/cte.',
        'concepto':             conceptoB,
        'asesor_direccion':     asesorBon.direccion || '',
        'banco':                asesorBon.banco || '',
        'tipo_cuenta':          asesorBon.tipo_cuenta || '',
        'numero_cuenta':        asesorBon.numero_cuenta || ''
      };
      var templateFileB = DriveApp.getFileById(TEMPLATE_CUENTA_COBRO_ID);
      var nombreCopiaB = 'Cuenta de cobro ' + (asesorBon.nombre||'') + ' - Bonificación ' + nombreMesB + ' ' + anoActualB;
      var copiaB = templateFileB.makeCopy(nombreCopiaB);
      try {
      var docB = DocumentApp.openById(copiaB.getId());
      var docBodyB = docB.getBody();
      Object.keys(reemplazosB).forEach(function(k){
        docBodyB.replaceText('\\{\\{' + k + '\\}\\}', String(reemplazosB[k]));
      });

      // Helper para formato de pesos en la tabla
      var fmtCop = function(n){ return '$' + Number(Math.round(n)).toLocaleString('es-CO'); };

      // Insertar tabla con detalle de negocios + bonificación al final del documento
      docBodyB.appendParagraph('').setSpacingBefore(12);
      var hd1 = docBodyB.appendParagraph('DETALLE DE NEGOCIOS — ' + nombreMesB.toUpperCase() + ' ' + anoActualB);
      hd1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      hd1.editAsText().setBold(true);

      if (cierresVntB.length > 0) {
        var pVnt = docBodyB.appendParagraph('Cierres de venta');
        pVnt.editAsText().setBold(true);
        var rowsVnt = [['Inmueble','Valor venta','% Com.','Comisión oficina','Mi part.','Comisión generada']];
        cierresVntB.forEach(function(v){
          var partV = sumPart(v.id_venta) * 0.5;
          var comGen = numVal(v.comision_oficina) * partV;
          rowsVnt.push([
            nomInmB(v.id_inmueble),
            fmtCop(numVal(v.valor_base_comision)),
            (numVal(v.pct_comision_oficina)*100).toFixed(2) + '%',
            fmtCop(numVal(v.comision_oficina)),
            (partV*100).toFixed(0) + '%',
            fmtCop(comGen)
          ]);
        });
        var tblVnt = docBodyB.appendTable(rowsVnt);
        tblVnt.getRow(0).editAsText().setBold(true);
      }

      if (cierresArrB.length > 0) {
        var pArr = docBodyB.appendParagraph('Cierres de arriendo');
        pArr.editAsText().setBold(true);
        var rowsArr = [['Inmueble','Canon+Admin','% Com.','Meses','Comisión total','Mi part.','Comisión generada']];
        cierresArrB.forEach(function(a){
          var mesesA = mesesContratoDe(a);
          var comTotalA = numVal(a.comision_oficina) * mesesA;
          var partA = sumPart(a.id_arriendo) * 0.5;
          var comGenA = comTotalA * partA;
          rowsArr.push([
            nomInmB(a.id_inmueble),
            fmtCop(numVal(a.valor_canon) + numVal(a.administracion)),
            (numVal(a.pct_comision_oficina)*100).toFixed(2) + '%',
            String(mesesA),
            fmtCop(comTotalA),
            (partA*100).toFixed(0) + '%',
            fmtCop(comGenA)
          ]);
        });
        var tblArr = docBodyB.appendTable(rowsArr);
        tblArr.getRow(0).editAsText().setBold(true);
      }

      // Liquidación de bonificación
      docBodyB.appendParagraph('').setSpacingBefore(8);
      var hd2 = docBodyB.appendParagraph('LIQUIDACIÓN DE BONIFICACIÓN');
      hd2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      hd2.editAsText().setBold(true);

      var rowsLiq = [
        ['Concepto','Valor'],
        ['Comisión generada a la oficina (total)', fmtCop(actualB.comisionGeneradaOficina)],
        ['Acciones comerciales del mes', String(actualB.numAcciones)],
        ['Categoría', actualB.categoria + (actualB.esMedio ? ' (1/2)' : '')],
        ['Vinculación', String(asesorBon.vinculacion || '')],
        ['% Variable aplicado', (pctVarB*100).toFixed(0) + '%'],
        ['Bonificación fija', fmtCop(fijoB)],
        ['Bonificación variable', fmtCop(variableB)]
      ];
      if (vincB === 'empleado') rowsLiq.push(['Factor empleado', 'Total ÷ 1.3']);
      rowsLiq.push(['TOTAL A COBRAR', fmtCop(totalB)]);
      var tblLiq = docBodyB.appendTable(rowsLiq);
      tblLiq.getRow(0).editAsText().setBold(true);
      tblLiq.getRow(rowsLiq.length - 1).editAsText().setBold(true);

      docB.saveAndClose();
      var pdfBlobB = copiaB.getAs('application/pdf').setName(nombreCopiaB + '.pdf');

      // Enviar correo (solo PDF)
      var asuntoB = 'Cuenta de cobro - Bonificación ' + nombreMesB + ' ' + anoActualB + ' - ' + asesorBon.nombre;
      var cuerpoB = 'Adjunto cuenta de cobro de bonificación generada automáticamente por el portal REDA3.\n\n' +
                    'Asesor: ' + asesorBon.nombre + '\n' +
                    'Mes liquidado: ' + nombreMesB + ' ' + anoActualB + '\n' +
                    'Categoría: ' + actualB.categoria + (actualB.esMedio ? ' (1/2)' : '') + '\n' +
                    'TOTAL: $' + Number(totalB).toLocaleString('es-CO') + '\n';
      var optsB = { attachments: [pdfBlobB] };
      var ccStrB = ccCobro_(asesorBon.email);
      if (ccStrB) optsB.cc = ccStrB;
      MailApp.sendEmail(GERENTE_EMAIL, asuntoB, cuerpoB, optsB);
      } catch (eEnvioBon) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'No se pudo generar o enviar la cuenta de cobro de bonificación: ' + eEnvioBon.message + '. Revisa los correos e inténtalo de nuevo.' });
      } finally {
        // Borrar copia temporal del Doc pase lo que pase (el PDF ya fue adjuntado)
        try { copiaB.setTrashed(true); } catch(_) {}
      }

      // Marcar la liquidación como cobrada (o crear fila si el mes no había sido liquidado)
      var ahoraCobro = new Date();
      if (liqExistenteB) {
        actualizarFila(HOJAS.bonificaciones_mes, 'id_bonmes', liqExistenteB.id_bonmes, { cobrada_en: ahoraCobro });
      } else {
        agregarFila(HOJAS.bonificaciones_mes, COLUMNAS.bonificaciones_mes, {
          id_bonmes: siguienteId(HOJAS.bonificaciones_mes, 'BNM'),
          id_asesor: idAsesorBon,
          'año': anoActualB,
          mes: mesBon,
          fecha: new Date(anoActualB, mesBon - 1, 1),
          categoria: actualB.categoria + (actualB.esMedio ? ' (1/2)' : ''),
          comision_generada: actualB.comisionGeneradaOficina,
          acciones_mes: actualB.numAcciones,
          fijo: fijoB,
          pct_variable: pctVarB,
          variable: variableB,
          total: totalB,
          continuidad: actualB.escalon ? (esContB ? 'CONTINUA' : 'INICIAL') : 'N/A',
          calculado_en: ahoraCobro,
          cobrada_en: ahoraCobro
        });
      }

      invalidarCacheHojas([HOJAS.bonificaciones_mes]);
      lock.releaseLock();
      return jsonResponse({
        ok: true,
        mensaje: 'Cuenta de cobro de bonificación enviada al gerente y a la dirección comercial' + (asesorBon.email ? ' (con copia a ' + asesorBon.email + ')' : ''),
        valor: totalB
      });
    }

    // --- REGISTRAR ACCIÓN COMERCIAL ---
    if (action === 'registrar_accion') {
      const datos = body.datos;
      if (!datos.id_asesor) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Falta id_asesor' });
      }
      // Validar asesor existe
      var asesorAcc = leerHoja(HOJAS.asesores).find(function(a){ return a.id_asesor === datos.id_asesor; });
      if (!asesorAcc) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Asesor "' + datos.id_asesor + '" no existe' });
      }
      // Validar fecha si viene
      if (datos.fecha) {
        var fechaP = new Date(datos.fecha);
        if (isNaN(fechaP.getTime())) {
          lock.releaseLock();
          return jsonResponse({ ok:false, error:'Fecha inválida: ' + datos.fecha });
        }
        if (!datos.mes) datos.mes = fechaP.getMonth() + 1;
      } else if (!datos.mes) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Falta fecha o mes de la acción' });
      }
      datos.id_accion = siguienteId(HOJAS.acciones, 'ACC');
      agregarFila(HOJAS.acciones, COLUMNAS.acciones, datos);
      invalidarCacheHojas([HOJAS.acciones]);
      lock.releaseLock();
      return jsonResponse({ ok: true, id: datos.id_accion, mensaje: 'Acción registrada' });
    }

    // --- ACTUALIZAR COBRO DE ARRIENDO (sólo gerente) ---
    if (action === 'actualizar_cobro_arriendo') {
      var idCobroU = body.id_cobro;
      var idAsesorU = body.id_asesor;
      if (!idCobroU || !idAsesorU) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Faltan parámetros (id_cobro, id_asesor)' });
      }
      // Validar rol gerente
      var asesorU = leerHoja(HOJAS.asesores).find(function(a){ return a.id_asesor === idAsesorU; });
      if (!esGestor_(asesorU)) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Sólo gerencia o dirección comercial pueden modificar cobros de arriendo' });
      }
      var cobroExist = leerHoja(HOJAS.cobros_arriendo).find(function(c){ return c.id_cobro === idCobroU; });
      if (!cobroExist) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Cobro no encontrado' });
      }
      var datosUpdCob = {};
      if (body.estado !== undefined) {
        var estU = String(body.estado).toUpperCase();
        if (['COBRADO','NO_COBRADO','CANCELADO'].indexOf(estU) === -1) {
          lock.releaseLock();
          return jsonResponse({ ok:false, error:'Estado inválido. Use COBRADO, NO_COBRADO o CANCELADO' });
        }
        datosUpdCob.estado = estU;
        // Inhabilitar o cancelar borra fecha_pago salvo que el gerente mande una explícita
        if ((estU === 'NO_COBRADO' || estU === 'CANCELADO') && body.fecha_pago === undefined) {
          datosUpdCob.fecha_pago = '';
        }
        // Rehabilitar a COBRADO sin fecha y sin fecha previa → usar hoy
        if (estU === 'COBRADO' && body.fecha_pago === undefined && !cobroExist.fecha_pago) {
          datosUpdCob.fecha_pago = new Date();
        }
      }
      if (body.fecha_pago !== undefined) datosUpdCob.fecha_pago = body.fecha_pago;
      if (body.valor_cobrado !== undefined) datosUpdCob.valor_cobrado = numVal(body.valor_cobrado);
      if (body.observacion !== undefined) datosUpdCob.observacion = body.observacion;

      actualizarFila(HOJAS.cobros_arriendo, 'id_cobro', idCobroU, datosUpdCob);
      invalidarCacheHojas([HOJAS.cobros_arriendo]);
      lock.releaseLock();
      return jsonResponse({ ok:true, mensaje:'Cobro actualizado' });
    }

    // --- SETUP COBROS ARRIENDO (one-shot, sólo gerente) ---
    if (action === 'setup_cobros_arriendo') {
      var idAsesorS = body.id_asesor;
      var asesorS = leerHoja(HOJAS.asesores).find(function(a){ return a.id_asesor === idAsesorS; });
      if (!esGestor_(asesorS)) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Sólo gerencia o dirección comercial pueden ejecutar el setup' });
      }
      var res = setupCobrosArriendo();
      lock.releaseLock();
      return jsonResponse({ ok:true, mensaje:'Setup ejecutado', resultado: res });
    }

    // --- LIQUIDAR BONIFICACIONES DEL MES (sólo gerente) ---
    if (action === 'liquidar_bonificaciones') {
      var idAsesorL = body.id_asesor;
      var asesorL = leerHoja(HOJAS.asesores).find(function(a){ return a.id_asesor === idAsesorL; });
      if (!esGestor_(asesorL)) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Sólo gerencia o dirección comercial pueden liquidar bonificaciones' });
      }
      var anioL = parseInt(body['año'] || body.anio || body.year, 10);
      var mesL = parseInt(body.mes, 10);
      try {
        var resultadoL = liquidarMes(anioL, mesL);
        invalidarCacheHojas([HOJAS.bonificaciones_mes]);
        lock.releaseLock();
        return jsonResponse(resultadoL);
      } catch (eL) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error: eL.message });
      }
    }

    // --- ACTUALIZAR PAGO (sólo gerente) ---
    if (action === 'actualizar_pago') {
      var idPago = body.id_pago;
      var idAsesorP = body.id_asesor;
      if (!idPago) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Falta id_pago' }); }
      if (!idAsesorP) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Falta id_asesor' }); }
      var asesorP = leerHoja(HOJAS.asesores).find(function(a){ return a.id_asesor === idAsesorP; });
      if (!esGestor_(asesorP)) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Sólo gerencia o dirección comercial pueden modificar pagos' });
      }
      // Buscar el pago para obtener id_venta
      var pagoActual = leerHoja(HOJAS.pagos).find(function(p) { return p.id_pago === idPago; });
      if (!pagoActual) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Pago no encontrado' }); }
      // Buscar la venta para calcular comisión proporcional
      var ventaPago = leerHoja(HOJAS.ventas).find(function(v) { return v.id_venta === pagoActual.id_venta; });
      if (!ventaPago) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Venta no encontrada' }); }

      var datosUpdate = {};
      if (body.fecha_pago !== undefined) {
        datosUpdate.fecha_pago = body.fecha_pago;
        var fp = body.fecha_pago ? new Date(body.fecha_pago + 'T12:00:00') : null;
        datosUpdate['año_pago'] = fp ? fp.getFullYear() : '';
        datosUpdate.mes_pago = fp ? (fp.getMonth() + 1) : '';
      }
      if (body.valor_pago !== undefined) {
        var valorBase = numVal(ventaPago.valor_base_comision);
        var comOficina = numVal(ventaPago.comision_oficina);
        var valorPago = numVal(body.valor_pago);
        if (valorPago < 0) {
          lock.releaseLock();
          return jsonResponse({ ok:false, error:'El valor del pago no puede ser negativo' });
        }
        if (valorBase > 0 && valorPago > valorBase) {
          lock.releaseLock();
          return jsonResponse({ ok:false, error:'El valor del pago (' + valorPago + ') excede el valor base de la venta (' + valorBase + ')' });
        }
        // Validar que la suma total de pagos (incluyendo el editado) no supere valorBase
        var todosPagos = leerHoja(HOJAS.pagos).filter(function(p){ return p.id_venta === pagoActual.id_venta; });
        var valorCobradoNuevo = valorBase > 0 ? Math.round((valorPago / valorBase) * comOficina) : 0;
        var sumCobradoFuturo = todosPagos.reduce(function(acc, p){
          var vc = numVal(p.valor_cobrado);
          return acc + (p.id_pago === idPago ? valorCobradoNuevo : vc);
        }, 0);
        if (comOficina > 0 && (sumCobradoFuturo - comOficina) > 1) {
          lock.releaseLock();
          return jsonResponse({ ok:false, error:'La suma de pagos de esta venta superaría la comisión total ($' + Math.round(sumCobradoFuturo).toLocaleString() + ' vs $' + Math.round(comOficina).toLocaleString() + ')' });
        }
        datosUpdate.valor_cobrado = valorCobradoNuevo;
      }
      if (body.observacion !== undefined) datosUpdate.observacion = body.observacion;

      actualizarFila(HOJAS.pagos, 'id_pago', idPago, datosUpdate);
      invalidarCacheHojas([HOJAS.pagos]);
      lock.releaseLock();
      return jsonResponse({ ok:true, mensaje:'Pago actualizado' });
    }

    // --- CANCELAR VENTA (sólo gerente) ---
    if (action === 'cancelar_venta') {
      var idVentaCancel = body.id_venta;
      var idAsesorCV = body.id_asesor;
      if (!idVentaCancel) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Falta id_venta' }); }
      if (!idAsesorCV) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Falta id_asesor' }); }
      var asesorCV = leerHoja(HOJAS.asesores).find(function(a){ return a.id_asesor === idAsesorCV; });
      if (!esGestor_(asesorCV)) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Sólo gerencia o dirección comercial pueden cancelar ventas' });
      }

      // Verificar que la venta exista y no esté ya cancelada
      var ventaC = leerHoja(HOJAS.ventas).find(function(v){ return v.id_venta === idVentaCancel; });
      if (!ventaC) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Venta "' + idVentaCancel + '" no encontrada' }); }
      if (String(ventaC.estado_venta).toUpperCase() === 'CANCELADA') {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'La venta ya estaba cancelada' });
      }

      actualizarFila(HOJAS.ventas, 'id_venta', idVentaCancel, { estado_venta: 'CANCELADA' });

      // Marcar todas las comisiones relacionadas como ANULADA (auditable, no se borran)
      var comSheet = getSheet(HOJAS.comisiones);
      var comData = comSheet.getDataRange().getValues();
      var comHeaders = comData[0];
      var idxNeg = comHeaders.indexOf('id_negocio');
      var idxEstado = comHeaders.indexOf('estado');
      var anuladas = 0;
      if (idxNeg !== -1 && idxEstado !== -1) {
        for (var r = 1; r < comData.length; r++) {
          if (String(comData[r][idxNeg]) === String(idVentaCancel)) {
            comSheet.getRange(r + 1, idxEstado + 1).setValue('ANULADA');
            anuladas++;
          }
        }
      }

      // Marcar todos los pagos de la venta cancelada como ANULADO
      // (evita pagos huerfanos apuntando a venta CANCELADA)
      asegurarColumnaEstadoPagos();
      var pagSheet = getSheet(HOJAS.pagos);
      var pagData = pagSheet.getDataRange().getValues();
      var pagHeaders = pagData[0];
      var idxPagVenta = pagHeaders.indexOf('id_venta');
      var idxPagEstado = pagHeaders.indexOf('estado');
      var pagosAnulados = 0;
      if (idxPagVenta !== -1 && idxPagEstado !== -1) {
        for (var p = 1; p < pagData.length; p++) {
          if (String(pagData[p][idxPagVenta]) === String(idVentaCancel)) {
            pagSheet.getRange(p + 1, idxPagEstado + 1).setValue('ANULADO');
            pagosAnulados++;
          }
        }
      }

      invalidarCacheHojas([HOJAS.ventas, HOJAS.comisiones, HOJAS.pagos]);
      lock.releaseLock();
      return jsonResponse({
        ok: true,
        mensaje: 'Venta cancelada. ' + anuladas + ' comision(es) anuladas, ' + pagosAnulados + ' pago(s) anulado(s).'
      });
    }

    // --- CANCELAR ARRIENDO (asesor dueño o gerente) ---
    // Anula: el arriendo, sus comisiones, y todos los cobros del contrato.
    if (action === 'cancelar_arriendo') {
      var idArrCancel = body.id_arriendo;
      var idAsesorCA = body.id_asesor;
      if (!idArrCancel) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Falta id_arriendo' }); }
      if (!idAsesorCA) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Falta id_asesor' }); }

      var asesorCA = leerHoja(HOJAS.asesores).find(function(a){ return a.id_asesor === idAsesorCA; });
      if (!asesorCA) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Asesor no encontrado' }); }
      var esGerenteCA = esGestor_(asesorCA);

      asegurarColumnaEstadoArriendo();
      var arrC = leerHoja(HOJAS.arriendos).find(function(a){ return a.id_arriendo === idArrCancel; });
      if (!arrC) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Arriendo "' + idArrCancel + '" no encontrado' }); }
      if (String(arrC.estado_arriendo || '').toUpperCase() === 'CANCELADO') {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'El arriendo ya estaba cancelado' });
      }

      // Si no es gerente, verificar que el asesor figure en las comisiones del arriendo
      if (!esGerenteCA) {
        var comisionesArr = leerHoja(HOJAS.comisiones).filter(function(c){
          return String(c.id_negocio) === String(idArrCancel);
        });
        var esSuyo = comisionesArr.some(function(c){ return c.id_asesor === idAsesorCA; });
        if (!esSuyo) {
          lock.releaseLock();
          return jsonResponse({ ok:false, error:'Solo puede cancelar arriendos en los que figura como asesor' });
        }
      }

      actualizarFila(HOJAS.arriendos, 'id_arriendo', idArrCancel, { estado_arriendo: 'CANCELADO' });

      // Anular todas las comisiones del arriendo
      var comSheetA = getSheet(HOJAS.comisiones);
      var comDataA = comSheetA.getDataRange().getValues();
      var comHeadA = comDataA[0];
      var idxNegA = comHeadA.indexOf('id_negocio');
      var idxEstadoA = comHeadA.indexOf('estado');
      var anuladasA = 0;
      if (idxNegA !== -1 && idxEstadoA !== -1) {
        for (var rA = 1; rA < comDataA.length; rA++) {
          if (String(comDataA[rA][idxNegA]) === String(idArrCancel)) {
            comSheetA.getRange(rA + 1, idxEstadoA + 1).setValue('ANULADA');
            anuladasA++;
          }
        }
      }

      // Anular todos los cobros del arriendo (pendientes y ya cobrados)
      var cobSheet = getSheet(HOJAS.cobros_arriendo);
      var cobAnulados = 0;
      if (cobSheet) {
        var cobData = cobSheet.getDataRange().getValues();
        var cobHead = cobData[0];
        var idxCobArr = cobHead.indexOf('id_arriendo');
        var idxCobEstado = cobHead.indexOf('estado');
        if (idxCobArr !== -1 && idxCobEstado !== -1) {
          for (var cR = 1; cR < cobData.length; cR++) {
            if (String(cobData[cR][idxCobArr]) === String(idArrCancel)) {
              cobSheet.getRange(cR + 1, idxCobEstado + 1).setValue('ANULADO');
              cobAnulados++;
            }
          }
        }
      }

      invalidarCacheHojas([HOJAS.arriendos, HOJAS.comisiones, HOJAS.cobros_arriendo]);
      lock.releaseLock();
      return jsonResponse({
        ok: true,
        mensaje: 'Arriendo cancelado. ' + anuladasA + ' comisión(es) anuladas, ' + cobAnulados + ' cobro(s) anulado(s).'
      });
    }

    // --- EDITAR ARRIENDO (gestor: gerente o directora) ---
    // Reescribe en sitio el arriendo, sus partes, comisiones y regenera los cobros proyectados.
    // body: { id_asesor, password, id_arriendo, datos, partes, comisiones_asesores }
    if (action === 'editar_arriendo') {
      var asesorEA = leerHoja(HOJAS.asesores).find(function(a){ return a.id_asesor === body.id_asesor; });
      if (!esGestor_(asesorEA)) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Sólo gerencia o dirección comercial pueden editar negocios' });
      }
      var idArrE = body.id_arriendo;
      if (!idArrE) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Falta id_arriendo' }); }
      var arrE = leerHoja(HOJAS.arriendos).find(function(a){ return a.id_arriendo === idArrE; });
      if (!arrE) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Arriendo "' + idArrE + '" no encontrado' }); }
      if (String(arrE.estado_arriendo || '').toUpperCase() === 'CANCELADO') {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'No se puede editar un arriendo cancelado' });
      }
      if (negocioBloqueadoPorCobro_(idArrE, 'arriendo')) {
        lock.releaseLock();
        return jsonResponse({ ok:false, requiere_cancelacion:true,
          error:'Este arriendo ya tiene bonificaciones cobradas de su mes; editarlo desincronizaría cuentas de cobro ya emitidas. Si está mal, cancélalo.' });
      }

      var datosEA = body.datos;
      if (!datosEA['año']) datosEA['año'] = arrE['año'] || new Date().getFullYear();
      if (numVal(datosEA.valor_canon) < 0 || numVal(datosEA.administracion) < 0 || numVal(datosEA.pct_comision_oficina) < 0) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Canon, administración y porcentaje no pueden ser negativos' });
      }
      if (numVal(datosEA.valor_canon) === 0) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'El valor del canon debe ser mayor a 0' });
      }
      var inmRefEA = leerHoja(HOJAS.inmuebles);
      var cliRefEA = leerHoja(HOJAS.clientes);
      var inmEA = inmRefEA.find(function(i){ return String(i.id_inmueble) === String(datosEA.id_inmueble); });
      if (!inmEA) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Inmueble "' + datosEA.id_inmueble + '" no existe' });
      }
      if (String(inmEA.estado || '').toLowerCase() === 'inactivo') {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'El inmueble "' + datosEA.id_inmueble + '" está inactivo. Reactívelo antes de editar el negocio.' });
      }
      // Duplicado: mismo inmueble + mes + año, excluyendo este mismo arriendo
      var dupEA = leerHoja(HOJAS.arriendos).find(function(a){
        return a.id_arriendo !== idArrE
          && String(a.id_inmueble) === String(datosEA.id_inmueble)
          && String(a.mes) === String(datosEA.mes)
          && String(a['año']) === String(datosEA['año']);
      });
      if (dupEA) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Ya existe otro arriendo para ese inmueble en ' + datosEA.mes + '/' + datosEA['año'] + ' (' + dupEA.id_arriendo + ')' });
      }
      var mesesEA = parseInt(datosEA.meses_contrato, 10);
      if (!mesesEA || mesesEA <= 0) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'meses_contrato es obligatorio y debe ser un entero positivo' });
      }
      datosEA.meses_contrato = mesesEA;

      // Validaciones (sin escribir nada todavía)
      var errComEA = validarComisionesAsesores(body.comisiones_asesores, leerHoja(HOJAS.asesores));
      if (errComEA) { lock.releaseLock(); return jsonResponse({ ok:false, error: errComEA }); }
      var errRefEA = validarReferidos(datosEA);
      if (errRefEA) { lock.releaseLock(); return jsonResponse({ ok:false, error: errRefEA }); }
      var errPartesEA = validarYGuardarPartes(idArrE, 'arriendo', ['arrendador','arrendatario'], body.partes, cliRefEA, false);
      if (errPartesEA) { lock.releaseLock(); return jsonResponse({ ok:false, error: errPartesEA }); }

      // Aplicar cambios
      datosEA.id_arriendo = idArrE;
      var canonTotalEA = numVal(datosEA.valor_canon) + numVal(datosEA.administracion);
      datosEA.comision_oficina = canonTotalEA * numVal(datosEA.pct_comision_oficina);
      actualizarFila(HOJAS.arriendos, 'id_arriendo', idArrE, datosEA);

      borrarFilasPorColumna_(HOJAS.partes, 'id_negocio', idArrE);
      validarYGuardarPartes(idArrE, 'arriendo', ['arrendador','arrendatario'], body.partes, cliRefEA, true);

      borrarFilasPorColumna_(HOJAS.comisiones, 'id_negocio', idArrE);
      if (body.comisiones_asesores && body.comisiones_asesores.length > 0) {
        body.comisiones_asesores.forEach(function(com){
          agregarFila(HOJAS.comisiones, COLUMNAS.comisiones, {
            id_asesor: com.id_asesor, id_negocio: idArrE, valor_comision: com.valor_comision,
            punta: com.punta, participacion: (numVal(com.participacion) || 100) / 100, estado: 'ACTIVA'
          });
        });
      }

      // Regenerar cobros proyectados con los nuevos valores (se pierden ediciones manuales de cobros)
      borrarFilasPorColumna_(HOJAS.cobros_arriendo, 'id_arriendo', idArrE);
      try { generarCobrosProyectados(datosEA); } catch (eCobEA) { /* no bloquear */ }

      invalidarCacheHojas([HOJAS.arriendos, HOJAS.comisiones, HOJAS.partes, HOJAS.cobros_arriendo]);
      lock.releaseLock();
      return jsonResponse({ ok:true, id: idArrE, mensaje:'Arriendo actualizado' });
    }

    // --- EDITAR VENTA (gestor: gerente o directora) ---
    // Reescribe en sitio la venta, sus partes, comisiones y plan de pagos.
    // body: { id_asesor, password, id_venta, datos, partes, comisiones_asesores, pagos }
    if (action === 'editar_venta') {
      var asesorEV = leerHoja(HOJAS.asesores).find(function(a){ return a.id_asesor === body.id_asesor; });
      if (!esGestor_(asesorEV)) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Sólo gerencia o dirección comercial pueden editar negocios' });
      }
      var idVntE = body.id_venta;
      if (!idVntE) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Falta id_venta' }); }
      var vntE = leerHoja(HOJAS.ventas).find(function(v){ return v.id_venta === idVntE; });
      if (!vntE) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Venta "' + idVntE + '" no encontrada' }); }
      if (String(vntE.estado_venta || '').toUpperCase() === 'CANCELADA') {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'No se puede editar una venta cancelada' });
      }
      if (negocioBloqueadoPorCobro_(idVntE, 'venta')) {
        lock.releaseLock();
        return jsonResponse({ ok:false, requiere_cancelacion:true,
          error:'Esta venta ya tiene bonificaciones cobradas de un mes de pago; editarla desincronizaría cuentas de cobro ya emitidas. Si está mal, cancélala.' });
      }

      var datosEV = body.datos;
      if (!datosEV['año']) datosEV['año'] = vntE['año'] || new Date().getFullYear();
      if (numVal(datosEV.valor_base_comision) < 0 || numVal(datosEV.pct_comision_oficina) < 0) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Valor base y porcentaje no pueden ser negativos' });
      }
      if (numVal(datosEV.valor_base_comision) === 0) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'El valor base de la venta debe ser mayor a 0' });
      }
      var inmRefEV = leerHoja(HOJAS.inmuebles);
      var cliRefEV = leerHoja(HOJAS.clientes);
      var inmEV = inmRefEV.find(function(i){ return String(i.id_inmueble) === String(datosEV.id_inmueble); });
      if (!inmEV) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Inmueble "' + datosEV.id_inmueble + '" no existe' });
      }
      if (String(inmEV.estado || '').toLowerCase() === 'inactivo') {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'El inmueble "' + datosEV.id_inmueble + '" está inactivo. Reactívelo antes de editar el negocio.' });
      }

      // Validaciones (sin escribir todavía)
      var errComEV = validarComisionesAsesores(body.comisiones_asesores, leerHoja(HOJAS.asesores));
      if (errComEV) { lock.releaseLock(); return jsonResponse({ ok:false, error: errComEV }); }
      var errRefEV = validarReferidos(datosEV);
      if (errRefEV) { lock.releaseLock(); return jsonResponse({ ok:false, error: errRefEV }); }
      var errPartesEV = validarYGuardarPartes(idVntE, 'venta', ['vendedor','comprador'], body.partes, cliRefEV, false);
      if (errPartesEV) { lock.releaseLock(); return jsonResponse({ ok:false, error: errPartesEV }); }

      var valorBaseEV = numVal(datosEV.valor_base_comision);
      if (body.pagos && body.pagos.length > 0) {
        var sumaPagosEV = body.pagos.reduce(function(acc, p){ return acc + numVal(p.valor_pago); }, 0);
        if (Math.abs(sumaPagosEV - valorBaseEV) > 1) {
          lock.releaseLock();
          return jsonResponse({ ok:false, error:'La suma de pagos ($' + Math.round(sumaPagosEV).toLocaleString() + ') no cuadra con el valor base ($' + Math.round(valorBaseEV).toLocaleString() + ')' });
        }
        for (var iPE = 0; iPE < body.pagos.length; iPE++) {
          if (numVal(body.pagos[iPE].valor_pago) < 0) {
            lock.releaseLock();
            return jsonResponse({ ok:false, error:'Los pagos no pueden ser negativos' });
          }
        }
      }

      // Aplicar cambios
      datosEV.id_venta = idVntE;
      datosEV.comision_oficina = valorBaseEV * numVal(datosEV.pct_comision_oficina);
      datosEV.comision_por_punta = datosEV.comision_oficina / 2;
      actualizarFila(HOJAS.ventas, 'id_venta', idVntE, datosEV);

      borrarFilasPorColumna_(HOJAS.partes, 'id_negocio', idVntE);
      validarYGuardarPartes(idVntE, 'venta', ['vendedor','comprador'], body.partes, cliRefEV, true);

      borrarFilasPorColumna_(HOJAS.comisiones, 'id_negocio', idVntE);
      if (body.comisiones_asesores && body.comisiones_asesores.length > 0) {
        body.comisiones_asesores.forEach(function(com){
          agregarFila(HOJAS.comisiones, COLUMNAS.comisiones, {
            id_asesor: com.id_asesor, id_negocio: idVntE, valor_comision: com.valor_comision,
            punta: com.punta, participacion: (numVal(com.participacion) || 100) / 100, estado: 'ACTIVA'
          });
        });
      }

      // Regenerar plan de pagos con los nuevos valores
      borrarFilasPorColumna_(HOJAS.pagos, 'id_venta', idVntE);
      if (body.pagos && body.pagos.length > 0) {
        body.pagos.forEach(function(pago){
          var fechaPagoEV = pago.fecha_pago ? new Date(pago.fecha_pago + 'T12:00:00') : null;
          var valorPagoEV = numVal(pago.valor_pago);
          var valorComEV = valorBaseEV > 0 ? (valorPagoEV / valorBaseEV) * datosEV.comision_oficina : 0;
          agregarFila(HOJAS.pagos, COLUMNAS.pagos, {
            id_pago: siguienteId(HOJAS.pagos, 'PAG'),
            id_venta: idVntE,
            fecha_pago: pago.fecha_pago || '',
            'año_pago': fechaPagoEV ? fechaPagoEV.getFullYear() : '',
            mes_pago: fechaPagoEV ? (fechaPagoEV.getMonth() + 1) : '',
            valor_cobrado: Math.round(valorComEV),
            observacion: pago.observacion || ''
          });
        });
      }

      invalidarCacheHojas([HOJAS.ventas, HOJAS.pagos, HOJAS.comisiones, HOJAS.partes]);
      lock.releaseLock();
      return jsonResponse({ ok:true, id: idVntE, mensaje:'Venta actualizada' });
    }

    // --- ELIMINAR NEGOCIO (gestor: gerente o directora) ---
    // Borrado REAL: elimina el negocio y todas sus comisiones, partes y pagos/cobros,
    // pero SÓLO si nada se ha cobrado (ninguna bonificación cobrada lo toca). Si ya hubo
    // cobro, devuelve requiere_cancelacion para que el frontend ofrezca cancelar (soft).
    // body: { id_asesor, password, id_negocio, tipo: 'arriendo' | 'venta' }
    if (action === 'eliminar_negocio') {
      var asesorEL = leerHoja(HOJAS.asesores).find(function(a){ return a.id_asesor === body.id_asesor; });
      if (!esGestor_(asesorEL)) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Sólo gerencia o dirección comercial pueden eliminar negocios' });
      }
      var idNegEL = body.id_negocio;
      var tipoEL = body.tipo;
      if (!idNegEL || !tipoEL) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Faltan parámetros (id_negocio, tipo)' }); }
      if (tipoEL !== 'arriendo' && tipoEL !== 'venta') { lock.releaseLock(); return jsonResponse({ ok:false, error:'Tipo inválido (use arriendo o venta)' }); }

      if (negocioBloqueadoPorCobro_(idNegEL, tipoEL)) {
        lock.releaseLock();
        return jsonResponse({ ok:false, requiere_cancelacion:true,
          error:'Este negocio ya tiene bonificaciones cobradas de su mes. No se puede borrar sin desincronizar cuentas de cobro. Usa "Cancelar" para anularlo dejando rastro.' });
      }

      var borrado;
      if (tipoEL === 'arriendo') {
        borrado = borrarFilasPorColumna_(HOJAS.arriendos, 'id_arriendo', idNegEL);
        if (!borrado) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Arriendo "' + idNegEL + '" no encontrado' }); }
        borrarFilasPorColumna_(HOJAS.cobros_arriendo, 'id_arriendo', idNegEL);
      } else {
        borrado = borrarFilasPorColumna_(HOJAS.ventas, 'id_venta', idNegEL);
        if (!borrado) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Venta "' + idNegEL + '" no encontrada' }); }
        borrarFilasPorColumna_(HOJAS.pagos, 'id_venta', idNegEL);
      }
      var comBorradas = borrarFilasPorColumna_(HOJAS.comisiones, 'id_negocio', idNegEL);
      var parBorradas = borrarFilasPorColumna_(HOJAS.partes, 'id_negocio', idNegEL);

      invalidarCacheHojas([HOJAS.arriendos, HOJAS.ventas, HOJAS.pagos, HOJAS.cobros_arriendo, HOJAS.comisiones, HOJAS.partes]);
      lock.releaseLock();
      return jsonResponse({ ok:true, mensaje:'Negocio ' + idNegEL + ' eliminado por completo (' + comBorradas + ' comisión(es), ' + parBorradas + ' parte(s) y sus pagos/cobros).' });
    }

    lock.releaseLock();
    return jsonResponse({ ok: false, error: 'Acción POST no reconocida: ' + action });

  } catch (err) {
    lock.releaseLock();
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ===== NÚMERO A LETRAS (español, enteros positivos hasta miles de millones) =====
function numeroALetras(num) {
  num = Math.floor(Math.abs(Number(num) || 0));
  if (num === 0) return 'Cero';
  var UNI = ['','uno','dos','tres','cuatro','cinco','seis','siete','ocho','nueve','diez',
             'once','doce','trece','catorce','quince','dieciséis','diecisiete','dieciocho','diecinueve',
             'veinte','veintiuno','veintidós','veintitrés','veinticuatro','veinticinco','veintiséis','veintisiete','veintiocho','veintinueve'];
  var DEC = ['','','','treinta','cuarenta','cincuenta','sesenta','setenta','ochenta','noventa'];
  var CEN = ['','ciento','doscientos','trescientos','cuatrocientos','quinientos','seiscientos','setecientos','ochocientos','novecientos'];

  function menor1000(n) {
    if (n === 0) return '';
    if (n === 100) return 'cien';
    var c = Math.floor(n/100), r = n%100;
    var s = '';
    if (c) s += CEN[c];
    if (r) {
      if (s) s += ' ';
      if (r < 30) s += UNI[r];
      else {
        var d = Math.floor(r/10), u = r%10;
        s += DEC[d] + (u ? ' y ' + UNI[u] : '');
      }
    }
    return s;
  }

  var partes = [];
  var millones = Math.floor(num / 1000000);
  var miles    = Math.floor((num % 1000000) / 1000);
  var resto    = num % 1000;

  if (millones) {
    partes.push(millones === 1 ? 'un millón' : menor1000(millones) + ' millones');
  }
  if (miles) {
    partes.push(miles === 1 ? 'mil' : menor1000(miles) + ' mil');
  }
  if (resto) {
    partes.push(menor1000(resto));
  }
  var txt = partes.join(' ').trim();
  // Capitalizar primera letra
  return txt.charAt(0).toUpperCase() + txt.slice(1);
}

// ===== INICIALIZAR HOJAS =====
// Ejecuta esta función UNA VEZ para crear las hojas con encabezados
function inicializarHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HOJAS).forEach(key => {
    const nombre = HOJAS[key];
    let sheet = ss.getSheetByName(nombre);
    if (!sheet) {
      sheet = ss.insertSheet(nombre);
    }
    // Poner encabezados si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLUMNAS[key]);
      // Formato encabezados
      const rango = sheet.getRange(1, 1, 1, COLUMNAS[key].length);
      rango.setFontWeight('bold');
      rango.setBackground('#1a3a5c');
      rango.setFontColor('#ffffff');
    }
  });
  SpreadsheetApp.getUi().alert('Hojas inicializadas correctamente');
}
