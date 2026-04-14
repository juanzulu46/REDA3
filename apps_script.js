// =====================================================================
// REDA3 — Google Apps Script (Backend)
// =====================================================================
// INSTRUCCIONES:
// 1. Abre tu Google Sheet "Negocios"
// 2. Ve a Extensiones → Apps Script
// 3. Borra todo el contenido de Code.gs y pega este archivo completo
// 4. Guarda (Ctrl+S)
// 5. Implementar → Nueva implementación → App web
//    - Ejecutar como: Yo
//    - Quién tiene acceso: Cualquier persona
// 6. Copia la URL generada y pégala en index.html (variable SCRIPT_URL)
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
  oficina: 'Oficina',
  origen: 'Origen',
  zona: 'Zona',
  acciones: 'Acciones',
  tipos_accion: 'TipoAccion',
  bonificaciones: 'Bonificaciones',
  parametros: 'Parametros'
};

// Columnas de cada hoja (en orden exacto)
const COLUMNAS = {
  asesores: ['id_asesor', 'nombre', 'vinculacion', 'estado',
             'cedula', 'ciudad_cc', 'direccion', 'banco', 'tipo_cuenta', 'numero_cuenta', 'email',
             'password', 'rol'],
  inmuebles: ['id_inmueble', 'codigo_plataforma', 'nombre', 'ciudad', 'zona', 'tipo', 'residencial_comercial', 'estado'],
  clientes: ['id_cliente', 'nombre', 'telefono', 'email'],
  arriendos: ['id_arriendo', 'año', 'mes', 'mercado', 'id_inmueble', 'id_arrendador', 'id_arrendatario',
              'valor_canon', 'administracion', 'pct_comision_oficina', 'comision_oficina',
              'oficina_captacion', 'origen_captacion', 'oficina_cierre', 'origen_cierre',
              'referido_captador', 'numero_captador_r', 'valor_ref_captador',
              'referido_cerrador', 'numero_cerrador_r', 'valor_ref_cerrador'],
  ventas: ['id_venta', 'año', 'mes', 'mercado', 'id_inmueble', 'id_vendedor', 'id_comprador',
           'valor_base_comision', 'pct_comision_oficina', 'comision_oficina',
           'comision_por_punta',
           'oficina_captacion', 'origen_captacion', 'oficina_cierre', 'origen_cierre',
           'referido_captador', 'numero_captador_r', 'valor_ref_captador',
           'referido_cerrador', 'numero_cerrador_r', 'valor_ref_cerrador',
           'estado_venta'],
  pagos: ['id_pago', 'id_venta', 'fecha_pago', 'año_pago', 'mes_pago', 'valor_cobrado', 'observacion'],
  comisiones: ['id_asesor', 'id_negocio', 'valor_comision', 'punta', 'participacion', 'estado'],
  oficina: ['id_oficina', 'nombre'],
  origen: ['id_origen', 'nombre', 'circulo'],
  zona: ['id_zona', 'comuna', 'ciudad'],
  acciones: ['id_accion', 'id_asesor', 'fecha', 'mes', 'tipo', 'descripcion'],
  tipos_accion: ['id_tipo', 'nombre', 'activo']
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

  // --- Arriendos: sin cambios, filtran por mes del arriendo ---
  arriendos.forEach(function(a) {
    if (parseInt(a.mes, 10) !== mes) return;
    if (negociosIds.indexOf(a.id_arriendo) === -1) return;
    comisionGeneradaOficina += (Number(a.comision_oficina) || 0) * 0.5 * sumParticipacion(a.id_arriendo);
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

  // Si aún no cae en ninguno: ARENA (cumple acciones mínimas pero no comisión) o ARENA MOVEDIZA
  var categoria;
  if (escalonAsignado) {
    categoria = String(escalonAsignado.categoria).toUpperCase();
  } else {
    // Buscar fila ARENA en la tabla para usar su min_acciones
    var arenaRow = escalones.find(function(e) { return String(e.categoria).toUpperCase() === 'ARENA'; });
    var arenaMin = arenaRow ? (Number(arenaRow.min_acciones) || 5) : 5;
    if (numAcciones >= arenaMin) {
      categoria = 'ARENA';
      escalonAsignado = arenaRow || null;
    } else {
      categoria = 'ARENA MOVEDIZA';
      escalonAsignado = null;
    }
  }

  return {
    categoria: categoria,
    esMedio: esMedio,
    escalon: escalonAsignado,
    comisionGeneradaOficina: comisionGeneradaOficina,
    totalRecibido: totalRecibido,
    numAcciones: numAcciones
  };
}

// Respuesta JSON con CORS
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== ENDPOINTS =====

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

    // --- LOGIN: devuelve SOLO lista para el dropdown (sin passwords ni datos sensibles) ---
    if (action === 'login') {
      var asesores = leerHoja(HOJAS.asesores);
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
      return jsonResponse({
        ok: true,
        inmuebles: leerHoja(HOJAS.inmuebles),
        clientes: leerHoja(HOJAS.clientes),
        oficinas: leerHoja(HOJAS.oficina),
        origenes: leerHoja(HOJAS.origen),
        zonas: leerHoja(HOJAS.zona),
        tipos_accion: leerHoja(HOJAS.tipos_accion)
      });
    }

    // --- MIS NEGOCIOS: todo lo del asesor ---
    if (action === 'mis_negocios') {
      var idAsesor = params.id_asesor || '';
      if (!idAsesor) return jsonResponse({ ok: false, error: 'Falta id_asesor' });

      var arriendos = leerHoja(HOJAS.arriendos);
      var ventas = leerHoja(HOJAS.ventas);
      var pagos = leerHoja(HOJAS.pagos);
      var comisiones = leerHoja(HOJAS.comisiones);
      var inmuebles = leerHoja(HOJAS.inmuebles);
      var clientes = leerHoja(HOJAS.clientes);

      var misComisiones = comisiones.filter(function(c) { return c.id_asesor === idAsesor; });
      var misNegocioIds = misComisiones.map(function(c) { return c.id_negocio; });

      var misArriendos = arriendos.filter(function(a) { return misNegocioIds.indexOf(a.id_arriendo) !== -1; });
      var misVentas = ventas.filter(function(v) { return misNegocioIds.indexOf(v.id_venta) !== -1; });
      var misVentaIds = misVentas.map(function(v) { return v.id_venta; });
      var misPagos = pagos.filter(function(p) { return misVentaIds.indexOf(p.id_venta) !== -1; });

      return jsonResponse({
        ok: true,
        arriendos: misArriendos,
        ventas: misVentas,
        pagos: misPagos,
        comisiones: misComisiones,
        inmuebles: inmuebles,
        clientes: clientes
      });
    }

    // --- SIGUIENTE ID ---
    if (action === 'siguiente_id') {
      var tipo = params.tipo || '';
      var prefijos = { arriendos: 'ARR', ventas: 'VNT', inmuebles: 'INM', clientes: 'CLI', pagos: 'PAG' };
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
        arriendos: leerHoja(HOJAS.arriendos),
        ventas: leerHoja(HOJAS.ventas),
        pagos: leerHoja(HOJAS.pagos),
        comisiones: leerHoja(HOJAS.comisiones),
        acciones: leerHoja(HOJAS.acciones),
        bonificaciones: leerHoja(HOJAS.bonificaciones)
      };

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
      var fijo = 0;
      var variable = 0;
      if (actual.escalon) {
        pctVariable = esContinuidad
          ? Number(actual.escalon.pct_variable_continuidad) || 0
          : Number(actual.escalon.pct_variable_inicial) || 0;
        fijo = actual.esMedio
          ? (Number(actual.escalon.fijo_medio) || 0)
          : (Number(actual.escalon.fijo) || 0);
        variable = actual.comisionGeneradaOficina * pctVariable;
      }

      var bonificacionTotal = fijo + variable;

      return jsonResponse({
        ok: true,
        mes: mesB,
        comision_generada_oficina: actual.comisionGeneradaOficina,
        total_recibido: actual.totalRecibido,
        num_acciones: actual.numAcciones,
        categoria: actual.categoria,
        es_medio: actual.esMedio,
        bonificacion_fija: fijo,
        bonificacion_variable: variable,
        bonificacion_total: bonificacionTotal,
        pct_variable: pctVariable,
        es_continuidad: esContinuidad,
        categoria_mes_anterior: catAnterior
      });
    }

    // --- MIS ACCIONES ---
    if (action === 'mis_acciones') {
      var idAsesorA = params.id_asesor || '';
      var mesA = params.mes ? parseInt(params.mes, 10) : null;
      var todasAcciones = leerHoja(HOJAS.acciones);
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
      var asesoresG = leerHoja(HOJAS.asesores);
      var asesorG = asesoresG.find(function(a) { return a.id_asesor === idAsesorG; });
      if (!asesorG || String(asesorG.rol).toLowerCase() !== 'gerente') {
        return jsonResponse({ ok: false, error: 'Acceso denegado' });
      }
      return jsonResponse({
        ok: true,
        arriendos: leerHoja(HOJAS.arriendos),
        ventas: leerHoja(HOJAS.ventas),
        pagos: leerHoja(HOJAS.pagos),
        comisiones: leerHoja(HOJAS.comisiones),
        inmuebles: leerHoja(HOJAS.inmuebles),
        clientes: leerHoja(HOJAS.clientes),
        asesores: asesoresG.map(function(a) { return { id_asesor: a.id_asesor, nombre: a.nombre }; })
      });
    }

    // --- VERIFICAR DUPLICADO ---
    if (action === 'verificar_duplicado') {
      var tipoDup = params.tipo || '';
      var idInmueble = params.id_inmueble || '';
      var mesDup = params.mes || '';
      var idComprador = params.id_comprador || '';
      var hojaDup = tipoDup === 'arriendos' ? HOJAS.arriendos : HOJAS.ventas;
      var datosDup = leerHoja(hojaDup);
      var duplicado = null;
      for (var i = 0; i < datosDup.length; i++) {
        var d = datosDup[i];
        // Verificar inmueble + mes (original)
        if (String(d.id_inmueble) === String(idInmueble) && String(d.mes) === String(mesDup)) {
          duplicado = d;
          break;
        }
        // Verificar inmueble + comprador (ventas)
        if (tipoDup === 'ventas' && idComprador &&
            String(d.id_inmueble) === String(idInmueble) && String(d.id_comprador) === String(idComprador)) {
          duplicado = d;
          break;
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

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message, stack: err.stack });
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

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

      // Validar integridad referencial: inmueble, arrendador, arrendatario
      var inmRef = leerHoja(HOJAS.inmuebles);
      var cliRef = leerHoja(HOJAS.clientes);
      if (!inmRef.some(function(i){ return String(i.id_inmueble) === String(datos.id_inmueble); })) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Inmueble "' + datos.id_inmueble + '" no existe' });
      }
      if (datos.id_arrendador && !cliRef.some(function(c){ return String(c.id_cliente) === String(datos.id_arrendador); })) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Arrendador "' + datos.id_arrendador + '" no existe' });
      }
      if (datos.id_arrendatario && !cliRef.some(function(c){ return String(c.id_cliente) === String(datos.id_arrendatario); })) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Arrendatario "' + datos.id_arrendatario + '" no existe' });
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
      // Calcular comisión (usar numVal para tolerar coma decimal)
      datos.comision_oficina = numVal(datos.valor_canon) * numVal(datos.pct_comision_oficina);
      // Guardar arriendo
      agregarFila(HOJAS.arriendos, COLUMNAS.arriendos, datos);

      // Guardar comisiones de los asesores
      if (body.comisiones_asesores && body.comisiones_asesores.length > 0) {
        body.comisiones_asesores.forEach(com => {
          agregarFila(HOJAS.comisiones, COLUMNAS.comisiones, {
            id_asesor: com.id_asesor,
            id_negocio: datos.id_arriendo,
            valor_comision: com.valor_comision,
            punta: com.punta,
            participacion: com.participacion || 100,
            estado: 'ACTIVA'
          });
        });
      }

      lock.releaseLock();
      return jsonResponse({ ok: true, id: datos.id_arriendo, mensaje: 'Arriendo registrado' });
    }

    // --- REGISTRAR VENTA ---
    if (action === 'registrar_venta') {
      const datos = body.datos;
      if (!datos['año']) datos['año'] = new Date().getFullYear();

      // Integridad referencial: inmueble, vendedor, comprador
      var inmRefV = leerHoja(HOJAS.inmuebles);
      var cliRefV = leerHoja(HOJAS.clientes);
      if (!inmRefV.some(function(i){ return String(i.id_inmueble) === String(datos.id_inmueble); })) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Inmueble "' + datos.id_inmueble + '" no existe' });
      }
      if (datos.id_vendedor && !cliRefV.some(function(c){ return String(c.id_cliente) === String(datos.id_vendedor); })) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Vendedor "' + datos.id_vendedor + '" no existe' });
      }
      if (datos.id_comprador && !cliRefV.some(function(c){ return String(c.id_cliente) === String(datos.id_comprador); })) {
        lock.releaseLock();
        return jsonResponse({ ok:false, error:'Comprador "' + datos.id_comprador + '" no existe' });
      }

      datos.id_venta = siguienteId(HOJAS.ventas, 'VNT');
      datos.comision_oficina = numVal(datos.valor_base_comision) * numVal(datos.pct_comision_oficina);
      datos.comision_por_punta = datos.comision_oficina / 2;
      datos.estado_venta = 'ACTIVA';
      agregarFila(HOJAS.ventas, COLUMNAS.ventas, datos);

      if (body.comisiones_asesores && body.comisiones_asesores.length > 0) {
        body.comisiones_asesores.forEach(com => {
          agregarFila(HOJAS.comisiones, COLUMNAS.comisiones, {
            id_asesor: com.id_asesor,
            id_negocio: datos.id_venta,
            valor_comision: com.valor_comision,
            punta: com.punta,
            participacion: com.participacion || 100,
            estado: 'ACTIVA'
          });
        });
      }

      // Registrar pagos/cuotas
      // El frontend envía valor_pago (monto del inmueble que se paga).
      // Convertimos a valor_cobrado (comisión proporcional que entra a la oficina).
      if (body.pagos && body.pagos.length > 0) {
        var valorBase = numVal(datos.valor_base_comision);
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

      lock.releaseLock();
      return jsonResponse({ ok: true, id: datos.id_venta, mensaje: 'Venta registrada' });
    }

    // --- REGISTRAR INMUEBLE ---
    if (action === 'registrar_inmueble') {
      const datos = body.datos;

      // Bloquear duplicado por codigo_plataforma (si viene)
      var codigo = String(datos.codigo_plataforma || '').trim();
      if (codigo) {
        var inmExist = leerHoja(HOJAS.inmuebles);
        var dupInm = inmExist.find(function(i){
          return String(i.codigo_plataforma || '').trim().toLowerCase() === codigo.toLowerCase();
        });
        if (dupInm) {
          lock.releaseLock();
          return jsonResponse({
            ok:false,
            error:'Ya existe un inmueble con código de plataforma "' + codigo + '" (' + dupInm.id_inmueble + ' — ' + (dupInm.nombre || '') + '). Use ese en vez de crear uno nuevo.',
            duplicado: dupInm
          });
        }
      }

      datos.id_inmueble = siguienteId(HOJAS.inmuebles, 'INM');
      datos.estado = 'Disponible';
      agregarFila(HOJAS.inmuebles, COLUMNAS.inmuebles, datos);
      lock.releaseLock();
      return jsonResponse({ ok: true, id: datos.id_inmueble, mensaje: 'Inmueble registrado' });
    }

    // --- REGISTRAR CLIENTE ---
    if (action === 'registrar_cliente') {
      const datos = body.datos;
      datos.id_cliente = siguienteId(HOJAS.clientes, 'CLI');
      agregarFila(HOJAS.clientes, COLUMNAS.clientes, datos);
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

      // Copiar template, reemplazar, exportar PDF
      var templateFile = DriveApp.getFileById(TEMPLATE_CUENTA_COBRO_ID);
      var nombreCopia = 'Cuenta de cobro ' + (asesor.nombre||'') + ' - ' + idNegocio;
      var copia = templateFile.makeCopy(nombreCopia);
      var doc = DocumentApp.openById(copia.getId());
      var docBody = doc.getBody();
      Object.keys(reemplazos).forEach(function(k){
        docBody.replaceText('\\{\\{' + k + '\\}\\}', String(reemplazos[k]));
      });
      doc.saveAndClose();

      var pdfBlob = copia.getAs('application/pdf').setName(nombreCopia + '.pdf');

      // Enviar correo
      var asunto = 'Cuenta de cobro — ' + (asesor.nombre||'') + ' — ' + idNegocio;
      var cuerpo = 'Adjunto cuenta de cobro generada automáticamente por el portal REDA3.\n\n' +
                   'Asesor: ' + (asesor.nombre||'') + '\n' +
                   'Negocio: ' + idNegocio + ' (' + tipoNeg + ')\n' +
                   'Concepto: ' + concepto + '\n' +
                   'Valor: ' + reemplazos.valor_numero + '\n';
      var opts = { attachments: [pdfBlob] };
      if (asesor.email) opts.cc = asesor.email;
      MailApp.sendEmail(GERENTE_EMAIL, asunto, cuerpo, opts);

      // Borrar copia temporal del Doc (el PDF ya fue enviado)
      copia.setTrashed(true);

      lock.releaseLock();
      return jsonResponse({
        ok:true,
        mensaje:'Cuenta de cobro enviada al gerente' + (asesor.email ? ' (con copia a ' + asesor.email + ')' : ''),
        valor: valorTotal
      });
    }

    // --- REGISTRAR ACCIÓN COMERCIAL ---
    if (action === 'registrar_accion') {
      const datos = body.datos;
      datos.id_accion = siguienteId(HOJAS.acciones, 'ACC');
      // Calcular mes desde fecha si no viene
      if (!datos.mes && datos.fecha) {
        var fechaP = new Date(datos.fecha);
        datos.mes = fechaP.getMonth() + 1;
      }
      agregarFila(HOJAS.acciones, COLUMNAS.acciones, datos);
      lock.releaseLock();
      return jsonResponse({ ok: true, id: datos.id_accion, mensaje: 'Acción registrada' });
    }

    // --- ACTUALIZAR PAGO ---
    if (action === 'actualizar_pago') {
      var idPago = body.id_pago;
      if (!idPago) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Falta id_pago' }); }
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
        if (valorBase > 0 && valorPago > valorBase) {
          lock.releaseLock();
          return jsonResponse({ ok:false, error:'El valor del pago (' + valorPago + ') excede el valor base de la venta (' + valorBase + ')' });
        }
        datosUpdate.valor_cobrado = valorBase > 0 ? Math.round((valorPago / valorBase) * comOficina) : 0;
      }
      if (body.observacion !== undefined) datosUpdate.observacion = body.observacion;

      actualizarFila(HOJAS.pagos, 'id_pago', idPago, datosUpdate);
      lock.releaseLock();
      return jsonResponse({ ok:true, mensaje:'Pago actualizado' });
    }

    // --- CANCELAR VENTA ---
    if (action === 'cancelar_venta') {
      var idVentaCancel = body.id_venta;
      if (!idVentaCancel) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Falta id_venta' }); }

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

      lock.releaseLock();
      return jsonResponse({ ok:true, mensaje:'Venta cancelada. ' + anuladas + ' comision(es) anuladas.' });
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
