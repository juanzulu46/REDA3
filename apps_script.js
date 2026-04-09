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
  pipeline: 'Pipeline',
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
             'password'],
  inmuebles: ['id_inmueble', 'codigo_plataforma', 'nombre', 'ciudad', 'zona', 'tipo', 'residencial_comercial', 'estado'],
  clientes: ['id_cliente', 'nombre', 'telefono', 'email'],
  arriendos: ['id_arriendo', 'año', 'mes', 'mercado', 'id_inmueble', 'id_arrendador', 'id_arrendatario',
              'valor_canon', 'administracion', 'pct_comision_oficina', 'comision_oficina',
              'oficina_captacion', 'origen_captacion', 'oficina_cierre', 'origen_cierre',
              'referido_captador', 'numero_captador_r', 'valor_ref_captador',
              'referido_cerrador', 'numero_cerrador_r', 'valor_ref_cerrador'],
  ventas: ['id_venta', 'año', 'mes', 'mercado', 'id_inmueble', 'id_vendedor', 'id_comprador',
           'valor_base_comision', 'pct_comision_oficina', 'comision_oficina',
           'comisiones_facturadas', 'comision_por_punta', 'pagos_efectuados',
           'pendiente_por_cobrar', 'fechas_cobro_comision',
           'oficina_captacion', 'origen_captacion', 'oficina_cierre', 'origen_cierre',
           'referido_captador', 'numero_captador_r', 'valor_ref_captador',
           'referido_cerrador', 'numero_cerrador_r', 'valor_ref_cerrador'],
  pipeline: ['id_pipeline', 'año', 'mes', 'mercado', 'id_inmueble', 'id_vendedor', 'id_comprador',
             'valor_base_comision', 'pct_comision_oficina', 'comision_oficina',
             'comisiones_facturadas', 'comision_por_punta', 'pagos_efectuados',
             'pendiente_por_cobrar', 'fechas_cobro_comision',
             'oficina_captacion', 'origen_captacion', 'oficina_cierre', 'origen_cierre',
             'referido_captador', 'valor_ref_captador', 'referido_cerrador', 'valor_ref_cerrador'],
  comisiones: ['id_asesor', 'id_negocio', 'valor_comision', 'punta', 'participacion'],
  oficina: ['id_oficina', 'nombre'],
  origen: ['id_origen', 'nombre', 'circulo'],
  zona: ['id_zona', 'comuna', 'ciudad'],
  acciones: ['id_accion', 'id_asesor', 'fecha', 'mes', 'tipo', 'descripcion'],
  tipos_accion: ['id_tipo', 'nombre', 'activo']
};

// ===== UTILIDADES =====

function getSheet(nombre) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
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

// Calcula la categoría de bonificación de un asesor en un mes específico.
// Usa los datos pre-cargados (datos.arriendos, ventas, comisiones, acciones, bonificaciones)
// para evitar releer las hojas.
// Retorna: { categoria, esMedio, escalon, comisionGeneradaOficina, totalRecibido, numAcciones }
function calcularCategoriaMes(idAsesor, mes, datos) {
  var arriendos = datos.arriendos;
  var ventas = datos.ventas;
  var comisiones = datos.comisiones;
  var acciones = datos.acciones;
  var bonificaciones = datos.bonificaciones;

  // Comisiones del asesor
  var misCom = comisiones.filter(function(c) { return c.id_asesor === idAsesor; });
  var negociosIds = misCom.map(function(c) { return c.id_negocio; });

  // Comisión generada a la oficina por las puntas que participa (50% por punta)
  var comisionGeneradaOficina = 0;
  arriendos.forEach(function(a) {
    if (parseInt(a.mes, 10) !== mes) return;
    if (negociosIds.indexOf(a.id_arriendo) === -1) return;
    var puntas = misCom.filter(function(c) { return c.id_negocio === a.id_arriendo; }).length;
    comisionGeneradaOficina += (Number(a.comision_oficina) || 0) * 0.5 * puntas;
  });
  ventas.forEach(function(v) {
    if (parseInt(v.mes, 10) !== mes) return;
    if (negociosIds.indexOf(v.id_venta) === -1) return;
    var puntas = misCom.filter(function(c) { return c.id_negocio === v.id_venta; }).length;
    comisionGeneradaOficina += (Number(v.comision_oficina) || 0) * 0.5 * puntas;
  });

  // Total recibido del mes (lo que efectivamente cobra)
  var totalRecibido = 0;
  misCom.forEach(function(c) {
    var negocio = arriendos.find(function(a) { return a.id_arriendo === c.id_negocio; }) ||
                  ventas.find(function(v) { return v.id_venta === c.id_negocio; });
    if (negocio && parseInt(negocio.mes, 10) === mes) {
      totalRecibido += Number(c.valor_comision) || 0;
    }
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

  // Evaluación en cascada: el primer escalón cuyas condiciones cumple el asesor
  var escalonAsignado = null;
  var esMedio = false;
  for (var i = 0; i < escalones.length; i++) {
    var e = escalones[i];
    var minCom = Number(e.min_comision_oficina) || 0;
    var minAcc = Number(e.min_acciones) || 0;
    if (comisionGeneradaOficina >= minCom && numAcciones >= minAcc) {
      escalonAsignado = e;
      esMedio = false;
      break;
    }
  }

  // Caso especial PIEDRA media: si no cumplió el umbral monetario de ningún escalón
  // pero sí hizo las acciones mínimas de PIEDRA Y generó algo de comisión, cae en PIEDRA con fijo medio.
  // Si comisión == 0, NO es PIEDRA media → cae en ARENA (solo hizo acciones, no comisionó).
  if (!escalonAsignado && comisionGeneradaOficina > 0) {
    for (var j = 0; j < escalones.length; j++) {
      var ej = escalones[j];
      if (String(ej.categoria).toUpperCase() === 'PIEDRA' &&
          Number(ej.fijo_medio) > 0 &&
          numAcciones >= (Number(ej.min_acciones) || 0)) {
        escalonAsignado = ej;
        esMedio = true;
        break;
      }
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
      var comisiones = leerHoja(HOJAS.comisiones);
      var inmuebles = leerHoja(HOJAS.inmuebles);
      var clientes = leerHoja(HOJAS.clientes);

      var misComisiones = comisiones.filter(function(c) { return c.id_asesor === idAsesor; });
      var misNegocioIds = misComisiones.map(function(c) { return c.id_negocio; });

      var misArriendos = arriendos.filter(function(a) { return misNegocioIds.indexOf(a.id_arriendo) !== -1; });
      var misVentas = ventas.filter(function(v) { return misNegocioIds.indexOf(v.id_venta) !== -1; });

      return jsonResponse({
        ok: true,
        arriendos: misArriendos,
        ventas: misVentas,
        comisiones: misComisiones,
        inmuebles: inmuebles,
        clientes: clientes
      });
    }

    // --- SIGUIENTE ID ---
    if (action === 'siguiente_id') {
      var tipo = params.tipo || '';
      var prefijos = { arriendos: 'ARR', ventas: 'VNT', inmuebles: 'INM', clientes: 'CLI', pipeline: 'PIP' };
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

    // --- VERIFICAR DUPLICADO ---
    if (action === 'verificar_duplicado') {
      var tipoDup = params.tipo || '';
      var idInmueble = params.id_inmueble || '';
      var mesDup = params.mes || '';
      var hojaDup = tipoDup === 'arriendos' ? HOJAS.arriendos : HOJAS.ventas;
      var datosDup = leerHoja(hojaDup);
      var duplicado = null;
      for (var i = 0; i < datosDup.length; i++) {
        if (String(datosDup[i].id_inmueble) === String(idInmueble) && String(datosDup[i].mes) === String(mesDup)) {
          duplicado = datosDup[i];
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
      // Generar ID
      datos.id_arriendo = siguienteId(HOJAS.arriendos, 'ARR');
      // Año automático
      if (!datos['año']) datos['año'] = new Date().getFullYear();
      // Calcular comisión
      datos.comision_oficina = (datos.valor_canon || 0) * (datos.pct_comision_oficina || 0);
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
            participacion: com.participacion || 100
          });
        });
      }

      lock.releaseLock();
      return jsonResponse({ ok: true, id: datos.id_arriendo, mensaje: 'Arriendo registrado' });
    }

    // --- REGISTRAR VENTA ---
    if (action === 'registrar_venta') {
      const datos = body.datos;
      datos.id_venta = siguienteId(HOJAS.ventas, 'VNT');
      if (!datos['año']) datos['año'] = new Date().getFullYear();
      datos.comision_oficina = (datos.valor_base_comision || 0) * (datos.pct_comision_oficina || 0);
      datos.comision_por_punta = datos.comision_oficina / 2;
      agregarFila(HOJAS.ventas, COLUMNAS.ventas, datos);

      if (body.comisiones_asesores && body.comisiones_asesores.length > 0) {
        body.comisiones_asesores.forEach(com => {
          agregarFila(HOJAS.comisiones, COLUMNAS.comisiones, {
            id_asesor: com.id_asesor,
            id_negocio: datos.id_venta,
            valor_comision: com.valor_comision,
            punta: com.punta,
            participacion: com.participacion || 100
          });
        });
      }

      lock.releaseLock();
      return jsonResponse({ ok: true, id: datos.id_venta, mensaje: 'Venta registrada' });
    }

    // --- REGISTRAR INMUEBLE ---
    if (action === 'registrar_inmueble') {
      const datos = body.datos;
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
    // body: { action, id_asesor, id_negocio, tipo: 'arriendo'|'venta' }
    // Llena el template, exporta PDF y envía correo a gerente con copia al asesor
    if (action === 'imprimir_cuenta_cobro') {
      var idAsesor  = body.id_asesor;
      var idNegocio = body.id_negocio;
      var tipoNeg   = body.tipo; // 'arriendo' | 'venta'
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
      } else {
        negocio = leerHoja(HOJAS.ventas).find(function(v){ return v.id_venta === idNegocio; });
        if (!negocio) { lock.releaseLock(); return jsonResponse({ ok:false, error:'Venta no encontrada' }); }
        mesNeg = parseInt(negocio.mes,10); anoNeg = negocio['año'];
        var inmV = inmuebles.find(function(i){ return i.id_inmueble === negocio.id_inmueble; });
        conceptoBase = 'Comisión por venta del inmueble ' + (inmV ? inmV.nombre : negocio.id_inmueble);
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
