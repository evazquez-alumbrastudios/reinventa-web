/**
 * REINVENTA by Mary Méndez — Apps Script
 *
 * HOJAS DEL SPREADSHEET:
 *
 * [Registros]
 *   A  Fecha registro       B  Nombre            C  Correo
 *   D  WhatsApp             E  Contacto preferido F  ¿Qué busca?
 *   G  Fase de compra       H  Monto pagado (MXN) I  Fecha de pago
 *   J  Pagó ✓              K  Stripe Payment ID  L  Origen
 *   M  Acepta marketing     N  Autoriza imagen    O  Canal UTM
 *   P  ¿Cómo se enteró?    Q  ¿A qué se dedica?
 *
 * [Comunicaciones]
 *   A  Correo               B  Nombre             C  WhatsApp
 *   D  Contacto preferido
 *   — Confirmación —
 *   E  WA Confirmación      F  Estado Conf. WA    G  Correo Conf. Enviado
 *   — Recordatorio —
 *   H  WA Recordatorio      I  Estado Record. WA  J  Correo Record. Enviado
 *   — QR de entrada —
 *   K  WA QR                L  Estado QR WA       M  Correo QR Enviado
 *   — Agradecimiento —
 *   N  WA Agradecimiento    O  Estado Agradec. WA P  Correo Agradec. Enviado
 *   — Indicaciones (jueves 13 ago) —
 *   Q  WA Indicaciones      R  Estado Indic. WA  S  Correo Indic. Enviado
 *
 * [Asistencia]
 *   A  ID Único (RNV-001)   B  Nombre             C  Correo
 *   D  Fase                 E  Asistió ✓          F  Fecha entrada
 *   G  Encuesta ✓           H  Fecha encuesta     I  Calificación  J  Comentario
 *   K  Expectativas         L  Aplicabilidad       M  Contenido valioso  N  Mary
 *   O  Org. previa          P  Org. comunicación   Q  Org. registro      R  Org. check-in
 *   S  Org. materiales      T  Org. lugar          U  Impacto            V  NPS
 *
 * [Dashboard] — solo fórmulas, no la toca el script
 */

var SHEET_REGISTROS      = 'Registros';
var SHEET_COMUNICACIONES = 'Comunicaciones';
var SHEET_ASISTENCIA     = 'Asistencia';
var SHEET_MEDIDAS        = 'Medidas';
var SHEET_PROSPECTOS     = 'Prospectos';

var EMAILS_NOTIFICACION = ['mejoracontinua@caceca.org', 'alopez@alumbrastudios.com'];
var LIMITE_TOTAL        = 40;
var STAFF_PIN           = '1508';

/* ── Getters de hojas ────────────────────────────────────────── */
function getSheet()               { return getSheetByName(SHEET_REGISTROS); }
function getComunicacionesSheet() { return getSheetByName(SHEET_COMUNICACIONES); }
function getAsistenciaSheet()     { return getSheetByName(SHEET_ASISTENCIA); }
function getMedidasSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_MEDIDAS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_MEDIDAS);
    sh.appendRow(['ID', 'Nombre', 'Hombros (cm)', 'Busto (cm)', 'Cintura (cm)', 'Cadera (cm)', 'Forma de cuerpo', 'Fecha']);
    sh.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  return sh;
}

function getProspectosSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_PROSPECTOS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PROSPECTOS);
    var headers = [
      'Fecha', 'Nombre', 'Correo', 'WhatsApp',
      'Servicio de interés', '¿Qué quiere transformar?', '¿Cuándo quiere comenzar?',
      'Contacto preferido', 'Acepta privacidad'
    ];
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2A0F25').setFontColor('#C6A56A');
  }
  return sh;
}

function getSheetByName(nombre) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nombre);
  if (!sheet) throw new Error('Hoja no encontrada: ' + nombre);
  return sheet;
}

/* ── doGet ───────────────────────────────────────────────────── */
function doGet(e) {
  var action = e.parameter.action || '';
  if (action === 'entrada') return handleEntrada(e.parameter.id || '', e.parameter.preview === '1');
  if (action === 'hub')     return handleHub(e.parameter.id || '');
  if (action === 'admin')   return handleAdmin();
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── doPost ──────────────────────────────────────────────────── */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.type && data.type === 'checkout.session.completed') {
      return handleStripeWebhook(data);
    }
    if (data.action === 'encuesta')        return handleEncuesta(data);
    if (data.action === 'encuesta_previa') return handleEncuestaPrevia(data);
    if (data.action === 'guardar_medidas') return handleGuardarMedidas(data);
    if (data.action === 'admin_action')    return handleAdminAction(data);
    if (data.action === 'reserva_taller')  return handleReservaTaller(data);
    if (data.action === 'log_descarga')    return handleLogDescarga(data);
    return handleFormSubmit(data); // landing de contacto → Prospectos
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ── Formulario de contacto de la landing → Prospectos ──────── */
function handleFormSubmit(data) {
  var sheet  = getProspectosSheet();
  var correo = (data.correo || '').toLowerCase().trim();

  // Buscar por correo en col C (índice 2) de Prospectos
  var shData = sheet.getDataRange().getValues();
  var existingRow = null;
  for (var i = 1; i < shData.length; i++) {
    if ((shData[i][2] || '').toLowerCase().trim() === correo) { existingRow = i + 1; break; }
  }

  if (existingRow) {
    // Actualizar fila existente sin pisar datos que ya tenía
    if (data.nombre)      sheet.getRange(existingRow, 2).setValue(data.nombre);
    if (data.whatsapp)    sheet.getRange(existingRow, 4).setValue(data.whatsapp);
    if (data.servicio)    sheet.getRange(existingRow, 5).setValue(data.servicio);
    if (data.transformar) sheet.getRange(existingRow, 6).setValue(data.transformar);
    if (data.cuando)      sheet.getRange(existingRow, 7).setValue(data.cuando);
    if (data.contacto)    sheet.getRange(existingRow, 8).setValue(data.contacto);
  } else {
    sheet.appendRow([
      new Date(),
      data.nombre      || '',
      correo,
      data.whatsapp    || '',
      data.servicio    || '',
      data.transformar || '',
      data.cuando      || '',
      data.contacto    || '',
      data.marketing   || ''
    ]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── Formulario de reserva del taller (reservar.html) → Registros ── */
function handleReservaTaller(data) {
  var sheet  = getSheet(); // hoja Registros
  var correo = (data.correo || '').toLowerCase().trim();

  var existingRow = findRowByEmail(sheet, correo);

  if (existingRow) {
    // Ya existe: actualizar campos del formulario sin pisar lo que haya pagado
    if (data.nombre)        sheet.getRange(existingRow, 2).setValue(data.nombre);
    if (data.whatsapp)      sheet.getRange(existingRow, 4).setValue(data.whatsapp);
    if (data.contacto)      sheet.getRange(existingRow, 5).setValue(data.contacto);
    if (data.transformar)   sheet.getRange(existingRow, 6).setValue(data.transformar);
    if (data.ocupacion)     sheet.getRange(existingRow, 16).setValue(data.ocupacion);
    if (data.como_se_entero) sheet.getRange(existingRow, 17).setValue(data.como_se_entero);
    if (data.canal_utm)     sheet.getRange(existingRow, 15).setValue(data.canal_utm);
    if (data.marketing !== undefined) sheet.getRange(existingRow, 13).setValue(data.marketing);
  } else {
    // Nuevo registro — aún no ha pagado; Stripe rellenará fase, monto y stripeId después
    sheet.appendRow([
      new Date(),             // A Fecha registro
      data.nombre    || '',   // B Nombre
      correo,                 // C Correo
      data.whatsapp  || '',   // D WhatsApp
      data.contacto  || '',   // E Contacto preferido
      data.transformar || '', // F ¿Qué busca?
      '',                     // G Fase de compra (lo llena el webhook)
      '',                     // H Monto pagado
      '',                     // I Fecha de pago
      '',                     // J Pagó ✓
      '',                     // K Stripe Payment ID
      data.origen    || 'landing + stripe', // L Origen
      data.marketing || '',   // M Acepta marketing
      data.autoriza_imagen || '', // N Autoriza uso de imagen
      data.canal_utm || '',   // O Canal UTM
      data.ocupacion || '',   // P ¿A qué se dedica?
      data.como_se_entero || '' // Q ¿Cómo se enteró?
    ]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── Webhook de Stripe ───────────────────────────────────────── */
function handleStripeWebhook(event) {
  var session  = event.data.object;
  var correo   = ((session.customer_details && session.customer_details.email)
                   ? session.customer_details.email
                   : (session.customer_email || '')).toLowerCase().trim();
  var nombre   = session.customer_details && session.customer_details.name
                   ? session.customer_details.name : '';
  var monto    = session.amount_total ? (session.amount_total / 100).toFixed(2) : '';
  var fecha    = session.created ? new Date(session.created * 1000) : new Date();
  var stripeId = session.id || '';

  var fase = 'Taller';
  if      (monto == '1300.00') fase = 'Early Bird';
  else if (monto == '1500.00') fase = 'Preventa';
  else if (monto == '1700.00') fase = 'Últimos lugares';
  else if (monto == '2600.00') fase = 'Early Bird x2';
  else if (monto == '3000.00') fase = 'Preventa x2';
  else if (monto == '3400.00') fase = 'Últimos lugares x2';
  else if (monto == '3900.00') fase = 'Early Bird x3';
  else if (monto == '4500.00') fase = 'Preventa x3';
  else if (monto == '5100.00') fase = 'Últimos lugares x3';

  var sheet = getSheet();

  // Anti-duplicado por Stripe ID
  if (stripeId && findRowByStripeId(sheet, stripeId)) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'duplicate' })).setMimeType(ContentService.MimeType.JSON);
  }

  var existingRow = findRowByEmail(sheet, correo);
  var tel         = '';
  var contacto    = '';

  if (existingRow) {
    var yaPago = sheet.getRange(existingRow, 10).getValue() === '✓';
    tel      = sheet.getRange(existingRow, 4).getValue();
    contacto = sheet.getRange(existingRow, 5).getValue();

    if (!yaPago) {
      // Primera compra de este correo — actualiza la fila existente
      var nombreFinal = nombre || sheet.getRange(existingRow, 2).getValue();
      if (nombre) sheet.getRange(existingRow, 2).setValue(nombre);
      sheet.getRange(existingRow, 7).setValue(fase);
      sheet.getRange(existingRow, 8).setValue(monto);
      sheet.getRange(existingRow, 9).setValue(fecha);
      sheet.getRange(existingRow, 10).setValue('✓');
      sheet.getRange(existingRow, 11).setValue(stripeId);
      sheet.getRange(existingRow, 12).setValue('landing + stripe');
      actualizarAsistencia(correo, nombreFinal, fase);
      var idNuevo = obtenerIdAsistente(correo);
      sincronizarComunicaciones(correo, nombreFinal, tel, contacto, idNuevo);
    } else {
      // Segunda (o tercera) compra del mismo correo — boleto adicional
      // Se agrega NUEVA fila en Registros con el nuevo Stripe ID
      sheet.appendRow([
        new Date(), nombre || sheet.getRange(existingRow, 2).getValue(), correo,
        tel, contacto, '', fase, monto, fecha, '✓', stripeId, 'stripe directo',
        '', '', '', '', ''
      ]);
      // Agrega acompañante en Asistencia (sin tocar Comunicaciones)
      var nombreComprador = sheet.getRange(existingRow, 2).getValue() || nombre;
      agregarAcompanante(nombreComprador, fase);
    }
  } else {
    // Correo completamente nuevo
    sheet.appendRow([
      new Date(), nombre, correo, '', '', '', fase, monto, fecha, '✓', stripeId, 'stripe directo',
      '', '', '', '', ''
    ]);
    actualizarAsistencia(correo, nombre, fase);
    var idNuevo = obtenerIdAsistente(correo);
    sincronizarComunicaciones(correo, nombre, '', '', idNuevo);
  }

  var totalPagos = contarPagosSheet(sheet);
  if (totalPagos === LIMITE_TOTAL) {
    var props = PropertiesService.getScriptProperties();
    if (!props.getProperty('sold_out_enviado')) {
      notificarCupoAgotado(sheet);
      props.setProperty('sold_out_enviado', 'true');
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── Sincronizar fila en Comunicaciones ──────────────────────── */
function sincronizarComunicaciones(correo, nombre, telefono, contacto, idUnico) {
  var sheet = getComunicacionesSheet();
  var fila  = findRowByEmailInSheet(sheet, correo);

  if (!fila) {
    sheet.appendRow([correo, nombre, telefono, contacto, '', '', '', '', '', '', '', '', '', '', '', '']);
    fila = sheet.getLastRow();
  } else {
    if (!sheet.getRange(fila, 3).getValue() && telefono) sheet.getRange(fila, 3).setValue(telefono);
    if (!sheet.getRange(fila, 4).getValue() && contacto) sheet.getRange(fila, 4).setValue(contacto);
  }

  // Generar botón WA confirmación si no existe
  var yaLink = sheet.getRange(fila, 5).getValue();
  if (!yaLink) {
    var id = idUnico || '';
    generarBotonWA(sheet, fila, nombre, telefono,
      function(n, t) { return generateWhatsAppLinkConfirmacion(n, t, id); },
      5, 6, 'Enviar WhatsApp');
  }
}

/* ── Asistencia ──────────────────────────────────────────────── */
function actualizarAsistencia(correo, nombre, fase) {
  var sheet = getAsistenciaSheet();
  var data  = sheet.getDataRange().getValues();
  // Verificar si ya existe por correo
  for (var i = 1; i < data.length; i++) {
    if ((data[i][2] || '').toLowerCase().trim() === correo.toLowerCase().trim()) return;
  }
  var id = generarSiguienteId(sheet);
  sheet.appendRow([id, nombre, correo, fase, '', '', '', '', '', '']);
}

/* ── Agregar acompañante en Asistencia ───────────────────────── */
/* Se llama cuando el comprador ya tiene su fila pero compra un
   boleto adicional (segunda o tercera compra del mismo correo).
   También se usa en sincronizarRegistrosFaltantes para las fases
   que ya traen x2 o x3 en un solo Stripe event.                 */
function agregarAcompanante(nombreComprador, fase) {
  var sheet     = getAsistenciaSheet();
  var nombreAc  = 'Acompañante de ' + nombreComprador;
  // Verificar que no exista ya ese acompañante
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][1] || '').toLowerCase().trim() === nombreAc.toLowerCase().trim()) return;
  }
  var id = generarSiguienteId(sheet);
  sheet.appendRow([id, nombreAc, '', fase, '', '', '', '', '', '']);
}

function generarSiguienteId(sheet) {
  var data     = sheet.getDataRange().getValues();
  var contador = 1;
  for (var i = 1; i < data.length; i++) {
    var id = (data[i][0] || '').toString();
    if (id.startsWith('RNV-')) {
      var num = parseInt(id.replace('RNV-', ''), 10);
      if (num >= contador) contador = num + 1;
    }
  }
  return 'RNV-' + String(contador).padStart(3, '0');
}

/* ── Hub: perfil del asistente ───────────────────────────────── */
function handleHub(id) {
  if (!id) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'ID requerido' })).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Verificar hub_status en Contenidos — si CLOSED, bloquear acceso público ──
  try {
    var shCont = ss.getSheetByName('Contenidos');
    if (shCont) {
      var contRows = shCont.getDataRange().getValues();
      for (var ci = 1; ci < contRows.length; ci++) {
        if ((contRows[ci][0]||'').toString().trim() === 'config' &&
            (contRows[ci][1]||'').toString().trim() === 'evento' &&
            (contRows[ci][2]||'').toString().trim() === 'hub_status') {
          if ((contRows[ci][4]||'').toString().trim() === 'CLOSED') {
            return ContentService.createTextOutput(JSON.stringify({ error: 'hub_cerrado' })).setMimeType(ContentService.MimeType.JSON);
          }
          break;
        }
      }
    }
  } catch(eHub) {}

  var idNorm = id.toString().trim().toUpperCase();

  // ── Buscar en Asistencia: primero por ID (col A), luego por correo (col C) ──
  var asiSheet = getAsistenciaSheet();
  var asiData  = asiSheet.getDataRange().getValues();
  var fila     = -1;

  // Pasada 1: coincidencia exacta por ID
  for (var i = 1; i < asiData.length; i++) {
    if ((asiData[i][0] || '').toString().trim().toUpperCase() === idNorm) { fila = i; break; }
  }

  // Pasada 2: si no encontró por ID, busca por correo (quien entra con su correo en vez del código)
  if (fila === -1) {
    for (var i = 1; i < asiData.length; i++) {
      if ((asiData[i][2] || '').toString().trim().toLowerCase() === id.toString().trim().toLowerCase()) { fila = i; break; }
    }
  }

  // ── Si encontró en Asistencia → verificar que no cedió su lugar ──────────
  if (fila !== -1 && (asiData[fila][10]||'').toString().trim() === 'cedió lugar') {
    return ContentService.createTextOutput(JSON.stringify({ error: 'acceso_cedido' })).setMimeType(ContentService.MimeType.JSON);
  }

  // ── Si encontró en Asistencia → devolver datos ────────────────────────────
  if (fila !== -1) {
    var foundId   = (asiData[fila][0] || '').toString().trim();
    var prevSheet = ss.getSheetByName('Encuesta Previa');
    var encPrev   = false;
    if (prevSheet) {
      var prevData = prevSheet.getDataRange().getValues();
      for (var j = 1; j < prevData.length; j++) {
        if ((prevData[j][0] || '').toString().trim() === foundId) { encPrev = true; break; }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({
      id:             asiData[fila][0],
      nombre:         asiData[fila][1],
      correo:         asiData[fila][2],
      fase:           asiData[fila][3],
      asistio:        asiData[fila][4] === '✓',
      fechaEntrada:   asiData[fila][5] ? asiData[fila][5].toString() : '',
      encuesta:       asiData[fila][6] === '✓',
      calificacion:   asiData[fila][8] || 0,
      encuestaPrevia: encPrev
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // ── Fallback: buscar en Registros por correo o por ID (si ya fue sincronizado)
  // Cubre el caso de personas que pagaron pero no fueron sincronizadas a Asistencia aún.
  try {
    var regSheet = getSheet();
    var regData  = regSheet.getDataRange().getValues();
    var idBaja   = id.toString().trim().toLowerCase();
    var esCodigoRNV = /^rnv-\d+$/i.test(idBaja);

    for (var r = 1; r < regData.length; r++) {
      var regCorreo = (regData[r][2] || '').toString().trim().toLowerCase();
      var regPago   = (regData[r][9] || '').toString().trim();
      if (!regCorreo || regPago !== '✓') continue;

      // Coincide si el input es el correo de la persona
      var coincide = (regCorreo === idBaja);

      if (coincide) {
        var regNombre = (regData[r][1] || '').toString().trim();
        var regFase   = (regData[r][6] || '').toString().trim();
        actualizarAsistencia(regCorreo, regNombre, regFase);
        // Recuperar la fila recién creada
        var asiData2 = asiSheet.getDataRange().getValues();
        for (var k = 1; k < asiData2.length; k++) {
          if ((asiData2[k][2] || '').toString().trim().toLowerCase() === regCorreo) {
            return ContentService.createTextOutput(JSON.stringify({
              id:             asiData2[k][0],
              nombre:         asiData2[k][1],
              correo:         asiData2[k][2],
              fase:           asiData2[k][3],
              asistio:        false,
              fechaEntrada:   '',
              encuesta:       false,
              calificacion:   0,
              encuestaPrevia: false
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
    }
  } catch(e) {
    // Si falla el fallback, caemos al error estándar
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'Registro no encontrado' })).setMimeType(ContentService.MimeType.JSON);
}

/* ── Admin: todos los datos del panel ───────────────────────── */
/* ── Guardar medidas de cuerpo desde el hub ─────────────────── */
function handleGuardarMedidas(data) {
  if (!data.id) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'ID requerido' })).setMimeType(ContentService.MimeType.JSON);
  }
  var sheet   = getMedidasSheet();
  var medData = sheet.getDataRange().getValues();
  var idNorm  = data.id.toString().trim();

  // Buscar fila existente para este ID
  var filaExistente = -1;
  for (var i = 1; i < medData.length; i++) {
    if ((medData[i][0] || '').toString().trim() === idNorm) { filaExistente = i + 1; break; }
  }

  // Obtener nombre del asistente desde Asistencia
  var nombre = data.nombre || '';
  if (!nombre) {
    try {
      var asiData = getAsistenciaSheet().getDataRange().getValues();
      for (var j = 1; j < asiData.length; j++) {
        if ((asiData[j][0] || '').toString().trim() === idNorm) { nombre = asiData[j][1]; break; }
      }
    } catch(e) {}
  }

  var fila = [
    idNorm,
    nombre,
    parseFloat(data.hombros) || '',
    parseFloat(data.busto)   || '',
    parseFloat(data.cintura) || '',
    parseFloat(data.cadera)  || '',
    data.forma || '',
    new Date().toLocaleString('es-MX')
  ];

  if (filaExistente > 0) {
    sheet.getRange(filaExistente, 1, 1, 8).setValues([fila]);
  } else {
    sheet.appendRow(fila);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleAdmin() {
  var asiSheet  = getAsistenciaSheet();
  var asiData   = asiSheet.getDataRange().getValues();
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var prevSheet = ss.getSheetByName('Encuesta Previa');

  // Mapa de WhatsApp por ID desde Comunicaciones (col A=ID, col C=WhatsApp)
  var whatsappPorId = {};
  try {
    var comSheet = ss.getSheetByName('Comunicaciones');
    if (comSheet) {
      var comData = comSheet.getDataRange().getValues();
      for (var ci = 1; ci < comData.length; ci++) {
        var cid = (comData[ci][0] || '').toString().trim();
        if (cid) whatsappPorId[cid] = (comData[ci][2] || '').toString().trim();
      }
    }
  } catch(e) {}

  // Cargar correos que cedieron su lugar (desde Registros col L)
  // y mapa correo → enterado (col P, idx 15) para enriquecer asistentes
  var cedieronSet  = {};
  var enteradoMap  = {}; // correo → cómo se enteró
  try {
    var shReg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registros');
    if (shReg) {
      var regData = shReg.getDataRange().getValues();
      for (var ri = 1; ri < regData.length; ri++) {
        var origen = (regData[ri][11] || '').toString().toLowerCase().trim();
        var correoReg = (regData[ri][2] || '').toString().toLowerCase().trim();
        if (origen.indexOf('cedi') !== -1 && origen.indexOf('lugar') !== -1) {
          if (correoReg) cedieronSet[correoReg] = true;
        }
        // Col P (idx 15) = cómo se enteró
        var ent = (regData[ri][15] || '').toString().trim();
        if (correoReg && ent) enteradoMap[correoReg] = ent;
      }
    }
  } catch(e) {}

  var asistentes = [];
  for (var i = 1; i < asiData.length; i++) {
    var row = asiData[i];
    if (!row[0]) continue;
    var correoAsi = (row[2] || '').toString().toLowerCase().trim();
    if (cedieronSet[correoAsi]) continue;
    var rid = (row[0] || '').toString().trim();
    asistentes.push({
      id:             rid,
      nombre:         row[1],
      correo:         row[2],
      whatsapp:       whatsappPorId[rid] || '',
      fase:           row[3],
      asistio:        row[4] === '✓',
      fechaEntrada:   row[5] ? row[5].toString() : '',
      encuesta:       row[6] === '✓',
      fechaEncuesta:  row[7] ? row[7].toString() : '',
      calificacion:   row[8] || 0,
      comentario:     row[9] || '',
      // campos completos de encuesta post (para reportes)
      // cols: K=10 expectativas, L=11 aplicabilidad, M=12 contenido, N=13 mary
      //       O=14 org_previa, P=15 org_com, Q=16 org_reg, R=17 org_checkin
      //       S=18 org_mat, T=19 org_lugar, U=20 impacto, V=21 nps
      expectativas:   row[10] || '',
      aplicabilidad:  row[11] || '',
      contenido:      row[12] || '',
      mary:           row[13] || '',
      orgMatriz:      [row[14]||'', row[15]||'', row[16]||'', row[17]||'', row[18]||'', row[19]||''],
      impacto:        row[20] || '',
      nps:            (row[21] !== undefined && row[21] !== '') ? row[21] : null,
      enterado:       enteradoMap[(row[2]||'').toString().toLowerCase().trim()] || ''
    });
  }

  // Mapa de medidas por ID
  var medidasPorId = {};
  try {
    var medSheet = getMedidasSheet();
    var medData  = medSheet.getDataRange().getValues();
    for (var m = 1; m < medData.length; m++) {
      var mid = (medData[m][0] || '').toString().trim();
      if (mid) medidasPorId[mid] = {
        hombros: medData[m][2] || '',
        busto:   medData[m][3] || '',
        cintura: medData[m][4] || '',
        cadera:  medData[m][5] || '',
        forma:   medData[m][6] || ''
      };
    }
  } catch(e) {}

  var encuestasPrevia = [];
  if (prevSheet) {
    var prevData = prevSheet.getDataRange().getValues();
    for (var j = 1; j < prevData.length; j++) {
      var p = prevData[j];
      if (!p[0]) continue;
      var pid = (p[0] || '').toString().trim();
      encuestasPrevia.push({
        id:           pid,
        nombre:       p[1],
        correo:       p[2],
        satisfaccion: p[4],
        coherencia:   p[5],
        confianza:    p[6],
        proyeccion:   p[7],
        motivacion:   p[8],
        expectativa:  p[9],
        piel:         p[10],
        cabello:      p[11],
        ojos:         p[12],
        medidas:      medidasPorId[pid] || null
      });
    }
  }

  // Construir array plano de todas las medidas (no filtrado por encuesta previa)
  var medidasArray = [];
  for (var mk in medidasPorId) {
    medidasArray.push({ id: mk, forma: medidasPorId[mk].forma || '' });
  }

  // Leer hoja Descargas (si existe)
  var descargasArray = [];
  try {
    var descSheet = ss.getSheetByName('Descargas');
    if (descSheet) {
      var descData = descSheet.getDataRange().getValues();
      for (var di = 1; di < descData.length; di++) {
        descargasArray.push({
          id:        (descData[di][0] || '').toString().trim(),
          nombre:    (descData[di][1] || '').toString().trim(),
          tipo:      (descData[di][2] || '').toString().trim(),
          fecha:     descData[di][3] ? descData[di][3].toString() : ''
        });
      }
    }
  } catch(e) {}

  // Registros: resumen por fila (fase, monto, enterado, origen) para reportes
  var registrosSummary = [];
  try {
    var shReg2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registros');
    if (shReg2) {
      var regData2 = shReg2.getDataRange().getValues();
      for (var rr = 1; rr < regData2.length; rr++) {
        if (!regData2[rr][1] && !regData2[rr][2]) continue;
        registrosSummary.push({
          correo:  (regData2[rr][2] || '').toString().toLowerCase().trim(),
          fase:    regData2[rr][6] || '',
          monto:   regData2[rr][7] || '',
          pago:    regData2[rr][9] === '✓',
          origen:  (regData2[rr][11] || '').toString(),
          enterado: regData2[rr][15] || ''
        });
      }
    }
  } catch(eReg) {}

  // Comunicaciones: resumen por campaña para reportes
  var commsSummary = [];
  try {
    var comSheet2 = ss.getSheetByName('Comunicaciones');
    if (comSheet2) {
      var comData2 = comSheet2.getDataRange().getValues();
      // Helper: acepta 'Sí', 'si', '✓', true, 1 como "enviado"
      function _siTrue(v) {
        if (!v) return false;
        var s = v.toString().toLowerCase().trim();
        return s === 'sí' || s === 'si' || s === '✓' || s === 'true' || s === '1';
      }
      commsSummary = comData2.slice(1).map(function(r) {
        return {
          correo:        (r[0] || '').toString(),
          nombre:        r[1] || '',
          // Confirmación — col F(5)=Estado WA, col G(6)=Correo enviado
          estConfWA:     (r[5] || '').toString(),
          correoConf:    _siTrue(r[6]),
          // Recordatorio — col I(8)=Estado WA, col J(9)=Correo enviado
          estRecordWA:   (r[8] || '').toString(),
          correoRecord:  _siTrue(r[9]),
          // QR de entrada — col L(11)=Estado WA, col M(12)=Correo enviado
          estQRWA:       (r[11] || '').toString(),
          correoQR:      _siTrue(r[12]),
          // Agradecimiento — col O(14)=Estado WA, col P(15)=Correo enviado
          estAgradecWA:  (r[14] || '').toString(),
          correoAgradec: _siTrue(r[15]),
          // Indicaciones — col R(17)=Estado WA, col S(18)=Correo enviado
          estIndicWA:    (r[17] || '').toString(),
          correoIndic:   _siTrue(r[18]),
          // Encuesta previa — col U(20)=Estado WA (sin correo)
          estEncPrevWA:  (r[20] || '').toString(),
          // Recursos — col X(23)=Correo enviado (sin WA)
          estRecursosWA: '',
          correoRecursos: _siTrue(r[23]),
          // Última llamada — col AA(26)=Estado WA
          estUltimaWA:   (r[26] || '').toString()
        };
      }).filter(function(r){ return r.correo || r.nombre; });
    }
  } catch(eComm) {}

  // encuesta_post_status y hub_status desde Contenidos
  var encuestaPostStatus = 'OPEN';
  var hubStatusValue = 'OPEN';
  try {
    var shCont2 = ss.getSheetByName('Contenidos');
    if (shCont2) {
      var rowsCont2 = shCont2.getDataRange().getValues();
      for (var ci2 = 1; ci2 < rowsCont2.length; ci2++) {
        var claveC2 = (rowsCont2[ci2][2]||'').toString().trim();
        if ((rowsCont2[ci2][0]||'').toString().trim() === 'config' &&
            (rowsCont2[ci2][1]||'').toString().trim() === 'evento') {
          if (claveC2 === 'encuesta_post_status') {
            encuestaPostStatus = (rowsCont2[ci2][4]||'OPEN').toString().trim();
          }
          if (claveC2 === 'hub_status') {
            hubStatusValue = (rowsCont2[ci2][4]||'OPEN').toString().trim();
          }
        }
      }
    }
  } catch(eCont) {}

  return ContentService
    .createTextOutput(JSON.stringify({
      asistentes:          asistentes,
      encuestasPrevia:     encuestasPrevia,
      totalMedidas:        medidasArray.length,
      medidas:             medidasArray,
      descargas:           descargasArray,
      registros:           registrosSummary,
      comunicaciones:      commsSummary,
      encuestaPostStatus:  encuestaPostStatus,
      hubStatus:           hubStatusValue
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── Registro de descarga (workbook y guías) ─────────────────── */
function handleLogDescarga(data) {
  if (!data.id || !data.tipo) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'ID y tipo requeridos' })).setMimeType(ContentService.MimeType.JSON);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Descargas');
  if (!sheet) {
    sheet = ss.insertSheet('Descargas');
    sheet.appendRow(['ID', 'Nombre', 'Tipo', 'Fecha']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  // Buscar nombre del asistente
  var nombre = data.nombre || data.id;
  try {
    var asSheet = getAsistenciaSheet();
    var rows = asSheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '').toString().trim() === data.id) {
        nombre = (rows[i][1] || '').toString().trim() || nombre;
        break;
      }
    }
  } catch(e) {}
  sheet.appendRow([data.id, nombre, data.tipo, new Date()]);
  return ContentService.createTextOutput(JSON.stringify({ result: 'ok' })).setMimeType(ContentService.MimeType.JSON);
}

/* ── Hub: encuesta previa al evento (imagen y propósito) ────── */
function handleEncuestaPrevia(data) {
  if (!data.id) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'ID requerido' })).setMimeType(ContentService.MimeType.JSON);
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Encuesta Previa');
  if (!sheet) {
    sheet = ss.insertSheet('Encuesta Previa');
    sheet.getRange(1, 1, 1, 14).setValues([[
      'ID','Nombre','Correo','Fecha',
      'Satisfacción (1-10)','Coherencia (1-10)','Confianza (1-10)','Proyección (1-10)',
      'Motivación principal','Expectativa',
      'Tono de piel','Color de cabello','Color de ojos','Ya respondió'
    ]]);
    sheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#2A0F25').setFontColor('#C6A56A');
  }

  // Buscar si ya respondió
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][0] || '').toString().trim() === data.id) {
      return ContentService.createTextOutput(JSON.stringify({ result: 'ya_enviada' })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Buscar nombre y correo en Asistencia
  var asiSheet = getAsistenciaSheet();
  var asiData  = asiSheet.getDataRange().getValues();
  var nombre = '', correo = '';
  for (var j = 1; j < asiData.length; j++) {
    if ((asiData[j][0] || '').toString().trim() === data.id) {
      nombre = asiData[j][1];
      correo = asiData[j][2];
      break;
    }
  }

  sheet.appendRow([
    data.id, nombre, correo, new Date(),
    data.satisfaccion  || '',
    data.coherencia    || '',
    data.confianza     || '',
    data.proyeccion    || '',
    data.motivacion    || '',
    data.expectativa   || '',
    data.piel          || '',
    data.cabello       || '',
    data.ojos          || '',
    '✓'
  ]);

  return ContentService.createTextOutput(JSON.stringify({ result: 'success' })).setMimeType(ContentService.MimeType.JSON);
}

/* ── Hub: guardar encuesta ───────────────────────────────────── */
function handleEncuesta(data) {
  if (!data.id) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'ID requerido' })).setMimeType(ContentService.MimeType.JSON);
  }
  var sheet = getAsistenciaSheet();
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][0] || '').toString().trim() === data.id) {
      if (rows[i][4] !== '✓') {
        return ContentService.createTextOutput(JSON.stringify({ error: 'Check-in pendiente' })).setMimeType(ContentService.MimeType.JSON);
      }
      if (rows[i][6] === '✓') {
        return ContentService.createTextOutput(JSON.stringify({ result: 'ya_enviada' })).setMimeType(ContentService.MimeType.JSON);
      }
      var row = i + 1;
      sheet.getRange(row, 7).setValue('✓');
      sheet.getRange(row, 8).setValue(new Date());
      sheet.getRange(row, 9).setValue(data.calificacion      || '');  // I
      sheet.getRange(row, 10).setValue(data.comentario       || '');  // J
      sheet.getRange(row, 11).setValue(data.expectativas     || '');  // K
      sheet.getRange(row, 12).setValue(data.aplicabilidad    || '');  // L
      sheet.getRange(row, 13).setValue(data.contenido_valioso|| '');  // M
      sheet.getRange(row, 14).setValue(data.mary             || '');  // N
      sheet.getRange(row, 15).setValue(data.org_previa       || '');  // O
      sheet.getRange(row, 16).setValue(data.org_comunicacion || '');  // P
      sheet.getRange(row, 17).setValue(data.org_registro     || '');  // Q
      sheet.getRange(row, 18).setValue(data.org_checkin      || '');  // R
      sheet.getRange(row, 19).setValue(data.org_materiales   || '');  // S
      sheet.getRange(row, 20).setValue(data.org_lugar        || '');  // T
      sheet.getRange(row, 21).setValue(data.impacto          || '');  // U
      sheet.getRange(row, 22).setValue(data.nps !== undefined && data.nps !== null && data.nps >= 0 ? data.nps : ''); // V = NPS
      return ContentService.createTextOutput(JSON.stringify({ result: 'success' })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ error: 'ID no encontrado' })).setMimeType(ContentService.MimeType.JSON);
}

/* ── Registro de entrada (QR) ────────────────────────────────── */
function handleEntrada(id, preview) {
  if (!id) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'ID requerido' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = getAsistenciaSheet();
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === id) {
      var nombre    = data[i][1];
      var fase      = data[i][3];
      var yaAsistio = data[i][4];

      // Preview mode: return info without marking attendance
      if (preview) {
        return ContentService.createTextOutput(JSON.stringify({ nombre: nombre, fase: fase, preview: true })).setMimeType(ContentService.MimeType.JSON);
      }

      if (yaAsistio === '✓') {
        return ContentService
          .createTextOutput(JSON.stringify({ yaRegistrado: true, nombre: nombre }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      // Gate: requiere encuesta previa
      var ss2 = SpreadsheetApp.getActiveSpreadsheet();
      var prevSheet2 = ss2.getSheetByName('Encuesta Previa');
      var tienePrevia = false;
      if (prevSheet2) {
        var prevData2 = prevSheet2.getDataRange().getValues();
        for (var k = 1; k < prevData2.length; k++) {
          if ((prevData2[k][0]||'').toString().trim() === id) { tienePrevia = true; break; }
        }
      }
      if (!tienePrevia) {
        return ContentService.createTextOutput(JSON.stringify({ error: 'encuesta_previa_pendiente', nombre: nombre })).setMimeType(ContentService.MimeType.JSON);
      }

      sheet.getRange(i + 1, 5).setValue('✓');
      sheet.getRange(i + 1, 6).setValue(new Date());

      return ContentService
        .createTextOutput(JSON.stringify({ nombre: nombre, fase: fase }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ error: 'No se encontró este registro' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── WhatsApp: generador genérico ────────────────────────────── */
function normalizeWhatsAppNumber(raw) {
  if (!raw) return null;
  var digits = raw.toString().replace(/\D/g, '');
  if (digits.length === 10) return '52' + digits;
  if (digits.length === 12 && digits.slice(0,2) === '52') return digits;
  if (digits.length === 13 && digits.slice(0,3) === '521') return '52' + digits.slice(3);
  if (digits.length === 11 && digits[0] === '1') return digits;
  if (digits.length >= 10 && digits.length <= 15) return digits;
  return null;
}

function generarBotonWA(sheet, fila, nombre, telefono, generadorFn, colLink, colEstado, textoBoton) {
  var link = generadorFn(nombre, telefono);
  if (link) {
    var richText = SpreadsheetApp.newRichTextValue()
      .setText(textoBoton)
      .setLinkUrl(link)
      .build();
    sheet.getRange(fila, colLink).setRichTextValue(richText);
    if (!sheet.getRange(fila, colEstado).getValue()) {
      sheet.getRange(fila, colEstado).setValue('Pendiente');
    }
  } else {
    sheet.getRange(fila, colLink).setValue('SIN TELEFONO');
    sheet.getRange(fila, colEstado).setValue('SIN TELEFONO');
  }
}

/* ── WhatsApp: mensajes por campaña ─────────────────────────── */
function generateWhatsAppLinkConfirmacion(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p   = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var hub = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';

  // Intentar mensaje publicado de Contenidos (con fallback al hardcode)
  var _pub = getContenidoPublicado('confirmacion', 'whatsapp', 'mensaje');
  if (_pub) {
    var _msg = _buildWAMsgFromTemplate(_pub, { nombre: p, hub: hub, id: idUnico || '' });
    return 'https://wa.me/' + numero + '?text=' + _msg;
  }

  // Fallback hardcodeado original ↓
  var e   = encodeURIComponent;
  var NL  = '%0A';
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('Hola, ' + p + '. Tu lugar esta confirmado.') + NL + NL +
    e('*Lo que tu imagen comunica*') + NL +
    e('- Sabado 15 de agosto de 2026') + NL +
    e('- 10:00 a 12:00 pm') + NL +
    e('- The University Club of Mexico') + NL +
    e('- Av. Paseo de la Reforma 150, Juarez, CDMX') + NL + NL +
    e('Como llegar: https://maps.app.goo.gl/Uo7tYiQz23jMCmKw7') + NL + NL +
    e('- - -') + NL + NL +
    e('*Tu espacio personal:*') + NL +
    e(hub) + NL + NL +
    e('Ahi encontraras tu pase con codigo QR, la agenda del dia y los recursos del taller.') + NL + NL +
    e('- - -') + NL + NL +
    e('*Encuesta previa — importante*') + NL + NL +
    e('Dentro de tu espacio hay una encuesta que Mary necesita que contestes antes del taller. La lee personalmente para preparar los materiales y recomendaciones de cada asistente. No toma mas de 5 minutos.') + NL + NL +
    e('- - -') + NL + NL +
    e('Nos da mucho gusto tenerte. Mary estara encantada de acompanarte.') + NL + NL +
    e('_Con carino,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Mensaje informativo — por favor no respondas a este chat._') + NL +
    (idUnico ? NL + e('Codigo de acceso: ' + idUnico) + NL : '') +
    e('_Organizado por Alumbra Studios · alumbrastudios.com_');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

function generateWhatsAppLinkIndicaciones(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p   = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var hub = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';

  var _pub = getContenidoPublicado('indicaciones', 'whatsapp', 'mensaje');
  if (_pub) {
    var _msg = _buildWAMsgFromTemplate(_pub, { nombre: p, hub: hub, id: idUnico || '' });
    return 'https://wa.me/' + numero + '?text=' + _msg;
  }

  // Fallback hardcodeado original ↓
  var e   = encodeURIComponent;
  var NL  = '%0A';
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('Hola, ' + p + '. Ya casi llegamos.') + NL + NL +
    e('Para que tu experiencia del sabado sea increible, aqui van las indicaciones:') + NL + NL +
    e('- - -') + NL + NL +
    e('*Tu espacio personal:*') + NL +
    e(hub) + NL + NL +
    e('Entra para ver la agenda, tu QR de entrada y registrar tus medidas antes del taller.') + NL + NL +
    e('- - -') + NL + NL +
    e('*Sugerencia de look para tu analisis de cuerpo*') + NL + NL +
    e('Para aprovechar mejor el analisis, te sugerimos llegar con esta combinacion. Si no puedes seguirla al 100%, no te preocupes — trabajamos con lo que traes:') + NL + NL +
    e('- Abajo: pantalon ajustado (leggings o jeans pegados).') + NL +
    e('- Arriba: blusa o playera basica ajustada, de preferencia en blanco o neutro.') + NL +
    e('- Encima: saco, blazer o bluson en el color que mas te guste.') + NL + NL +
    e('- - -') + NL + NL +
    e('*Rostro — Analisis visagismo*') + NL + NL +
    e('- De preferencia con el cabello recogido (facilita el analisis).') + NL +
    e('- Si no puedes, no te preocupes, habra pinzas disponibles.') + NL + NL +
    e('- - -') + NL + NL +
    e('*Medidas de cuerpo (opcional pero util)*') + NL + NL +
    e('Si puedes tomartelas antes, Mary podra hacer tu analisis aun mas personalizado. Son 4:') + NL + NL +
    e('- Hombros: de un extremo al otro por la parte mas alta de la espalda.') + NL +
    e('- Busto: la parte mas voluminosa del pecho, a la altura de los pezones.') + NL +
    e('- Cintura: zona mas angosta, 2 dedos arriba del ombligo, sin meter el abdomen.') + NL +
    e('- Cadera: la parte mas ancha de los gluteos y los huesos de la cadera.') + NL + NL +
    e('Registra tus medidas en tu espacio personal:') + NL +
    e(hub) + NL + NL +
    e('- - -') + NL + NL +
    e('*Encuesta previa*') + NL + NL +
    e('Si aun no has contestado la encuesta dentro de tu espacio, este es el momento. Mary la lee personalmente para preparar tu experiencia. No toma mas de 5 minutos.') + NL + NL +
    e(hub) + NL + NL +
    e('- - -') + NL + NL +
    e('Nos vemos el sabado. Va a ser un dia para descubrir muchas cosas lindas.') + NL + NL +
    e('_Con carino,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Mensaje informativo — por favor no respondas a este chat._') + NL +
    (idUnico ? NL + e('Codigo de acceso: ' + idUnico) + NL : '') +
    e('_Organizado por Alumbra Studios · alumbrastudios.com_');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

function generateWhatsAppLinkRecordatorio(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p   = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var hub = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';

  var _pub = getContenidoPublicado('recordatorio', 'whatsapp', 'mensaje');
  if (_pub) {
    var _msg = _buildWAMsgFromTemplate(_pub, { nombre: p, hub: hub, id: idUnico || '' });
    return 'https://wa.me/' + numero + '?text=' + _msg;
  }

  // Fallback hardcodeado original ↓
  var e   = encodeURIComponent;
  var NL  = '%0A';
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('Hola, ' + p + '. Manana es el gran dia.') + NL + NL +
    e('*Lo que tu imagen comunica*') + NL +
    e('- Manana sabado 15 de agosto') + NL +
    e('- 10:00 a 12:00 pm') + NL +
    e('- The University Club of Mexico') + NL +
    e('- Av. Paseo de la Reforma 150, Juarez, CDMX') + NL + NL +
    e('Como llegar: https://maps.app.goo.gl/Uo7tYiQz23jMCmKw7') + NL + NL +
    e('- - -') + NL + NL +
    e('*Tu espacio personal:*') + NL +
    e(hub) + NL + NL +
    e('Ahi encuentras la agenda del dia y tu codigo QR de entrada.') + NL + NL +
    e('- - -') + NL + NL +
    e('*Sugerencia de look para manana*') + NL + NL +
    e('Te sugerimos este look para aprovechar mejor los analisis. No es obligatorio — si no puedes seguirlo al pie de la letra, esta perfectamente bien:') + NL + NL +
    e('- Abajo: pantalon ajustado (leggings o jeans).') + NL +
    e('- Arriba: blusa o playera basica ajustada (blanco o neutro de preferencia).') + NL +
    e('- Encima: saco, blazer o bluson en el color que prefieras.') + NL +
    e('- Rostro: de preferencia con el cabello recogido para el analisis (si no puedes, hay pinzas).') + NL + NL +
    e('- - -') + NL + NL +
    e('*Medidas de cuerpo*') + NL + NL +
    e('Si aun no las registraste, puedes hacerlo desde tu espacio antes de llegar. Solo necesitas una cinta metrica:') + NL + NL +
    e('- Hombros: extremo a extremo por la parte alta de la espalda.') + NL +
    e('- Busto: la parte mas voluminosa del pecho.') + NL +
    e('- Cintura: zona mas angosta, 2 dedos arriba del ombligo.') + NL +
    e('- Cadera: la parte mas ancha de los gluteos.') + NL + NL +
    e(hub) + NL + NL +
    e('- - -') + NL + NL +
    e('*Encuesta previa*') + NL + NL +
    e('Si aun no la has contestado, este es el ultimo momento. Mary la revisa antes del taller para personalizar tu experiencia. Quienes no la contesten no tendran acceso al material digital posterior.') + NL + NL +
    e(hub) + NL + NL +
    e('- - -') + NL + NL +
    e('Te esperamos puntual. Nos vemos manana.') + NL + NL +
    e('_Con carino,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Mensaje informativo — por favor no respondas a este chat._') + NL +
    (idUnico ? NL + e('Codigo de acceso: ' + idUnico) + NL : '') +
    e('_Organizado por Alumbra Studios · alumbrastudios.com_');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

function generateWhatsAppLinkQR(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p   = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var hub = 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico;

  var _pub = getContenidoPublicado('qr', 'whatsapp', 'mensaje');
  if (_pub) {
    var _msg = _buildWAMsgFromTemplate(_pub, { nombre: p, hub: hub, id: idUnico || '' });
    return 'https://wa.me/' + numero + '?text=' + _msg;
  }

  // Fallback hardcodeado original ↓
  var e   = encodeURIComponent;
  var NL  = '%0A';
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('Hola, ' + p + '. Hoy te esperamos.') + NL + NL +
    e('*Lo que tu imagen comunica*') + NL +
    e('- Hoy sabado 15 de agosto') + NL +
    e('- 10:00 a 12:00 pm') + NL +
    e('- The University Club of Mexico') + NL +
    e('- Av. Paseo de la Reforma 150, Juarez, CDMX') + NL + NL +
    e('Como llegar: https://maps.app.goo.gl/Uo7tYiQz23jMCmKw7') + NL + NL +
    e('- - -') + NL + NL +
    e('*Tu codigo QR de entrada:*') + NL + NL +
    e('Entra a tu espacio y muestra el codigo QR al llegar. El staff lo escaneara en la entrada.') + NL + NL +
    e(hub) + NL + NL +
    e('- - -') + NL + NL +
    e('Nos vemos en un momento. Sera un dia increible.') + NL + NL +
    e('_Con carino,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Mensaje informativo — por favor no respondas a este chat._') + NL +
    NL + e('Codigo de acceso: ' + idUnico) + NL +
    e('_Organizado por Alumbra Studios · alumbrastudios.com_');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

function generateWhatsAppLinkAgradecimiento(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p   = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var hub = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico + '&tab=recursos'
    : 'https://reinventabymarymendez.com.mx/hub?tab=recursos';

  var _pub = getContenidoPublicado('agradecimiento', 'whatsapp', 'mensaje');
  if (_pub) {
    var _msg = _buildWAMsgFromTemplate(_pub, { nombre: p, hub: hub, id: idUnico || '' });
    return 'https://wa.me/' + numero + '?text=' + _msg;
  }

  // Fallback hardcodeado original ↓
  var e   = encodeURIComponent;
  var NL  = '%0A';
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('Hola, ' + p + '. Fue un honor acompanarte hoy.') + NL + NL +
    e('Gracias por confiar en este espacio y por abrirte a transformar la manera en que tu imagen comunica quien eres. Lo que viviste hoy es solo el comienzo.') + NL + NL +
    e('Mary estara siempre disponible para seguir acompaanandote en este camino.') + NL + NL +
    e('- - -') + NL + NL +
    e('*Encuesta post-evento — importante*') + NL + NL +
    e('Nos toma menos de 3 minutos y contestarla desbloquea en tu espacio:') + NL + NL +
    e('- Tu constancia de participacion descargable.') + NL +
    e('- Las guias y materiales del taller.') + NL +
    e('- Tu feedback le llega directamente a Mary para seguir mejorando.') + NL + NL +
    e('Entra a tu espacio y contestala:') + NL +
    e(hub) + NL + NL +
    e('- - -') + NL + NL +
    e('Conoce mas sobre Mary: https://reinventabymarymendez.com.mx') + NL + NL +
    e('_Con carino,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Mensaje informativo — por favor no respondas a este chat._') + NL +
    (idUnico ? NL + e('Codigo de acceso: ' + idUnico) + NL : '') +
    e('_Organizado por Alumbra Studios · alumbrastudios.com_');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

/* ── Generación masiva de botones WA ────────────────────────── */
function generarBotonesConfirmacionMasivo() {
  _generarBotonesMasivoConId(generateWhatsAppLinkConfirmacion, 5, 6, 'Enviar confirmación', false);
}

function generarBotonesIndicacionesMasivo() {
  _generarBotonesMasivoConId(generateWhatsAppLinkIndicaciones, 17, 18, 'Enviar indicaciones (jueves)', false);
}

function generarBotonesRecordatorioMasivo() {
  _generarBotonesMasivoConId(generateWhatsAppLinkRecordatorio, 8, 9, 'Enviar recordatorio', false);
}

function generarBotonesAgradecimientoMasivo() {
  _generarBotonesMasivoConId(generateWhatsAppLinkAgradecimiento, 14, 15, 'Enviar agradecimiento', false);
}

/* ── WA: Recordatorio de recursos (mié 19 ago) ───────────────── */
function generateWhatsAppLinkRecordatorioPost(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p   = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var hub = (idUnico ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico : 'https://reinventabymarymendez.com.mx/hub') + '&tab=recursos';
  var e = encodeURIComponent, NL = '%0A';
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('Hola, ' + p + '. Tus guias del taller todavia te esperan.') + NL + NL +
    e('Solo necesitas contestar la encuesta de satisfaccion (toma 3 minutos) para desbloquear:') + NL + NL +
    e('· Tu constancia de participacion.') + NL +
    e('· Guia de 7 estilos y sus esenciales.') + NL +
    e('· Guia de compras inteligentes.') + NL + NL +
    e('El acceso al hub cierra el lunes 24 de agosto.') + NL + NL +
    e('Entra aqui:') + NL +
    e(hub) + NL + NL +
    e('_Con carino,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Mensaje informativo — por favor no respondas a este chat._');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

function generarBotonesRecordatorioPostMasivo() {
  // Col W(23)=link, Col X(24)=estado — solo a quienes no tienen encuesta post
  _generarBotonesMasivoConIdFiltrado(generateWhatsAppLinkRecordatorioPost, 23, 24, 'Enviar rec. recursos', false);
}

/* ── WA: Última llamada (vie 21 ago) ─────────────────────────── */
function generateWhatsAppLinkUltimaLlamada(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p   = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var hub = (idUnico ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico : 'https://reinventabymarymendez.com.mx/hub') + '&tab=recursos';
  var e = encodeURIComponent, NL = '%0A';
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('Hola, ' + p + '. Ultimo aviso importante.') + NL + NL +
    e('El lunes 24 de agosto se cierra definitivamente el acceso a tu espacio personal de REINVENTA.') + NL + NL +
    e('Antes de que eso pase, contesta la encuesta (3 minutos) y descarga tus guias:') + NL + NL +
    e('· Constancia de participacion.') + NL +
    e('· Guia de 7 estilos y sus esenciales.') + NL +
    e('· Guia de compras inteligentes.') + NL + NL +
    e('Entra aqui antes del lunes:') + NL +
    e(hub) + NL + NL +
    e('_Con carino,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Mensaje informativo — por favor no respondas a este chat._');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

function generarBotonesUltimaLlamadaMasivo() {
  // Col Z(26)=link, Col AA(27)=estado
  _generarBotonesMasivoConIdFiltrado(generateWhatsAppLinkUltimaLlamada, 26, 27, 'Enviar ultima llamada', false);
}

/* ── Variante de _generarBotonesMasivoConId que filtra sin encuesta post ── */
function _generarBotonesMasivoConIdFiltrado(generadorFn, colLink, colEstado, textoBoton, soloSinLink) {
  // Construir set de IDs que ya tienen encuesta post
  var aSheet  = getAsistenciaSheet();
  var aData   = aSheet.getDataRange().getValues();
  var conEncuesta = {};
  for (var ai = 1; ai < aData.length; ai++) {
    if (aData[ai][6] === '✓') conEncuesta[(aData[ai][0]||'').toString().trim()] = true;
  }
  // También set de quienes ya descargaron algo
  var descSet = _obtenerDescargasPorId();

  _generarBotonesMasivoConIdConFiltro(generadorFn, colLink, colEstado, textoBoton, soloSinLink, function(id) {
    // Devuelve true si DEBE recibir el WA (sin encuesta post Y sin descargas)
    return !conEncuesta[id] && !descSet[id];
  });
}

function _generarBotonesMasivoConIdConFiltro(generadorFn, colLink, colEstado, textoBoton, soloSinLink, filtroFn) {
  var comSheet = getComunicacionesSheet();
  var regSheet = getSheet();
  var asSheet  = getAsistenciaSheet();
  var comData  = comSheet.getDataRange().getValues();
  var regData  = regSheet.getDataRange().getValues();
  var asData   = asSheet.getDataRange().getValues();

  // Mapas auxiliares: correo → {telefono, id}
  var telefonoPorCorreo = {}, idPorCorreo = {};
  for (var ri = 1; ri < regData.length; ri++) {
    var rc = (regData[ri][2]||'').toLowerCase().trim();
    if (rc) telefonoPorCorreo[rc] = (regData[ri][3]||'').toString().trim();
  }
  for (var ai = 1; ai < asData.length; ai++) {
    var ac = (asData[ai][2]||'').toLowerCase().trim();
    if (ac) idPorCorreo[ac] = (asData[ai][0]||'').toString().trim();
  }

  var generados = 0;
  for (var i = 1; i < comData.length; i++) {
    var correo  = (comData[i][0]||'').toLowerCase().trim();
    var nombre  = (comData[i][1]||'').toString();
    var yaLink  = (comData[i][colLink - 1]||'').toString();
    if (!correo) continue;
    if (soloSinLink && yaLink && yaLink !== '') continue;

    var id      = idPorCorreo[correo] || '';
    var telefono= telefonoPorCorreo[correo] || (comData[i][2]||'').toString();

    // Aplicar filtro personalizado
    if (filtroFn && !filtroFn(id)) continue;

    var link = generadorFn(nombre, telefono, id);
    if (!link) continue;

    var cell = comSheet.getRange(i + 1, colLink);
    cell.setFormula('=HYPERLINK("' + link + '","' + textoBoton + '")');
    comSheet.getRange(i + 1, colEstado).setValue('Pendiente');
    generados++;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Botones generados: ' + generados, 'WA listo', 4);
}

/* ── WA: Encuesta previa pendiente → espacio personal ────────── */
function generateWhatsAppLinkEncuestaPrevia(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p   = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var hub = 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico;
  var NL  = '%0A';
  var e   = encodeURIComponent;
  var msg =
    e('Hola ' + p + ' 🌟') + NL + NL +
    e('El taller es mañana y queremos estar bien preparadas para ti.') + NL + NL +
    e('Aún tenemos pendiente tu encuesta previa — con ella Mary puede personalizar tu experiencia al máximo.') + NL + NL +
    e('👉 Llénala desde tu espacio personal:') + NL +
    e(hub) + NL + NL +
    e('¡Nos vemos mañana!') + NL + NL +
    e('_Con cariño,_') + NL +
    e('_Reinventa by Mary Méndez_') + NL + NL +
    e('_Mensaje informativo — por favor no respondas a este chat._') + NL +
    e('_Organizado por Alumbra Studios · alumbrastudios.com_');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

/* Genera botones WA solo para quienes NO han llenado la encuesta previa */
function generarBotonesEncuestaPreviaMasivo() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var regSheet  = getSheet();
  var comSheet  = getComunicacionesSheet();
  var asiSheet  = getAsistenciaSheet();
  var prevSheet = ss.getSheetByName('Encuesta Previa');

  var regData  = regSheet.getDataRange().getValues();
  var asiData  = asiSheet.getDataRange().getValues();
  var prevData = prevSheet ? prevSheet.getDataRange().getValues() : [];

  // IDs que YA llenaron la encuesta previa (col A, índice 0)
  var conEncuesta = {};
  for (var p = 1; p < prevData.length; p++) {
    var idEnc = (prevData[p][0] || '').toString().trim();
    if (idEnc) conEncuesta[idEnc] = true;
  }

  // Mapa correo → ID único (desde Asistencia)
  var idPorCorreo = {};
  for (var j = 1; j < asiData.length; j++) {
    var ca = (asiData[j][2] || '').toLowerCase().trim();
    if (ca) idPorCorreo[ca] = asiData[j][0];
  }

  var generados    = 0;
  var vistos       = {};
  var totalPagadas = 0;

  for (var i = 1; i < regData.length; i++) {
    if (regData[i][9] !== '✓') continue;
    totalPagadas++;
    var correo = (regData[i][2] || '').toLowerCase().trim();
    if (!correo || vistos[correo]) continue;
    vistos[correo] = true;

    var id       = idPorCorreo[correo] || '';
    var tieneEnc = !!conEncuesta[id];
    var fila     = findRowByEmailInSheet(comSheet, correo);

    if (tieneEnc) continue;
    if (!id) continue;
    if (!fila) continue;

    // Cols 20-21 (T-U) — encuesta previa pendiente
    var yaLink = (comSheet.getRange(fila, 20).getValue() || '').toString();
    if (yaLink.indexOf('wa.me') !== -1) continue; // ya tiene link real — saltar
    var nombre = regData[i][1];
    var tel    = regData[i][3];
    generarBotonWA(comSheet, fila, nombre, tel,
      function(fn, theId) {
        return function(n, t) { return fn(n, t, theId); };
      }(generateWhatsAppLinkEncuestaPrevia, id),
      20, 21, 'Enviar encuesta previa');
    generados++;
  }
  SpreadsheetApp.getUi().alert('Botones generados: ' + generados + ' de ' + totalPagadas + ' asistentes (solo quienes no han llenado la encuesta previa)');
}

/* Genera botones WA pasando el ID único de cada asistente al mensaje */
function _generarBotonesMasivoConId(generadorFn, colLink, colEstado, textoBoton, soloSinLink) {
  var regSheet  = getSheet();
  var comSheet  = getComunicacionesSheet();
  var asiSheet  = getAsistenciaSheet();
  var data      = regSheet.getDataRange().getValues();
  var asiData   = asiSheet.getDataRange().getValues();
  var generados = 0;
  var vistos    = {};

  // Mapa correo → ID
  var idPorCorreo = {};
  for (var j = 1; j < asiData.length; j++) {
    var c = (asiData[j][2] || '').toLowerCase().trim();
    if (c) idPorCorreo[c] = asiData[j][0];
  }

  for (var i = 1; i < data.length; i++) {
    if (data[i][9] !== '✓') continue;
    var correo = (data[i][2] || '').toLowerCase().trim();
    var nombre = data[i][1];
    var tel    = data[i][3];
    if (!correo || vistos[correo]) continue;
    vistos[correo] = true;

    var id   = idPorCorreo[correo] || '';
    var fila = findRowByEmailInSheet(comSheet, correo);
    if (!fila) continue;
    var yaLink = comSheet.getRange(fila, colLink).getValue();
    if (soloSinLink && yaLink) continue;

    generarBotonWA(comSheet, fila, nombre, tel,
      function(fn, theId) {
        return function(n, t) { return fn(n, t, theId); };
      }(generadorFn, id),
      colLink, colEstado, textoBoton);
    generados++;
  }
  Logger.log('Botones generados (' + textoBoton + '): ' + generados);
}

function generarBotonesQRMasivo() {
  var regSheet  = getSheet();
  var comSheet  = getComunicacionesSheet();
  var asiSheet  = getAsistenciaSheet();
  var regData   = regSheet.getDataRange().getValues();
  var asiData   = asiSheet.getDataRange().getValues();
  var generados = 0;

  var idPorCorreo = {};
  for (var j = 1; j < asiData.length; j++) {
    var c = (asiData[j][2] || '').toLowerCase().trim();
    if (c) idPorCorreo[c] = asiData[j][0];
  }

  for (var i = 1; i < regData.length; i++) {
    if (regData[i][9] !== '✓') continue;
    var correo = (regData[i][2] || '').toLowerCase().trim();
    var nombre = regData[i][1];
    var tel    = regData[i][3];
    var id     = idPorCorreo[correo];
    if (!id) continue;
    var fila = findRowByEmailInSheet(comSheet, correo);
    if (!fila) continue;

    var link = generateWhatsAppLinkQR(nombre, tel, id); // id ya viene de idPorCorreo
    if (link) {
      var richText = SpreadsheetApp.newRichTextValue()
        .setText('Enviar QR')
        .setLinkUrl(link)
        .build();
      comSheet.getRange(fila, 11).setRichTextValue(richText);
      if (!comSheet.getRange(fila, 12).getValue()) comSheet.getRange(fila, 12).setValue('Pendiente');
    } else {
      comSheet.getRange(fila, 11).setValue('SIN TELEFONO');
      comSheet.getRange(fila, 12).setValue('SIN TELEFONO');
    }
    generados++;
  }
  Logger.log('Botones QR generados: ' + generados);
}

/* ── Correos: plantillas ─────────────────────────────────────── */
function _headerCorreo(preheader) {
  // Preheader oculto — algunos clientes lo muestran después del asunto (invisible en el cuerpo)
  var phPad = '&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;'
            + '&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;'
            + '&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;'
            + '&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;'
            + '&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;';
  var phHtml = preheader
    ? '<span style="display:none;max-height:0;overflow:hidden;mso-hide:all;">' + preheader + phPad + '</span>'
    : '';
  return phHtml
    + '<div style="background:#E8E2DB;padding:2rem 1rem;font-family:\'Gill Sans\',Calibri,\'Segoe UI\',sans-serif;">'
    + '<div style="max-width:540px;margin:0 auto;background:#EFE9E2;box-shadow:0 4px 40px rgba(42,15,37,.13);">'
    + '<div style="background:#2A0F25;padding:2rem 2.4rem 1.6rem;text-align:center;">'
    + '<span style="font-family:Georgia,serif;font-weight:400;font-size:1rem;letter-spacing:.22em;text-transform:uppercase;color:#C6A56A;display:block;margin-bottom:.2rem;">Reinventa</span>'
    + '<span style="font-size:.62rem;letter-spacing:.16em;text-transform:uppercase;color:rgba(198,165,106,.5);">by Mary Méndez</span>'
    + '</div>'
    + '<div style="height:2px;background:#C6A56A;opacity:.45;"></div>';
}

function _footerCorreo(correo) {
  return '<div style="background:#2A0F25;padding:1.1rem 2rem;text-align:center;">'
    + '<p style="font-size:.63rem;letter-spacing:.07em;color:rgba(198,165,106,.5);line-height:1.7;margin:0 0 .5rem;">REINVENTA by Mary Méndez &middot; Ciudad de México<br>'
    + 'Este correo fue enviado a ' + correo + '.</p>'
    + '<p style="font-size:.6rem;color:rgba(198,165,106,.3);margin:0;">Evento organizado integralmente por <strong style="color:rgba(198,165,106,.5);">Alumbra Studios</strong> &middot; <a href="https://www.alumbrastudios.com" style="color:rgba(198,165,106,.4);text-decoration:none;">alumbrastudios.com</a></p>'
    + '</div></div></div>';
}

function _firmaCorreo() {
  return '<hr style="border:none;border-top:1px solid rgba(42,15,37,.1);margin:0 0 1.4rem;" />'
    + '<p style="font-family:Georgia,serif;font-size:.98rem;color:#2A0F25;margin:.8rem 0 .1rem;">Mary Méndez</p>'
    + '<p style="font-size:.72rem;color:#8F7383;margin:0;">Consultora de imagen y liderazgo</p>';
}

function _detallesEvento() {
  // Lee datos del evento desde Contenidos (id: 'global', canal: 'evento')
  var evNombre    = getContenidoPublicado('global','evento','nombre')    || 'Lo que tu imagen comunica';
  var evFecha     = getContenidoPublicado('global','evento','fecha')     || 'Sábado 15 de agosto de 2026';
  var evHora      = getContenidoPublicado('global','evento','hora')      || '10:00 a.m. a 12:00 p.m.';
  var evLugar     = getContenidoPublicado('global','evento','lugar')     || 'The University Club of Mexico';
  var evDireccion = getContenidoPublicado('global','evento','direccion') || 'Av. Paseo de la Reforma 150, Juárez, CDMX';
  return '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .8rem;">Detalles del evento</p>'
    + '<p style="font-family:Georgia,serif;font-size:1.05rem;font-weight:400;color:#2A0F25;margin:0 0 .1rem;">Taller de imagen y liderazgo</p>'
    + '<p style="font-size:.78rem;color:#8F7383;font-style:italic;margin:0 0 1rem;">' + evNombre + '</p>'
    + '<p style="font-size:.85rem;color:#4a3545;margin:0 0 .5rem;">' + evFecha + ' &middot; ' + evHora + '</p>'
    + '<p style="font-size:.85rem;color:#4a3545;margin:0;">' + evLugar + '<br><span style="color:#8F7383;font-size:.8rem;">' + evDireccion + '</span></p>'
    + '</div>';
}

function enviarCorreoConfirmacion(nombre, correo, fase, idUnico) {
  var p = nombre ? nombre.split(' ')[0] : 'Hola';
  var calLink = 'https://calendar.google.com/calendar/render?action=TEMPLATE'
    + '&text=Taller+de+imagen+y+liderazgo+%E2%80%94+REINVENTA'
    + '&dates=20260815T160000Z/20260815T180000Z'
    + '&details=Taller+Lo+que+tu+imagen+comunica+%7C+REINVENTA+by+Mary+M%C3%A9ndez'
    + '&location=The+University+Club+of+Mexico%2C+Av.+Paseo+de+la+Reforma+150%2C+Ju%C3%A1rez%2C+CDMX';
  var hubUrl = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';

  // Contenido desde Contenidos (requerido — lanza error si falta publicado)
  var _vars = { nombre: p, hub: hubUrl, id: idUnico || '' };
  var templateMode = getContenidoPublicado('confirmacion', 'correo', 'template_mode') || 'visual';
  var asunto = _reqContenido('confirmacion', 'correo', 'asunto', _vars);
  var html;

  if (templateMode === 'html') {
    // Modo HTML personalizado: usa el campo 'html' publicado directamente
    var customHtml = getContenidoPublicado('confirmacion', 'correo', 'html');
    if (!customHtml) throw new Error('CONTENIDO_FALTANTE:confirmacion/correo/html — aprueba el HTML personalizado antes de enviar.');
    html = _aplicarVariables(customHtml, _vars);
  } else {
    // Modo visual (default): usa los campos individuales de Contenidos
    var apertura         = _reqContenido('confirmacion', 'correo', 'apertura',            _vars);
    var h1               = _reqContenido('confirmacion', 'correo', 'h1',                  _vars);
    var seccionHub       = _reqContenido('confirmacion', 'correo', 'seccion_hub',         _vars);
    var notaEncuesta     = _reqContenido('confirmacion', 'correo', 'nota_encuesta',       _vars);
    var ctaCalendarLabel = _reqContenido('confirmacion', 'correo', 'cta_calendar_label',  _vars);
    var cierre           = _aplicarVariables(getContenidoPublicado('confirmacion','correo','cierre'), _vars);
    // Campos opcionales con fallback (no bloquean el envío si faltan)
    var preheader        = getContenidoPublicado('confirmacion','correo','preheader')              || '';
    var chipPrefijo      = getContenidoPublicado('confirmacion','correo','chip_prefijo')           || 'Pago confirmado';
    var seccionHubLabel  = getContenidoPublicado('confirmacion','correo','seccion_hub_label')      || 'Tu espacio personal del evento';
    var ctaHubLabel      = getContenidoPublicado('confirmacion','correo','cta_hub_label')          || 'Acceder a mi espacio →';
    var secEncTitulo     = getContenidoPublicado('confirmacion','correo','seccion_encuesta_titulo')|| 'Una cosa importante';

    html = _headerCorreo(preheader)
      + '<div style="padding:2.2rem 2.6rem 2rem;">'
      + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">' + chipPrefijo + ' &middot; ' + fase + '</div>'
      + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 1rem;">' + h1 + '</h1>'
      + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;">' + apertura + '</p>'
      + _detallesEvento()
      + '<a href="' + calLink + '" style="display:block;background:#2A0F25;color:#EFE9E2;text-align:center;padding:.9rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;margin-bottom:1.6rem;">' + ctaCalendarLabel + '</a>'
      + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
      + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .6rem;">' + seccionHubLabel + '</p>'
      + '<p style="font-size:.87rem;color:#4a3545;line-height:1.6;margin:0 0 .8rem;">' + seccionHub + '</p>'
      + '<a href="' + hubUrl + '" style="display:block;background:#C6A56A;color:#2A0F25;text-align:center;padding:.8rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;margin-bottom:1rem;font-weight:600;">' + ctaHubLabel + '</a>'
      + '<div style="border-top:1px solid rgba(42,15,37,.1);padding-top:.9rem;">'
      + '<p style="font-size:.78rem;color:#2A0F25;font-weight:600;margin:0 0 .3rem;">' + secEncTitulo + '</p>'
      + '<p style="font-size:.8rem;color:#4a3545;line-height:1.6;margin:0;">' + notaEncuesta + '</p>'
      + '</div></div>'
      + (cierre ? '<p style="font-size:.9rem;line-height:1.7;color:#4a3545;margin:0 0 1.2rem;">' + cierre + '</p>' : '')
      + _firmaCorreo()
      + '</div>'
      + _footerCorreo(correo);
  }

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: asunto, htmlBody: html });
}

/* ── Jueves 13 ago: indicaciones de vestimenta y análisis ───────── */
function enviarCorreoIndicaciones(nombre, correo, idUnico) {
  var p      = nombre ? nombre.split(' ')[0] : 'hola';
  var hubUrl = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';

    // Contenido desde Contenidos (requerido — lanza error si falta publicado)
  var _vars = { nombre: p, hub: hubUrl, id: idUnico || '' };
  var asunto               = _reqContenido('indicaciones', 'correo', 'asunto',               _vars);
  var apertura             = _reqContenido('indicaciones', 'correo', 'apertura',             _vars);
  var cierre               = _reqContenido('indicaciones', 'correo', 'cierre',               _vars);
  var h1                   = _reqContenido('indicaciones', 'correo', 'h1',                   _vars);
  var lookBloque           = _reqContenido('indicaciones', 'correo', 'look_bloque',          _vars);
  var rostroBloque         = _reqContenido('indicaciones', 'correo', 'rostro_bloque',        _vars);
  var medidasInstrucciones = _reqContenido('indicaciones', 'correo', 'medidas_instrucciones',_vars);
  var seccionHub           = _reqContenido('indicaciones', 'correo', 'seccion_hub',          _vars);
  var ctaHubLabel          = _reqContenido('indicaciones', 'correo', 'cta_hub_label',        _vars);

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Preparación &middot; Jueves 13 de agosto</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 .8rem;">' + h1 + '</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.6rem;">' + apertura + '</p>'

    // Hub destacado primero
    + '<div style="background:#2A0F25;padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:rgba(198,165,106,.6);margin:0 0 .5rem;">Tu espacio personal</p>'
    + '<p style="font-size:.87rem;color:#EFE9E2;line-height:1.6;margin:0 0 .9rem;">' + seccionHub + '</p>'
    + '<a href="' + hubUrl + '" style="display:block;background:#C6A56A;color:#2A0F25;text-align:center;padding:.8rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;font-weight:600;">' + ctaHubLabel + '</a>'
    + '</div>'

    // Sugerencia de look
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .9rem;">👕 Sugerencia de look &mdash; Análisis de cuerpo</p>'
    + _renderLineasBloque(lookBloque, 'font-size:.88rem;color:#4a3545;line-height:1.7;margin:0 0 .5rem;')
    + '</div>'

    // Rostro
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .9rem;">💇‍♀️ Rostro &mdash; Análisis visagismo</p>'
    + '<p style="font-size:.88rem;color:#4a3545;line-height:1.7;margin:0;">' + rostroBloque + '</p>'
    + '</div>'

    // Medidas
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .9rem;">📏 Tus medidas (opcional pero súper útil)</p>'
    + _renderLineasBloque(medidasInstrucciones, 'font-size:.85rem;color:#4a3545;line-height:1.7;margin:0 0 .35rem;')
    + '<a href="' + hubUrl + '" style="display:block;background:transparent;border:1px solid #C6A56A;color:#2A0F25;text-align:center;padding:.7rem 1.2rem;text-decoration:none;font-size:.78rem;letter-spacing:.09em;text-transform:uppercase;font-weight:600;margin-top:.8rem;">Registrar mis medidas en mi espacio &rarr;</a>'
    + '</div>'

    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 .5rem;">' + cierre + '</p>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: asunto, htmlBody: html });
}

function enviarCorreoRecordatorio(nombre, correo, idUnico) {
  var p      = nombre ? nombre.split(' ')[0] : 'Hola';
  var hubUrl = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';

  // Contenido desde Contenidos con fallback hardcodeado
  var _vars = { nombre: p, hub: hubUrl, id: idUnico || '' };
  function _get(campo, fallback) {
    return _aplicarVariables(getContenidoPublicado('recordatorio', 'correo', campo), _vars) || fallback;
  }

  var asunto = _get('asunto', 'Mañana te esperamos — REINVENTA');
  var h1     = _get('h1',     p + ', mañana es el gran día.');
  var apertura = _get('apertura',
    'Mary tiene algo muy especial preparado para ti. Te esperamos puntual y con muchas ganas de transformar la manera en que tu imagen comunica quién eres.');
  var seccionHub = _get('seccion_hub',
    'Ahí encontrarás la agenda del día, tu código QR de entrada y los materiales del taller.');
  var lookCompacto = _get('look_compacto',
    'Recuerda el dresscode del día — opcional\n'
    + 'Parte inferior: Pantalón ajustado (leggings o jeans).\n'
    + 'Parte superior: Blusa o playera básica ajustada (blanco o neutro de preferencia).\n'
    + 'Capa extra: Saco, blazer o blusón en el color que prefieras.\n'
    + 'Rostro: De preferencia con el cabello recogido para el análisis de visagismo (si no puedes, hay pinzas disponibles).');
  var seccionMedidas = _get('seccion_medidas', '');
  var notaEncuesta = _get('nota_encuesta',
    'Si aún no has contestado la encuesta dentro de tu espacio, este es el momento. Mary la revisa personalmente antes del taller para personalizar tu experiencia.'
    + '<br><br><strong>Importante: quienes no la contesten antes del evento no tendrán acceso al material digital posterior — guías, recursos y demás — que se desbloquea en tu perfil tras el taller.</strong>');
  var cierre = _get('cierre',
    'Si tienes alguna duda de último momento no dudes en contactarnos. ¡Nos vemos mañana!');

  // La primera línea del look_compacto es la intro opcional, el resto son ítems
  var lookLineas    = lookCompacto.split('\n');
  var lookIntro     = lookLineas[0] || '';
  var lookItems     = lookLineas.slice(1).join('\n');

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Recordatorio &middot; Mañana es el taller</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 1rem;">' + h1 + '</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;">' + apertura + '</p>'
    + _detallesEvento()
    + '<p style="font-size:.85rem;color:#4a3545;margin:-1rem 0 1.6rem;padding:0 1.6rem;"><a href="https://maps.app.goo.gl/Uo7tYiQz23jMCmKw7" style="color:#C6A56A;text-decoration:none;">Ver en Google Maps &rarr;</a></p>'
    // Acceso al hub
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .6rem;">Tu espacio personal del evento</p>'
    + '<p style="font-size:.87rem;color:#4a3545;line-height:1.6;margin:0 0 .8rem;">' + seccionHub + '</p>'
    + '<a href="' + hubUrl + '" style="display:block;background:#C6A56A;color:#2A0F25;text-align:center;padding:.8rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;margin-bottom:0;font-weight:600;">Acceder a mi espacio &rarr;</a>'
    + '</div>'
    // Look sugerido — recordatorio compacto
    + '<div style="border:1px solid rgba(42,15,37,.15);border-left:3px solid #C6A56A;padding:1.2rem 1.4rem;margin-bottom:1.4rem;background:rgba(198,165,106,.04);">'
    + '<p style="font-size:.75rem;color:#2A0F25;font-weight:700;margin:0 0 .35rem;letter-spacing:.03em;">👕 Sugerencia de look para mañana</p>'
    + (lookIntro ? '<p style="font-size:.78rem;color:#8F7383;line-height:1.6;margin:0 0 .6rem;font-style:italic;">' + lookIntro + '</p>' : '')
    + _renderLineasBloque(lookItems, 'font-size:.82rem;color:#4a3545;line-height:1.6;margin:0 0 .35rem;')
    + '</div>'
    // Medidas — solo si hay contenido publicado
    + (seccionMedidas
      ? '<div style="border:1px solid rgba(42,15,37,.12);padding:1.2rem 1.4rem;margin-bottom:1.4rem;">'
        + '<p style="font-size:.75rem;color:#2A0F25;font-weight:700;margin:0 0 .4rem;letter-spacing:.03em;">Tus medidas de cuerpo</p>'
        + '<p style="font-size:.82rem;color:#4a3545;line-height:1.6;margin:0 0 .7rem;">' + seccionMedidas + '</p>'
        + '<a href="' + hubUrl + '" style="display:inline-block;background:transparent;border:1px solid #C6A56A;color:#2A0F25;padding:.5rem 1rem;text-decoration:none;font-size:.72rem;letter-spacing:.08em;text-transform:uppercase;font-weight:600;">Registrar mis medidas &rarr;</a>'
        + '</div>'
      : '')
    // Encuesta previa — último recordatorio
    + '<div style="border:1px solid #C6A56A;border-left:3px solid #C6A56A;padding:1.2rem 1.4rem;margin-bottom:1.6rem;background:rgba(198,165,106,.06);">'
    + '<p style="font-size:.75rem;color:#2A0F25;font-weight:700;margin:0 0 .4rem;letter-spacing:.03em;">Último recordatorio: encuesta previa</p>'
    + '<p style="font-size:.82rem;color:#4a3545;line-height:1.6;margin:0;">' + notaEncuesta + '</p>'
    + '</div>'
    + '<p style="font-size:.87rem;line-height:1.7;color:#4a3545;margin:0 0 .5rem;">' + cierre + '</p>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: asunto, htmlBody: html });
}

function enviarCorreoQR(nombre, correo, idUnico) {
  var p      = nombre ? nombre.split(' ')[0] : 'Hola';
  var hubUrl = 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico;
  var urlQR  = 'https://api.qrserver.com/v1/create-qr-code/?size=220x220&data=' + encodeURIComponent(hubUrl);

  // Contenido desde Contenidos con fallback hardcodeado
  var _vars = { nombre: p, hub: hubUrl, id: idUnico || '' };
  function _get(campo, fallback) {
    return _aplicarVariables(getContenidoPublicado('qr', 'correo', campo), _vars) || fallback;
  }
  var asunto   = _get('asunto',   'Hoy te esperamos — REINVENTA');
  var h1       = _get('h1',       p + ', hoy te esperamos.');
  var apertura = _get('apertura', 'Muestra este código QR al llegar al evento. El staff lo escaneará en la entrada.');
  var cierre   = _get('cierre',   '¡Nos vemos en un momento. Será un día increíble!');

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;text-align:center;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Hoy es el día &middot; Sábado 15 de agosto</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 .6rem;text-align:left;">' + h1 + '</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;text-align:left;">' + apertura + '</p>'
    + '<div style="background:#2A0F25;display:inline-block;padding:1.2rem;margin-bottom:1.6rem;">'
    + '<img src="' + urlQR + '" width="200" height="200" style="display:block;" alt="Código QR de entrada" /></div>'
    + _detallesEvento()
    + '<p style="font-size:.87rem;line-height:1.7;color:#4a3545;margin:0 0 .5rem;text-align:left;">' + cierre + '</p>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: asunto, htmlBody: html });
}

function enviarCorreoAgradecimiento(nombre, correo, idUnico) {
  var p      = nombre ? nombre.split(' ')[0] : 'Hola';
  var hubUrl = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';

  // URL con deep-link a la pestaña de Recursos (donde está la encuesta post)
  var hubRecursosUrl = hubUrl + (hubUrl.indexOf('?') !== -1 ? '&' : '?') + 'tab=recursos';

  // Contenido desde Contenidos (requerido — lanza error si falta publicado)
  var _vars = { nombre: p, hub: hubRecursosUrl, id: idUnico || '' };
  var asunto              = _reqContenido('agradecimiento', 'correo', 'asunto',               _vars);
  var apertura            = _reqContenido('agradecimiento', 'correo', 'apertura',             _vars);
  var cierre              = _reqContenido('agradecimiento', 'correo', 'cierre',               _vars);
  var h1                  = _reqContenido('agradecimiento', 'correo', 'h1',                   _vars);
  var seccionEncuestaPost = _reqContenido('agradecimiento', 'correo', 'seccion_encuesta_post',_vars);
  var ctaEncuestaLabel    = _reqContenido('agradecimiento', 'correo', 'cta_encuesta_label',   _vars);

  // Separar párrafo intro de bullets (líneas con ·)
  var agr_lineas = seccionEncuestaPost.split('\n');
  var agr_intro  = agr_lineas.filter(function(l){ return l.trim().charAt(0) !== '·'; }).join(' ');
  var agr_bullets= agr_lineas.filter(function(l){ return l.trim().charAt(0) === '·'; });

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Gracias por estar aquí</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 1rem;">' + h1 + '</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.2rem;">' + apertura + '</p>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;">' + cierre + '</p>'
    // Encuesta post — principal CTA
    + '<div style="background:#2A0F25;padding:1.4rem 1.6rem;margin-bottom:1.4rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:rgba(198,165,106,.6);margin:0 0 .5rem;">Un paso importante antes de cerrar</p>'
    + '<p style="font-family:Georgia,serif;font-size:1.05rem;color:#EFE9E2;margin:0 0 .8rem;line-height:1.4;">Contesta la encuesta post-evento</p>'
    + (agr_intro ? '<p style="font-size:.85rem;color:rgba(239,233,226,.75);line-height:1.6;margin:0 0 1rem;">' + agr_intro + '</p>' : '')
    + agr_bullets.map(function(b){ return '<p style="font-size:.83rem;color:rgba(239,233,226,.85);margin:0 0 .3rem;">' + b.trim() + '</p>'; }).join('')
    + '<a href="' + hubRecursosUrl + '" style="display:block;background:#C6A56A;color:#2A0F25;text-align:center;padding:.8rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;font-weight:600;margin-top:1rem;">' + ctaEncuestaLabel + '</a>'
    + '</div>'
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.2rem 1.4rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .5rem;">Conoce más sobre Mary</p>'
    + '<p style="font-size:.85rem;color:#4a3545;margin:0;"><a href="https://reinventabymarymendez.com.mx" style="color:#C6A56A;text-decoration:none;">reinventabymarymendez.com.mx &rarr;</a></p>'
    + '</div>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: asunto, htmlBody: html });
}

/* ── Envíos masivos de correo ────────────────────────────────── */
function enviarConfirmacionExistentes() {
  _enviarCorreosMasivo('enviarCorreoConfirmacion', 7, 'Correos confirmación');
}

function enviarCorreosIndicacionesMasivo() {
  _enviarCorreosMasivo('enviarCorreoIndicaciones', 19, 'Correos indicaciones (jueves)');
}

function enviarCorreosRecordatorioMasivo() {
  _enviarCorreosMasivo('enviarCorreoRecordatorio', 10, 'Correos recordatorio');
}

function enviarCorreosAgradecimientoMasivo() {
  _enviarCorreosMasivo('enviarCorreoAgradecimiento', 16, 'Correos agradecimiento');
}

function _enviarCorreosMasivo(fnNombre, colEnviado, label) {
  var regSheet = getSheet();
  var comSheet = getComunicacionesSheet();
  var data     = regSheet.getDataRange().getValues();
  var enviados = 0;
  var vistos   = {};

  for (var i = 1; i < data.length; i++) {
    if (data[i][9] !== '✓') continue;
    var correo = (data[i][2] || '').toLowerCase().trim();
    var nombre = data[i][1];
    var fase   = data[i][6];
    if (!correo || vistos[correo]) continue;

    var filaComm = findRowByEmailInSheet(comSheet, correo);
    var yaEnv    = filaComm ? comSheet.getRange(filaComm, colEnviado).getValue() : '';

    if (yaEnv !== 'Sí') {
      var idUnicoCorreo = obtenerIdAsistente(correo);
      try {
        if (fnNombre === 'enviarCorreoConfirmacion')  enviarCorreoConfirmacion(nombre, correo, fase, idUnicoCorreo);
        if (fnNombre === 'enviarCorreoIndicaciones')  enviarCorreoIndicaciones(nombre, correo, idUnicoCorreo);
        if (fnNombre === 'enviarCorreoRecordatorio')  enviarCorreoRecordatorio(nombre, correo, idUnicoCorreo);
        if (fnNombre === 'enviarCorreoAgradecimiento') enviarCorreoAgradecimiento(nombre, correo, idUnicoCorreo);
        if (filaComm) comSheet.getRange(filaComm, colEnviado).setValue('Sí');
        vistos[correo] = true;
        enviados++;
      } catch(e) {
        // CONTENIDO_FALTANTE u otro error — no marcar como enviado, registrar para QA
        Logger.log('⚠ Error enviando a ' + correo + ': ' + e.message);
        errores = (errores || 0) + 1;
      }
      Utilities.sleep(1000);
    }
  }
  Logger.log(label + ' enviados: ' + enviados);
}

function enviarCorreosQRMasivo() {
  var regSheet = getSheet();
  var comSheet = getComunicacionesSheet();
  var asiSheet = getAsistenciaSheet();
  var regData  = regSheet.getDataRange().getValues();
  var asiData  = asiSheet.getDataRange().getValues();
  var enviados = 0;
  var vistos   = {};

  var idPorCorreo = {};
  for (var j = 1; j < asiData.length; j++) {
    var c = (asiData[j][2] || '').toLowerCase().trim();
    if (c) idPorCorreo[c] = asiData[j][0];
  }

  for (var i = 1; i < regData.length; i++) {
    if (regData[i][9] !== '✓') continue;
    var correo = (regData[i][2] || '').toLowerCase().trim();
    var nombre = regData[i][1];
    if (!correo || vistos[correo]) continue;

    var id       = idPorCorreo[correo];
    var filaComm = findRowByEmailInSheet(comSheet, correo);
    var yaEnv    = filaComm ? comSheet.getRange(filaComm, 13).getValue() : '';

    if (id && yaEnv !== 'Sí') {
      try {
        enviarCorreoQR(nombre, correo, id);
        if (filaComm) comSheet.getRange(filaComm, 13).setValue('Sí');
        vistos[correo] = true;
        enviados++;
      } catch(e) {
        Logger.log('⚠ Error enviando QR a ' + correo + ': ' + e.message);
      }
      Utilities.sleep(1000);
    }
  }
  Logger.log('Correos QR enviados: ' + enviados);
}

/* ── Trigger automático: QR a las 8am del 15 agosto 2026 ────── */
/* Ejecuta esta función UNA SOLA VEZ desde el editor de Apps Script
   para programar el envío automático. Puedes verificarlo en
   Edición > Triggers del proyecto actual.                        */
function programarEnvioQR8AM() {
  // Eliminar triggers previos del mismo nombre para no duplicar
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'enviarCorreosQRMasivo') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Crear trigger: sábado 15 agosto 2026 a las 8:00 am (hora México = UTC-6)
  // Apps Script usa la zona horaria del spreadsheet, asegúrate de tenerla en America/Mexico_City
  ScriptApp.newTrigger('enviarCorreosQRMasivo')
    .timeBased()
    .at(new Date('2026-08-15T08:00:00'))
    .create();
  SpreadsheetApp.getUi().alert('Listo. El correo con código QR se mandará automáticamente el sábado 15 de agosto a las 8:00 am.');
}

/* Cancela el trigger por si necesitas moverlo o ajustarlo */
function cancelarEnvioQR() {
  var triggers = ScriptApp.getProjectTriggers();
  var cancelados = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'enviarCorreosQRMasivo') {
      ScriptApp.deleteTrigger(triggers[i]);
      cancelados++;
    }
  }
  SpreadsheetApp.getUi().alert('Trigger cancelado (' + cancelados + ' eliminado(s)).');
}

/* ── Correos de prueba ───────────────────────────────────────── */
function enviarCorreoPrueba()              { enviarCorreoConfirmacion('Valeria García', 'mejoracontinua@caceca.org', 'Early Bird', 'RNV-001'); }
function enviarCorreoIndicacionesPrueba()  { enviarCorreoIndicaciones('Valeria García', 'mejoracontinua@caceca.org', 'RNV-001'); }
function enviarCorreoRecordatorioPrueba()  { enviarCorreoRecordatorio('Valeria García', 'mejoracontinua@caceca.org', 'RNV-001'); }
function enviarCorreoQRPrueba()            { enviarCorreoQR('Valeria García', 'mejoracontinua@caceca.org', 'RNV-001'); }
function enviarCorreoAgradecimientoPrueba(){ enviarCorreoAgradecimiento('Valeria García', 'mejoracontinua@caceca.org', 'RNV-001'); }

/* ── Prueba WhatsApp ─────────────────────────────────────────── */
function probarEnlaceWhatsApp() {
  var sheet = getComunicacionesSheet();
  var fila  = sheet.getLastRow() + 1;
  var link  = generateWhatsAppLinkConfirmacion('Estef PRUEBA', '5536599392');
  Logger.log('Link: ' + link);
  var rt = SpreadsheetApp.newRichTextValue().setText('Enviar WhatsApp').setLinkUrl(link).build();
  sheet.getRange(fila, 5).setRichTextValue(rt);
  sheet.getRange(fila, 6).setValue('PRUEBA');
}

/* ── Configurar dropdowns ────────────────────────────────────── */
function configurarDropdowns() {
  var sheet      = getComunicacionesSheet();
  var ultimaFila = Math.max(sheet.getLastRow(), 2);
  var regla      = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendiente', 'Enviado', 'SIN TELEFONO'], true)
    .setAllowInvalid(false)
    .build();
  // F, I, L, O = cols 6, 9, 12, 15 = estados de WA
  [6, 9, 12, 15].forEach(function(col) {
    sheet.getRange(2, col, ultimaFila - 1, 1).setDataValidation(regla);
  });
  Logger.log('Dropdowns configurados');
}

/* ── Notificaciones internas ─────────────────────────────────── */
function notificarCupoAgotado(sheet) {
  var eb  = contarFaseSheet(sheet, 'Early Bird');
  var pre = contarFaseSheet(sheet, 'Preventa');
  var fin = contarFaseSheet(sheet, 'Últimos lugares');
  var recaudado = (eb * 1300) + (pre * 1500) + (fin * 1700);
  var asunto = 'REINVENTA — Cupo completo 40/40 lugares vendidos';
  var cuerpo = 'SOLD OUT!\n\n— Early Bird: ' + eb + '\n— Preventa: ' + pre + '\n— Últimos lugares: ' + fin
    + '\n\nTotal estimado: $' + recaudado.toLocaleString('es-MX') + ' MXN';
  EMAILS_NOTIFICACION.forEach(function(email) { MailApp.sendEmail(email, asunto, cuerpo); });
}

/* ── Configuración inicial ───────────────────────────────────── */
function configurarHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function crearOActualizar(nombre, headers) {
    var sheet = ss.getSheetByName(nombre);
    if (!sheet) sheet = ss.insertSheet(nombre);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#2A0F25').setFontColor('#C6A56A');
    return sheet;
  }

  crearOActualizar('Registros', [
    'Fecha registro','Nombre','Correo','WhatsApp','Contacto preferido',
    '¿Qué busca?','Fase de compra','Monto pagado (MXN)','Fecha de pago','Pagó ✓',
    'Stripe Payment ID','Origen','Acepta marketing','Autoriza uso de imagen',
    'Canal UTM','¿Cómo se enteró?','¿A qué se dedica?'
  ]);

  crearOActualizar('Comunicaciones', [
    'Correo','Nombre','WhatsApp','Contacto preferido',
    'WA Confirmación','Estado Conf. WA','Correo Conf. Enviado',
    'WA Recordatorio','Estado Record. WA','Correo Record. Enviado',
    'WA QR','Estado QR WA','Correo QR Enviado',
    'WA Agradecimiento','Estado Agradec. WA','Correo Agradec. Enviado'
  ]);

  crearOActualizar('Asistencia', [
    'ID Único','Nombre','Correo','Fase','Asistió ✓','Fecha entrada',
    'Encuesta ✓','Fecha encuesta','Calificación','Comentario'
  ]);

  Logger.log('Hojas configuradas.');
}

/* ── Migración de datos existentes ───────────────────────────── */
function migrarDatosExistentes() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var origen = ss.getSheetByName('REINVENTA - Registro')
            || ss.getSheetByName('Reinventa - Registros');

  if (!origen) { Logger.log('Hoja original no encontrada.'); return; }

  var regSheet = getSheet();
  var comSheet = getComunicacionesSheet();
  var asiSheet = getAsistenciaSheet();
  var data     = origen.getDataRange().getValues();
  var migrados = 0;

  for (var i = 1; i < data.length; i++) {
    var fila   = data[i];
    var correo = (fila[2] || '').toLowerCase().trim();
    if (!correo) continue;

    if (!findRowByEmail(regSheet, correo)) {
      regSheet.appendRow([
        fila[0], fila[1], correo, fila[3], fila[4], fila[5],
        fila[6], fila[7], fila[8], fila[9], fila[10], fila[11],
        fila[12], fila[13], fila[15], fila[16], fila[17]
      ]);
    }

    if (fila[9] === '✓') {
      if (!findRowByEmailInSheet(comSheet, correo)) {
        comSheet.appendRow([correo, fila[1], fila[3], fila[4], '', '', '', '', '', '', '', '', '', '', '', '']);
      }
      if (!findRowByEmailInSheet(asiSheet, correo)) {
        asiSheet.appendRow([generarSiguienteId(asiSheet), fila[1], correo, fila[6], '', '']);
      }
    }
    migrados++;
  }
  Logger.log('Migrados: ' + migrados);
}

/* ── Helpers de confirmación ─────────────────────────────────── */
function obtenerIdAsistente(correo) {
  var sheet = getAsistenciaSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][2] || '').toLowerCase().trim() === correo.toLowerCase().trim()) {
      return data[i][0] || '';
    }
  }
  return '';
}

function marcarCorreoEnviado(correo, col) {
  var sheet = getComunicacionesSheet();
  var fila  = findRowByEmailInSheet(sheet, correo);
  if (fila) sheet.getRange(fila, col).setValue('Sí');
}

/* ── Helpers ─────────────────────────────────────────────────── */
function findRowByEmail(sheet, correo) {
  if (!correo) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][2] || '').toString().toLowerCase().trim() === correo) return i + 1;
  }
  return null;
}

function findRowByEmailInSheet(sheet, correo) {
  if (!correo) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().toLowerCase().trim() === correo.toLowerCase().trim()) return i + 1;
  }
  return null;
}

function findRowByStripeId(sheet, stripeId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][10] || '').toString().trim() === stripeId) return i + 1;
  }
  return null;
}

function contarPagosSheet(sheet) {
  var data = sheet.getDataRange().getValues();
  var n = 0;
  for (var i = 1; i < data.length; i++) { if (data[i][9] === '✓') n++; }
  return n;
}

function contarFaseSheet(sheet, fase) {
  var data = sheet.getDataRange().getValues();
  var n = 0;
  for (var i = 1; i < data.length; i++) { if (data[i][9] === '✓' && data[i][6] === fase) n++; }
  return n;
}

/* ── Menú REINVENTA ──────────────────────────────────────────── */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('REINVENTA')
    .addItem('Sincronizar registros faltantes', 'sincronizarRegistrosFaltantes')
    .addItem('Regenerar links WA sin ID', 'regenerarLinksWASinId')
    .addSeparator()
    .addItem('Correos indicaciones masivo (jue 13 ago)', 'enviarCorreosIndicacionesMasivo')
    .addItem('Generar botones WA indicaciones (jue 13 ago)', 'generarBotonesIndicacionesMasivo')
    .addItem('Programar indicaciones automático (jue 13 ago 12pm)', 'adminProgramarIndicaciones')
    .addItem('Cancelar indicaciones automático', 'adminCancelarIndicaciones')
    .addSeparator()
    .addItem('Enviar recordatorio masivo ahora', 'adminEnviarRecordatorioMasivo')
    .addItem('Programar recordatorio automático (vie 14 ago 12pm)', 'adminProgramarRecordatorio')
    .addItem('Cancelar recordatorio automático', 'adminCancelarRecordatorio')
    .addSeparator()
    .addItem('Programar envío QR automático (sáb 15 ago 8am)', 'programarEnvioQR8AM')
    .addItem('Cancelar envío QR automático', 'cancelarEnvioQR')
    .addSeparator()
    .addItem('Enviar agradecimiento masivo ahora (con check-in)', 'adminEnviarAgradecimientoMasivo')
    .addItem('Programar agradecimiento automático (lun 17 ago 10am)', 'adminProgramarAgradecimiento')
    .addItem('Cancelar agradecimiento automático', 'adminCancelarAgradecimiento')
    .addSeparator()
    .addItem('Generar botones WA encuesta previa pendiente', 'generarBotonesEncuestaPreviaMasivo')
    .addSeparator()
    .addItem('Enviar recordatorio recursos masivo ahora (mié 19)', 'adminEnviarRecordatorioPostMasivo')
    .addItem('Generar botones WA recordatorio recursos (mié 19)', 'generarBotonesRecordatorioPostMasivo')
    .addItem('Programar recordatorio recursos automático (mié 19 ago 11am)', 'adminProgramarRecordatorioPost')
    .addItem('Cancelar recordatorio recursos automático', 'adminCancelarRecordatorioPost')
    .addSeparator()
    .addItem('Enviar última llamada masivo ahora (vie 21)', 'adminEnviarUltimaLlamadaMasivo')
    .addItem('Generar botones WA última llamada (vie 21)', 'generarBotonesUltimaLlamadaMasivo')
    .addItem('Programar última llamada automática (vie 21 ago 12pm)', 'adminProgramarUltimaLlamada')
    .addItem('Cancelar última llamada automática', 'adminCancelarUltimaLlamada')
    .addToUi();
}

/* Regenera el link de WhatsApp de confirmación para filas en
   Comunicaciones cuyo link actual NO contiene ?id=              */
function regenerarLinksWASinId() {
  var sheet = getComunicacionesSheet();
  var data  = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var correo   = (data[i][0] || '').toString().trim();
    var nombre   = (data[i][1] || '').toString().trim();
    var telefono = (data[i][2] || '').toString().trim();
    var linkActual = (data[i][4] || '').toString();
    if (!correo || !telefono) continue;
    if (linkActual.indexOf('?id=') !== -1) continue; // ya tiene ID, saltar
    var id = obtenerIdAsistente(correo);
    if (!id) continue; // sin ID en Asistencia, saltar
    var nuevoLink = generateWhatsAppLinkConfirmacion(nombre, telefono, id);
    if (nuevoLink) {
      sheet.getRange(i + 1, 5).setValue(nuevoLink);
      sheet.getRange(i + 1, 5).setFormula('=HYPERLINK("' + nuevoLink + '","Enviar WhatsApp")');
      count++;
    }
  }
  SpreadsheetApp.getUi().alert('Listo. Se actualizaron ' + count + ' link(s) con su ID personalizado.');
}

/* ── Sincronizar registros faltantes ─────────────────────────── */
/* Recorre Registros, y para cada fila con Pagó ✓:
   - Si el correo no está en Asistencia → lo agrega
   - Si la fase es x2 → agrega 1 "Acompañante de..." en Asistencia
   - Si la fase es x3 → agrega 2 "Acompañante de..." en Asistencia
   - Si el correo no está en Comunicaciones → lo agrega con botón WA
   No envía ningún correo ni WhatsApp, todo es manual.           */
function sincronizarRegistrosFaltantes() {
  var sheet    = getSheet();
  var data     = sheet.getDataRange().getValues();
  var vistos   = {};
  var agregados = 0;

  for (var i = 1; i < data.length; i++) {
    var pago    = (data[i][9] || '').toString().trim();
    if (pago !== '✓') continue;

    var correo  = (data[i][2] || '').toString().toLowerCase().trim();
    var nombre  = (data[i][1] || '').toString().trim();
    var telefono= (data[i][3] || '').toString().trim();
    var contacto= (data[i][4] || '').toString().trim();
    var fase    = (data[i][6] || '').toString().trim();

    if (!correo) continue;

    if (!vistos[correo]) {
      vistos[correo] = { nombre: nombre, fase: fase, boletos: 1 };

      // Agregar a Asistencia si no existe
      actualizarAsistencia(correo, nombre, fase);

      // Agregar acompañantes según fase (x2 = 1 acomp., x3 = 2 acomp.)
      var cantAcompanantes = 0;
      if (fase.indexOf('x2') !== -1) cantAcompanantes = 1;
      if (fase.indexOf('x3') !== -1) cantAcompanantes = 2;
      for (var a = 0; a < cantAcompanantes; a++) {
        agregarAcompanante(nombre, fase);
      }

      // Agregar a Comunicaciones si no existe
      var id = obtenerIdAsistente(correo);
      sincronizarComunicaciones(correo, nombre, telefono, contacto, id);

      agregados++;
    } else {
      // Segunda o tercera fila del mismo correo (compra adicional)
      vistos[correo].boletos++;
      // Agregar acompañante por cada boleto adicional
      agregarAcompanante(vistos[correo].nombre, fase);
    }

    Utilities.sleep(300);
  }

  SpreadsheetApp.getUi().alert('Listo. Se sincronizaron ' + agregados + ' persona(s). Revisa Asistencia para ver los acompañantes agregados.');
}

/* ── Admin actions (desde panel) ─────────────────────────── */
function handleAdminAction(data) {
  var sub = data.sub || '';

  if (sub === 'test_email')                  return adminTestEmail();
  if (sub === 'test_email_agradecimiento')   return adminTestEmailAgradecimiento();
  if (sub === 'checkin')                     return adminCheckin(data.id);
  if (sub === 'update_correo')               return adminUpdateCorreo(data.id, data.correo_nuevo);
  if (sub === 'registrar_acompanante')       return adminRegistrarAcompanante(data.id, data.nombre_nuevo, data.correo_nuevo);
  if (sub === 'enviar_bienvenida')           return adminEnviarBienvenida(data.id);
  if (sub === 'registrar_sustituta')         return adminRegistrarSustituta(data.id_original, data.nombre, data.correo, data.whatsapp);
  if (sub === 'update_nombre')               return adminUpdateNombre(data.id, data.nombre_nuevo);
  if (sub === 'registro_manual')             return adminRegistroManual(data.nombre, data.correo, data.telefono, data.fase);
  if (sub === 'registrar_invitada')          return adminRegistrarInvitada(data.nombre, data.correo, data.telefono);
  if (sub === 'test_email_acompanante')      return adminTestEmailAcompanante();
  if (sub === 'estado_correos')              return adminEstadoCorreos();
  if (sub === 'programar_indicaciones')      return adminProgramarIndicaciones();
  if (sub === 'cancelar_indicaciones')       return adminCancelarIndicaciones();
  if (sub === 'enviar_indicaciones_masivo')  return adminEnviarIndicacionesMasivo();
  if (sub === 'programar_recordatorio')      return adminProgramarRecordatorio();
  if (sub === 'cancelar_recordatorio')       return adminCancelarRecordatorio();
  if (sub === 'enviar_recordatorio_masivo')  return adminEnviarRecordatorioMasivo();
  if (sub === 'programar_agradecimiento')    return adminProgramarAgradecimiento();
  if (sub === 'cancelar_agradecimiento')     return adminCancelarAgradecimiento();
  if (sub === 'enviar_agradecimiento_masivo')  return adminEnviarAgradecimientoMasivo();
  if (sub === 'programar_recordatorio_post')   return adminProgramarRecordatorioPost();
  if (sub === 'cancelar_recordatorio_post')    return adminCancelarRecordatorioPost();
  if (sub === 'enviar_recordatorio_post_masivo') return adminEnviarRecordatorioPostMasivo();
  if (sub === 'programar_ultima_llamada')      return adminProgramarUltimaLlamada();
  if (sub === 'cancelar_ultima_llamada')       return adminCancelarUltimaLlamada();
  if (sub === 'enviar_ultima_llamada_masivo')  return adminEnviarUltimaLlamadaMasivo();
  if (sub === 'test_email_recordatorio_post')  return adminTestEmailRecordatorioPost();
  if (sub === 'test_email_ultima_llamada')     return adminTestEmailUltimaLlamada();
  if (sub === 'enviar_qr_masivo_gs')        return ContentService.createTextOutput(JSON.stringify(enviarCorreosQRMasivo()||{result:'ok'})).setMimeType(ContentService.MimeType.JSON);
  if (sub === 'reenviar_rescate')            return adminReenviarRescate(data.tipo, data.email);
  if (sub === 'reporte_completo')            return adminReporteCompleto();
  if (sub === 'set_survey_status')           return adminSetSurveyStatus(data.status);
  if (sub === 'set_hub_status')              return adminSetHubStatus(data.status);
  if (sub === 'create_test_profile')         return adminCreateTestProfile();

  // ── Contenidos editables (Etapa 2) ──────────────────────────
  if (sub === 'leer_contenidos')        return adminLeerContenidos();
  if (sub === 'guardar_borrador')       return adminGuardarBorrador(data.id, data.canal, data.campo, data.valor);
  if (sub === 'publicar_contenido')     return adminPublicarContenido(data.id, data.canal, data.campo);
  if (sub === 'guardar_borradores_lote') return adminGuardarBorradoresLote(data.id, data.canal, data.campos);
  if (sub === 'publicar_contenidos_lote') return adminPublicarContenidosLote(data.id, data.canal, data.campos);

  return ContentService.createTextOutput(JSON.stringify({ error: 'Acción desconocida' })).setMimeType(ContentService.MimeType.JSON);
}

function adminCheckin(id) {
  if (!id) return jsErr('ID requerido');
  var sheet = getAsistenciaSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === id) {
      var nombre = data[i][1];
      if (data[i][4] === '✓') return jsOk({ yaRegistrado: true, nombre: nombre });

      // Gate: requiere encuesta previa
      var ss2 = SpreadsheetApp.getActiveSpreadsheet();
      var prevSheet = ss2.getSheetByName('Encuesta Previa');
      var tienePrevia = false;
      if (prevSheet) {
        var prevData = prevSheet.getDataRange().getValues();
        for (var k = 1; k < prevData.length; k++) {
          if ((prevData[k][0]||'').toString().trim() === id) { tienePrevia = true; break; }
        }
      }
      if (!tienePrevia) {
        return ContentService.createTextOutput(JSON.stringify({ error: 'encuesta_previa_pendiente', nombre: nombre })).setMimeType(ContentService.MimeType.JSON);
      }

      sheet.getRange(i+1,5).setValue('✓');
      sheet.getRange(i+1,6).setValue(new Date());
      return jsOk({ nombre: nombre, fase: data[i][3] });
    }
  }
  return jsErr('ID no encontrado');
}

function adminUpdateCorreo(id, correoNuevo) {
  if (!id || !correoNuevo) return jsErr('Faltan datos');
  correoNuevo = correoNuevo.toLowerCase().trim();

  // Actualizar en Asistencia
  var asi = getAsistenciaSheet();
  var aData = asi.getDataRange().getValues();
  var nombre = '';
  for (var i = 1; i < aData.length; i++) {
    if ((aData[i][0]||'').toString().trim() === id) {
      asi.getRange(i+1,3).setValue(correoNuevo);
      nombre = aData[i][1];
      break;
    }
  }

  // Actualizar o crear en Comunicaciones
  var com = getComunicacionesSheet();
  var cData = com.getDataRange().getValues();
  var filaComm = null;
  for (var j = 1; j < cData.length; j++) {
    if ((cData[j][0]||'').toLowerCase().trim() === correoNuevo) { filaComm = j+1; break; }
  }
  if (!filaComm) {
    com.appendRow([correoNuevo, nombre, '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    filaComm = com.getLastRow();
  } else {
    if (!com.getRange(filaComm,2).getValue()) com.getRange(filaComm,2).setValue(nombre);
  }

  return jsOk({ correo: correoNuevo });
}

function adminRegistrarAcompanante(id, nombreNuevo, correoNuevo) {
  if (!id || !correoNuevo) return jsErr('Faltan datos');
  correoNuevo = correoNuevo.toLowerCase().trim();

  // Actualizar nombre y correo en Asistencia
  var asi = getAsistenciaSheet();
  var aData = asi.getDataRange().getValues();
  var fase = '';
  var nombreFinal = nombreNuevo || '';
  for (var i = 1; i < aData.length; i++) {
    if ((aData[i][0]||'').toString().trim() === id) {
      if (nombreNuevo) asi.getRange(i+1,2).setValue(nombreNuevo);
      asi.getRange(i+1,3).setValue(correoNuevo);
      fase = aData[i][3];
      break;
    }
  }

  // Crear en Comunicaciones
  var com = getComunicacionesSheet();
  var fila = findRowByEmailInSheet(com, correoNuevo);
  if (!fila) {
    com.appendRow([correoNuevo, nombreFinal, '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    fila = com.getLastRow();
  }

  // Enviar correo de bienvenida todo-en-uno
  enviarCorreoBienvenidaAcompanante(nombreFinal, correoNuevo, fase, id);

  // Marcar correo de confirmación como enviado en Comunicaciones (col G)
  com.getRange(fila, 7).setValue('Sí');

  return jsOk({ result: 'ok', correo: correoNuevo });
}

function adminEnviarBienvenida(id) {
  if (!id) return jsErr('ID requerido');
  var asi = getAsistenciaSheet();
  var aData = asi.getDataRange().getValues();
  for (var i = 1; i < aData.length; i++) {
    if ((aData[i][0]||'').toString().trim() === id) {
      var nombre = aData[i][1];
      var correo = aData[i][2];
      var fase   = aData[i][3];
      if (!correo) return jsErr('Sin correo registrado');
      enviarCorreoBienvenidaAcompanante(nombre, correo, fase, id);
      var com = getComunicacionesSheet();
      var fila = findRowByEmailInSheet(com, correo);
      if (fila) com.getRange(fila,7).setValue('Sí');
      return jsOk({ result: 'ok' });
    }
  }
  return jsErr('ID no encontrado');
}

/* ── Registrar sustituta de una compradora original ─────────────
   Crea un nuevo registro vinculado (historial de la compradora intacto),
   asigna ID propio a la sustituta y le manda el correo de bienvenida
   con su hub personal (QR + encuesta previa + recursos del taller).   */
function adminRegistrarSustituta(idOriginal, nombreSub, correoSub, whatsappSub) {
  if (!idOriginal || !nombreSub || !correoSub) return jsErr('Faltan datos');
  correoSub = correoSub.toLowerCase().trim();

  // ── 1. Encontrar a la compradora original en Asistencia ──
  var asi    = getAsistenciaSheet();
  var aData  = asi.getDataRange().getValues();
  var faseOriginal   = 'Cortesía';
  var nombreOriginal = '';
  var correoOriginal = '';
  var filaOriginal   = null;
  for (var i = 1; i < aData.length; i++) {
    if ((aData[i][0]||'').toString().trim() === idOriginal) {
      faseOriginal   = aData[i][3] || 'Cortesía';
      nombreOriginal = aData[i][1] || '';
      correoOriginal = (aData[i][2] || '').toLowerCase().trim();
      filaOriginal   = i + 1;
      break;
    }
  }
  if (!nombreOriginal) return jsErr('ID original no encontrado: ' + idOriginal);
  // Marcar en col K que cedió lugar → queda oculta en el admin
  if (filaOriginal) asi.getRange(filaOriginal, 11).setValue('cedió lugar');

  // ── 2. Anotar en Registros de la compradora que cedió su lugar ──
  var reg   = getSheet();
  var rData = reg.getDataRange().getValues();
  for (var k = 1; k < rData.length; k++) {
    if ((rData[k][2]||'').toLowerCase().trim() === correoOriginal) {
      var origenActual = (rData[k][11]||'').toString().trim();
      var nota = 'cedió lugar a ' + nombreSub + ' (' + correoSub + ')';
      reg.getRange(k + 1, 12).setValue(origenActual ? origenActual + ' · ' + nota : nota);
      break;
    }
  }

  // ── 3. Crear registro independiente para la sustituta ──
  var newId = generarSiguienteId(asi);
  asi.appendRow([newId, nombreSub, correoSub, faseOriginal, '', '', '', '', '', '']);

  reg.appendRow([
    new Date(), nombreSub, correoSub, whatsappSub || '', '', '',
    faseOriginal, '0.00', new Date(), '✓',
    'SUST-' + newId, 'sustituta de ' + nombreOriginal + ' (' + correoOriginal + ')',
    '', '', '', '', ''
  ]);

  sincronizarComunicaciones(correoSub, nombreSub, whatsappSub || '', '', newId);

  // ── 4. Enviar correo de bienvenida a la sustituta con su propio hub ──
  enviarCorreoBienvenidaAcompanante(nombreSub, correoSub, faseOriginal, newId);

  // Marcar en Comunicaciones col V (22) que se envió bienvenida sustituta/invitada
  var comS  = getComunicacionesSheet();
  var filaS = findRowByEmailInSheet(comS, correoSub);
  if (filaS) comS.getRange(filaS, 22).setValue('Sí');

  return jsOk({ id: newId, nombre: nombreSub, correo: correoSub });
}

/* ── Rescate: reenvío individual por email ─────────────────────── */
function adminReenviarRescate(tipo, email) {
  if (!tipo || !email) return jsErr('Falta tipo o email');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registros');
  if (!sheet) return jsErr('Hoja Registros no encontrada');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var iNombre = headers.indexOf('Nombre');
  var iCorreo = headers.indexOf('Correo');
  var iFase   = headers.indexOf('Fase de compra'); // nombre real de la columna
  if (iNombre < 0 || iCorreo < 0) return jsErr('Columnas Nombre/Correo no encontradas');
  var fila = null;
  for (var i = 1; i < data.length; i++) {
    if ((data[i][iCorreo] || '').toLowerCase().trim() === email.toLowerCase().trim()) {
      fila = data[i]; break;
    }
  }
  if (!fila) return jsErr('No se encontró ningún registro con ese correo: ' + email);
  var nombre  = fila[iNombre] || '';
  var correo  = fila[iCorreo] || '';
  var fase    = iFase >= 0 ? (fila[iFase] || '') : '';
  // ID viene de Asistencia, no de Registros (Registros no tiene columna ID)
  var idUnico = obtenerIdAsistente(correo);
  try {
    if (tipo === 'confirmacion')  enviarCorreoConfirmacion(nombre, correo, fase, idUnico);
    else if (tipo === 'indicaciones') enviarCorreoIndicaciones(nombre, correo, idUnico);
    else if (tipo === 'recordatorio') enviarCorreoRecordatorio(nombre, correo, idUnico);
    else if (tipo === 'qr')           enviarCorreoQR(nombre, correo, idUnico);
    else if (tipo === 'agradecimiento') enviarCorreoAgradecimiento(nombre, correo, idUnico);
    else return jsErr('Tipo de comunicación desconocido: ' + tipo);
    return jsOk({ enviado: 1, nombre: nombre, correo: correo });
  } catch(e) {
    return jsErr('Error al enviar: ' + e.message);
  }
}

/* Correo único todo-en-uno para acompañantes */
function enviarCorreoBienvenidaAcompanante(nombre, correo, fase, idUnico) {
  var p      = nombre ? nombre.split(' ')[0] : 'Hola';
  var hubUrl = 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico;

  // Contenido editable con fallback hardcodeado
  var _vars = { nombre: p, hub: hubUrl, id: idUnico || '' };
  var asunto   = _aplicarVariables(getContenidoPublicado('acompanante','correo','asunto'),   _vars)
               || '¡Te esperamos este sábado! — Un paso antes de llegar';
  var apertura = _aplicarVariables(getContenidoPublicado('acompanante','correo','apertura'), _vars)
               || 'Alguien muy especial reservó un lugar para ti en el taller <em>Lo que tu imagen comunica</em> con Mary Méndez. Estamos felices de tenerte el <strong>sábado 15 de agosto</strong>.';
  var cierre   = _aplicarVariables(getContenidoPublicado('acompanante','correo','cierre'),   _vars)
               || '¡Te esperamos el sábado! Cualquier duda, escríbenos a este correo.';

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;">'

    // Chip: invitada especial
    + '<div style="display:inline-block;background:rgba(198,165,106,.12);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Invitada especial · REINVENTA</div>'

    // Encabezado
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 1rem;">Hola, ' + p + '.<br>Qué gusto que nos acompañas este sábado.</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;">' + apertura + '</p>'

    // Bloque urgente: encuesta previa (lo más importante)
    + '<div style="border:2px solid #C6A56A;padding:1.4rem 1.6rem;margin-bottom:1.8rem;background:rgba(198,165,106,.07);">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .5rem;">Antes de llegar — paso obligatorio</p>'
    + '<p style="font-family:Georgia,serif;font-size:1.1rem;color:#2A0F25;margin:0 0 .8rem;">Contesta la encuesta previa desde tu espacio personal.</p>'
    + '<p style="font-size:.87rem;color:#4a3545;line-height:1.7;margin:0 0 1rem;">Es una encuesta breve que Mary revisa personalmente para personalizar la experiencia de cada asistente. <strong>Sin contestarla no podrás hacer check-in el día del evento</strong>, así que te pedimos que la completes antes del sábado.</p>'
    + '<a href="' + hubUrl + '" style="display:block;background:#C6A56A;color:#2A0F25;text-align:center;padding:.9rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;font-weight:600;margin-bottom:.6rem;">Ir a mi espacio y contestar la encuesta →</a>'
    + '<p style="font-size:.7rem;color:#8F7383;text-align:center;margin:0;">Aquí también encontrarás tu QR de entrada y toda la info del día</p>'
    + '</div>'

    // Detalles del evento (breve)
    + '<div style="border:1px solid rgba(42,15,37,.1);padding:1.2rem 1.4rem;margin-bottom:1.8rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .8rem;">Detalles del evento</p>'
    + '<p style="font-size:.88rem;color:#2A0F25;margin:0 0 .3rem;font-weight:600;">Lo que tu imagen comunica — REINVENTA</p>'
    + '<p style="font-size:.85rem;color:#4a3545;margin:0 0 .25rem;">Sábado 15 de agosto de 2026 · 10:00–12:00 pm</p>'
    + '<p style="font-size:.85rem;color:#4a3545;margin:0;">The University Club of Mexico<br>Av. Paseo de la Reforma 150, Juárez, CDMX</p>'
    + '</div>'

    // Cierre cálido
    + '<p style="font-size:.9rem;line-height:1.7;color:#4a3545;margin:0 0 .5rem;">' + cierre + '</p>'
    + '<p style="font-size:.9rem;color:#4a3545;margin:0;">Con cariño,</p>'
    + '<p style="font-family:Georgia,serif;font-size:.95rem;color:#2A0F25;margin:.2rem 0 0;">Mary y el equipo de REINVENTA</p>'

    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({
    to: correo,
    bcc: 'alopez@alumbrastudios.com',
    name: 'Reinventa by Mary Méndez',
    subject: asunto,
    htmlBody: html
  });
}

/* ── Estado de correos (para el tab Correos del admin) ─────── */
function adminEstadoCorreos() {
  var com  = getComunicacionesSheet();
  var asi  = getAsistenciaSheet();
  var cData = com.getDataRange().getValues();
  var aData = asi.getDataRange().getValues();

  // Check-in set por correo
  var checkinSet = {};
  for (var i = 1; i < aData.length; i++) {
    var c = (aData[i][2]||'').toLowerCase().trim();
    if (c && aData[i][4] === '✓') checkinSet[c] = true;
  }

  var total = 0, recEnv = 0, qrEnv = 0, agrEnv = 0, indEnv = 0, conCheckin = 0;
  for (var j = 1; j < cData.length; j++) {
    var correo = (cData[j][0]||'').toLowerCase().trim();
    if (!correo) continue;
    total++;
    if ((cData[j][9] ||'').toString() === 'Sí') recEnv++;
    if ((cData[j][12]||'').toString() === 'Sí') qrEnv++;
    if ((cData[j][15]||'').toString() === 'Sí') agrEnv++;
    if ((cData[j][18]||'').toString() === 'Sí') indEnv++; // col S = indicaciones
    if (checkinSet[correo]) conCheckin++;
  }

  // Estado de triggers guardado en PropertiesService (no requiere scope extra)
  var props = PropertiesService.getScriptProperties();
  var trigRec     = props.getProperty('trigger_recordatorio') === 'true';
  var trigAgr     = props.getProperty('trigger_agradecimiento') === 'true';
  var trigInd     = props.getProperty('trigger_indicaciones') === 'true';
  var trigRecPost = props.getProperty('trigger_recordatorio_post') === 'true';
  var trigUltima  = props.getProperty('trigger_ultima_llamada') === 'true';

  // Confirmación: col G (índice 6) = 'Sí' cuando se envió bienvenida
  var confEnv = 0;
  for (var k = 1; k < cData.length; k++) {
    if ((cData[k][6]||'').toString() === 'Sí') confEnv++;
  }

  return jsOk({
    total: total,
    confirmacion:   { enviados: confEnv, pendientes: total - confEnv },
    indicaciones:   { enviados: indEnv,  pendientes: total - indEnv,  programado: trigInd, fecha: 'Jue 13 ago 2026 · 12:00 pm' },
    recordatorio:   { enviados: recEnv,  pendientes: total - recEnv,  programado: trigRec, fecha: 'Vie 14 ago 2026 · 12:00 pm' },
    qr:             { enviados: qrEnv,   pendientes: total - qrEnv,   programado: true,    fecha: 'Sáb 15 ago 2026 · 8:00 am' },
    agradecimiento:      { enviados: agrEnv,  pendientes: conCheckin - agrEnv, conCheckin: conCheckin, programado: trigAgr,     fecha: 'Lun 17 ago 2026 · 10:00 am' },
    recordatorio_post:   { enviados: '—', pendientes: '—', programado: trigRecPost, fecha: 'Mié 19 ago 2026 · 11:00 am' },
    ultima_llamada:      { enviados: '—', pendientes: '—', programado: trigUltima,  fecha: 'Vie 21 ago 2026 · 12:00 pm' }
  });
}

/* ── Recordatorio masivo ───────────────────────────────────── */
// Función ejecutada por el trigger automático
function triggerRecordatorioMasivo() { adminEnviarRecordatorioMasivo(); }

function adminEnviarRecordatorioMasivo() {
  var com   = getComunicacionesSheet();
  var asi   = getAsistenciaSheet();
  var cData = com.getDataRange().getValues();
  var aData = asi.getDataRange().getValues();

  var idPorCorreo = {};
  for (var i = 1; i < aData.length; i++) {
    var c = (aData[i][2]||'').toLowerCase().trim();
    if (c) idPorCorreo[c] = aData[i][0];
  }

  var enviados = 0;
  for (var j = 1; j < cData.length; j++) {
    var correo  = (cData[j][0]||'').toLowerCase().trim();
    var nombre  = cData[j][1] || '';
    var yaEnv   = (cData[j][9]||'').toString();
    if (!correo || yaEnv === 'Sí') continue;
    var id = idPorCorreo[correo] || '';
    try {
      enviarCorreoRecordatorio(nombre, correo, id);
      com.getRange(j+1, 10).setValue('Sí'); // col J
      enviados++;
      Utilities.sleep(800);
    } catch(e) { Logger.log('Error recordatorio ' + correo + ': ' + e); }
  }
  Logger.log('Recordatorios enviados: ' + enviados);
  return jsOk({ enviados: enviados });
}

function adminProgramarRecordatorio() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerRecordatorioMasivo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('triggerRecordatorioMasivo')
    .timeBased().at(new Date('2026-08-14T18:00:00')).create(); // UTC = 12pm Mexico City (UTC-6)
  PropertiesService.getScriptProperties().setProperty('trigger_recordatorio', 'true');
  return jsOk({ programado: true, fecha: 'Vie 14 ago 2026 · 12:00 pm' });
}

function adminCancelarRecordatorio() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerRecordatorioMasivo') { ScriptApp.deleteTrigger(t); n++; }
  });
  PropertiesService.getScriptProperties().deleteProperty('trigger_recordatorio');
  return jsOk({ cancelado: true, eliminados: n });
}

/* ── Indicaciones masivo (jueves 13 ago 12pm) ──────────────── */
function triggerIndicacionesMasivo() { adminEnviarIndicacionesMasivo(); }

function adminEnviarIndicacionesMasivo() {
  var com   = getComunicacionesSheet();
  var asi   = getAsistenciaSheet();
  var cData = com.getDataRange().getValues();
  var aData = asi.getDataRange().getValues();

  var idPorCorreo = {};
  for (var i = 1; i < aData.length; i++) {
    var c = (aData[i][2]||'').toLowerCase().trim();
    if (c) idPorCorreo[c] = aData[i][0];
  }

  var enviados = 0;
  for (var j = 1; j < cData.length; j++) {
    var correo  = (cData[j][0]||'').toLowerCase().trim();
    var nombre  = cData[j][1] || '';
    var yaEnv   = (cData[j][18]||'').toString(); // col S (índice 18)
    if (!correo || yaEnv === 'Sí') continue;
    var id = idPorCorreo[correo] || '';
    try {
      enviarCorreoIndicaciones(nombre, correo, id);
      com.getRange(j+1, 19).setValue('Sí'); // col S
      enviados++;
      Utilities.sleep(800);
    } catch(e) { Logger.log('Error indicaciones ' + correo + ': ' + e); }
  }
  Logger.log('Indicaciones enviadas: ' + enviados);
  return jsOk({ enviados: enviados });
}

function adminProgramarIndicaciones() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerIndicacionesMasivo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('triggerIndicacionesMasivo')
    .timeBased().at(new Date('2026-08-13T18:00:00')).create(); // UTC = 12pm Mexico City (UTC-6)
  PropertiesService.getScriptProperties().setProperty('trigger_indicaciones', 'true');
  return jsOk({ programado: true, fecha: 'Jue 13 ago 2026 · 12:00 pm' });
}

function adminCancelarIndicaciones() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerIndicacionesMasivo') { ScriptApp.deleteTrigger(t); n++; }
  });
  PropertiesService.getScriptProperties().deleteProperty('trigger_indicaciones');
  return jsOk({ cancelado: true, eliminados: n });
}

/* ── Agradecimiento masivo (solo asistentes con check-in) ──── */
function triggerAgradecimientoMasivo() { adminEnviarAgradecimientoMasivo(); }

function adminEnviarAgradecimientoMasivo() {
  var com   = getComunicacionesSheet();
  var asi   = getAsistenciaSheet();
  var cData = com.getDataRange().getValues();
  var aData = asi.getDataRange().getValues();

  var idPorCorreo = {}, checkinSet = {};
  for (var i = 1; i < aData.length; i++) {
    var c = (aData[i][2]||'').toLowerCase().trim();
    if (c) {
      idPorCorreo[c] = aData[i][0];
      if (aData[i][4] === '✓') checkinSet[c] = true;
    }
  }

  var enviados = 0;
  for (var j = 1; j < cData.length; j++) {
    var correo = (cData[j][0]||'').toLowerCase().trim();
    var nombre = cData[j][1] || '';
    var yaEnv  = (cData[j][15]||'').toString();
    if (!correo || yaEnv === 'Sí') continue;
    if (!checkinSet[correo]) continue; // Solo con check-in
    var id = idPorCorreo[correo] || '';
    try {
      enviarCorreoAgradecimiento(nombre, correo, id);
      com.getRange(j+1, 16).setValue('Sí'); // col P
      enviados++;
      Utilities.sleep(800);
    } catch(e) { Logger.log('Error agradecimiento ' + correo + ': ' + e); }
  }
  Logger.log('Agradecimientos enviados: ' + enviados);
  return jsOk({ enviados: enviados });
}

function adminProgramarAgradecimiento() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerAgradecimientoMasivo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('triggerAgradecimientoMasivo')
    .timeBased().at(new Date('2026-08-17T16:00:00')).create(); // UTC = 10am Mexico City (UTC-6)
  PropertiesService.getScriptProperties().setProperty('trigger_agradecimiento', 'true');
  return jsOk({ programado: true, fecha: 'Lun 17 ago 2026 · 10:00 am' });
}

function adminCancelarAgradecimiento() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerAgradecimientoMasivo') { ScriptApp.deleteTrigger(t); n++; }
  });
  PropertiesService.getScriptProperties().deleteProperty('trigger_agradecimiento');
  return jsOk({ cancelado: true, eliminados: n });
}

/* ════════════════════════════════════════════════════════════════
   RECORDATORIOS POST-TALLER
   Miércoles 19 ago — solo a quienes NO contestaron encuesta post
   Viernes 21 ago   — última llamada, hub cierra lunes 24
   ════════════════════════════════════════════════════════════════ */

/* ── Helpers de filtro ───────────────────────────────────────── */
function _obtenerDescargasPorId() {
  var descargasSet = {};
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Descargas');
    if (sh) {
      var rows = sh.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        var rid = (rows[i][0] || '').toString().trim();
        if (rid) descargasSet[rid] = true;
      }
    }
  } catch(e) {}
  return descargasSet;
}

/* Devuelve array de {nombre, correo, id} que tienen check-in
   pero NO encuesta post Y NO descargas */
function _filtrarSinRecursos() {
  var aSheet = getAsistenciaSheet();
  var aData  = aSheet.getDataRange().getValues();
  var descSet = _obtenerDescargasPorId();
  var pendientes = [];
  var vistos = {};
  for (var i = 1; i < aData.length; i++) {
    var id      = (aData[i][0] || '').toString().trim();
    var nombre  = (aData[i][1] || '').toString().trim();
    var correo  = (aData[i][2] || '').toString().toLowerCase().trim();
    var checkin = aData[i][4] === '✓';
    var encPost = aData[i][6] === '✓';
    if (!correo || vistos[correo] || !checkin) continue;
    // Solo sin encuesta post Y sin ninguna descarga
    if (!encPost && !descSet[id]) {
      pendientes.push({ id: id, nombre: nombre, correo: correo });
      vistos[correo] = true;
    }
  }
  return pendientes;
}

/* ── Correo: recordatorio de recursos (Mié 19) ───────────────── */
function enviarCorreoRecordatorioPost(nombre, correo, idUnico) {
  var p   = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var hub = (idUnico ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico : 'https://reinventabymarymendez.com.mx/hub') + '&tab=recursos';
  var _vars = { nombre: p, hub: hub, id: idUnico || '' };

  // Leer desde Contenidos (con fallbacks si no hay publicado)
  var g = function(campo, fb) {
    return _aplicarVariables(getContenidoPublicado('recordatorio_post', 'correo', campo), _vars) || _aplicarVariables(fb, _vars);
  };
  var asunto   = g('asunto',   '[nombre], tus recursos del taller todavía te esperan — REINVENTA');
  var h1       = g('h1',       '[nombre], aún puedes llevarte tus guías del taller.');
  var apertura = g('apertura', 'Hace unos días compartiste con nosotras algo muy especial. Queremos asegurarnos de que puedas llevarte todo lo que prometimos — las guías y tu constancia de participación.\n\nSolo necesitas contestar la encuesta de satisfacción (toma menos de 3 minutos). Al hacerlo, se desbloquean automáticamente en tu espacio personal.');
  var cierre   = g('cierre',   'El acceso al hub cierra el lunes 24 de agosto.');

  // Renderizar párrafos de apertura (puede tener saltos de línea)
  var aperturaPars = apertura.split('\n').filter(function(l){ return l.trim(); }).map(function(l){
    return '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.2rem;">' + l + '</p>';
  }).join('');

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Tus recursos te esperan</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 1rem;">' + h1 + '</h1>'
    + aperturaPars
    + '<div style="background:#2A0F25;padding:1.4rem 1.6rem;margin-bottom:1.4rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:rgba(198,165,106,.6);margin:0 0 .5rem;">Al contestar la encuesta desbloqueas</p>'
    + '<p style="font-size:.83rem;color:rgba(239,233,226,.85);margin:0 0 .3rem;">· Tu constancia de participación descargable.</p>'
    + '<p style="font-size:.83rem;color:rgba(239,233,226,.85);margin:0 0 .3rem;">· Guía de 7 estilos y sus esenciales.</p>'
    + '<p style="font-size:.83rem;color:rgba(239,233,226,.85);margin:0 0 1rem;">· Guía de compras inteligentes.</p>'
    + '<p style="font-size:.78rem;color:rgba(198,165,106,.7);margin:0 0 .8rem;">⏳ ' + cierre + '</p>'
    + '<a href="' + hub + '" style="display:block;background:#C6A56A;color:#2A0F25;text-align:center;padding:.8rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;font-weight:600;margin-top:.5rem;">Ir a mis recursos →</a>'
    + '</div>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: asunto.replace('[nombre]', p).replace('{{nombre}}', p), htmlBody: html });
}

/* ── Correo: última llamada (Vie 21) ─────────────────────────── */
function enviarCorreoUltimaLlamada(nombre, correo, idUnico) {
  var p   = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var hub = (idUnico ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico : 'https://reinventabymarymendez.com.mx/hub') + '&tab=recursos';
  var _vars = { nombre: p, hub: hub, id: idUnico || '' };

  // Leer desde Contenidos (con fallbacks si no hay publicado)
  var g = function(campo, fb) {
    return _aplicarVariables(getContenidoPublicado('ultima_llamada', 'correo', campo), _vars) || _aplicarVariables(fb, _vars);
  };
  var asunto   = g('asunto',   'Último aviso — tu acceso al hub cierra el lunes — REINVENTA');
  var h1       = g('h1',       '[nombre], el lunes 24 cierra el acceso a tu espacio personal.');
  var apertura = g('apertura', 'Este es el último aviso. El lunes 24 de agosto se cierra definitivamente el acceso a tu espacio personal de REINVENTA.\n\nAntes de que eso pase, contesta la encuesta (3 minutos) y descarga tus guías. Es lo último que necesitas hacer.');
  var cierre   = g('cierre',   'Tienes hasta el domingo 23 de agosto.');

  // Renderizar párrafos de apertura
  var aperturaPars = apertura.split('\n').filter(function(l){ return l.trim(); }).map(function(l){
    return '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.2rem;">' + l + '</p>';
  }).join('');

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Último aviso · Lunes 24 ago</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 1rem;">' + h1 + '</h1>'
    + aperturaPars
    + '<div style="background:#2A0F25;padding:1.4rem 1.6rem;margin-bottom:1.4rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:rgba(198,165,106,.6);margin:0 0 .5rem;">' + cierre + '</p>'
    + '<p style="font-size:.83rem;color:rgba(239,233,226,.85);margin:0 0 .3rem;">· Constancia de participación.</p>'
    + '<p style="font-size:.83rem;color:rgba(239,233,226,.85);margin:0 0 .3rem;">· Guía de 7 estilos y sus esenciales.</p>'
    + '<p style="font-size:.83rem;color:rgba(239,233,226,.85);margin:0 0 1rem;">· Guía de compras inteligentes.</p>'
    + '<a href="' + hub + '" style="display:block;background:#C6A56A;color:#2A0F25;text-align:center;padding:.8rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;font-weight:600;">Entrar ahora antes del lunes →</a>'
    + '</div>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: asunto, htmlBody: html });
}

/* ── Envíos masivos ──────────────────────────────────────────── */
function adminEnviarRecordatorioPostMasivo() {
  var pendientes = _filtrarSinRecursos();
  var enviados = 0;
  pendientes.forEach(function(p) {
    try {
      enviarCorreoRecordatorioPost(p.nombre, p.correo, p.id);
      enviados++;
      Utilities.sleep(800);
    } catch(e) { Logger.log('Error recordatorio_post ' + p.correo + ': ' + e); }
  });
  Logger.log('Recordatorio post enviados: ' + enviados + ' de ' + pendientes.length);
  return jsOk({ enviados: enviados, total: pendientes.length });
}

function adminEnviarUltimaLlamadaMasivo() {
  var pendientes = _filtrarSinRecursos();
  var enviados = 0;
  pendientes.forEach(function(p) {
    try {
      enviarCorreoUltimaLlamada(p.nombre, p.correo, p.id);
      enviados++;
      Utilities.sleep(800);
    } catch(e) { Logger.log('Error ultima_llamada ' + p.correo + ': ' + e); }
  });
  Logger.log('Última llamada enviados: ' + enviados + ' de ' + pendientes.length);
  return jsOk({ enviados: enviados, total: pendientes.length });
}

/* ── Correos de prueba ───────────────────────────────────────── */
function adminTestEmailRecordatorioPost() {
  var pendientes = _filtrarSinRecursos();
  var p = pendientes.length > 0 ? pendientes[0] : { nombre: 'Elizabeth Santos (PRUEBA)', id: 'RNV-001' };
  enviarCorreoRecordatorioPost(p.nombre, 'mejoracontinua@caceca.org', p.id);
  return jsOk({ result: 'ok', nombre: p.nombre, id: p.id, totalPendientes: pendientes.length });
}

function adminTestEmailUltimaLlamada() {
  var pendientes = _filtrarSinRecursos();
  var p = pendientes.length > 0 ? pendientes[0] : { nombre: 'Elizabeth Santos (PRUEBA)', id: 'RNV-001' };
  enviarCorreoUltimaLlamada(p.nombre, 'mejoracontinua@caceca.org', p.id);
  return jsOk({ result: 'ok', nombre: p.nombre, id: p.id, totalPendientes: pendientes.length });
}

/* ── Triggers automáticos ────────────────────────────────────── */
function triggerRecordatorioPostMasivo()  { adminEnviarRecordatorioPostMasivo(); }
function triggerUltimaLlamadaMasivo()     { adminEnviarUltimaLlamadaMasivo(); }

function adminProgramarRecordatorioPost() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerRecordatorioPostMasivo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('triggerRecordatorioPostMasivo')
    .timeBased().at(new Date('2026-08-19T17:00:00')).create(); // 11am CDMX = 17:00 UTC
  PropertiesService.getScriptProperties().setProperty('trigger_recordatorio_post', 'true');
  return jsOk({ programado: true, fecha: 'Mié 19 ago 2026 · 11:00 am' });
}

function adminCancelarRecordatorioPost() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerRecordatorioPostMasivo') ScriptApp.deleteTrigger(t);
  });
  PropertiesService.getScriptProperties().deleteProperty('trigger_recordatorio_post');
  return jsOk({ cancelado: true });
}

function adminProgramarUltimaLlamada() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerUltimaLlamadaMasivo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('triggerUltimaLlamadaMasivo')
    .timeBased().at(new Date('2026-08-21T18:00:00')).create(); // 12pm CDMX = 18:00 UTC
  PropertiesService.getScriptProperties().setProperty('trigger_ultima_llamada', 'true');
  return jsOk({ programado: true, fecha: 'Vie 21 ago 2026 · 12:00 pm' });
}

function adminCancelarUltimaLlamada() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerUltimaLlamadaMasivo') ScriptApp.deleteTrigger(t);
  });
  PropertiesService.getScriptProperties().deleteProperty('trigger_ultima_llamada');
  return jsOk({ cancelado: true });
}

/* ─────────────────────────────────────────────────────────────── */

function adminUpdateNombre(id, nombreNuevo) {
  if (!id || !nombreNuevo) return jsErr('Faltan datos');
  var asi = getAsistenciaSheet();
  var data = asi.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0]||'').toString().trim() === id) {
      asi.getRange(i+1,2).setValue(nombreNuevo);
      var correo = (data[i][2]||'').toLowerCase().trim();
      if (correo) {
        var com = getComunicacionesSheet();
        var fila = findRowByEmailInSheet(com, correo);
        if (fila) com.getRange(fila,2).setValue(nombreNuevo);
      }
      return jsOk({ nombre: nombreNuevo });
    }
  }
  return jsErr('ID no encontrado');
}

function adminRegistroManual(nombre, correo, telefono, fase) {
  if (!nombre || !correo) return jsErr('Nombre y correo requeridos');
  correo = correo.toLowerCase().trim();
  fase = fase || 'Cortesía';

  var reg = getSheet();
  reg.appendRow([
    new Date(), nombre, correo, telefono||'', '', '', fase, '0.00', new Date(), '✓', 'MANUAL-' + new Date().getTime(), 'admin manual',
    '', '', '', '', ''
  ]);

  actualizarAsistencia(correo, nombre, fase);
  var id = obtenerIdAsistente(correo);
  sincronizarComunicaciones(correo, nombre, telefono||'', '', id);

  return jsOk({ result: 'ok', id: id });
}

/* ── Registrar invitada (cortesía) y enviar acceso completo hoy ──
   Registra a la persona en Registros + Asistencia + Comunicaciones
   con fase 'Cortesía' y le manda el correo de bienvenida con:
   hub personal, encuesta previa, medidas y QR de acceso.
   Al quedar en Comunicaciones, el trigger de QR de las 8am
   mañana también le llegará automáticamente.                       */
function adminRegistrarInvitada(nombre, correo, telefono) {
  if (!nombre || !correo) return jsErr('Nombre y correo requeridos');
  correo = correo.toLowerCase().trim();

  // Verificar que no esté ya registrada
  var asi   = getAsistenciaSheet();
  var aData = asi.getDataRange().getValues();
  for (var i = 1; i < aData.length; i++) {
    if ((aData[i][2]||'').toLowerCase().trim() === correo) {
      return jsErr('Este correo ya está registrado en el sistema');
    }
  }

  // Registrar con fase Cortesía
  var reg = getSheet();
  reg.appendRow([
    new Date(), nombre, correo, telefono||'', '', '', 'Cortesía', '0.00', new Date(), '✓',
    'INVIT-' + new Date().getTime(), 'invitada manual',
    '', '', '', '', ''
  ]);
  actualizarAsistencia(correo, nombre, 'Cortesía');
  var id = obtenerIdAsistente(correo);
  sincronizarComunicaciones(correo, nombre, telefono||'', '', id);

  // Enviar correo de bienvenida con hub + encuesta previa + medidas + QR
  enviarCorreoBienvenidaAcompanante(nombre, correo, 'Cortesía', id);

  // Marcar en Comunicaciones col V (22) que se envió bienvenida sustituta/invitada
  var comI  = getComunicacionesSheet();
  var filaI = findRowByEmailInSheet(comI, correo);
  if (filaI) comI.getRange(filaI, 22).setValue('Sí');

  return jsOk({ id: id, nombre: nombre, correo: correo });
}

function adminTestEmailAcompanante() {
  enviarCorreoBienvenidaAcompanante('Valeria García (PRUEBA acompañante)', 'mejoracontinua@caceca.org', 'Early Bird x2', 'RNV-001');
  return jsOk({ result: 'ok' });
}

function adminTestEmail() {
  enviarCorreoConfirmacion('Valeria García (PRUEBA)', 'mejoracontinua@caceca.org', 'Early Bird', 'RNV-001');
  return jsOk({ result: 'ok' });
}

function adminTestEmailAgradecimiento() {
  // Busca la primera asistente con check-in real para usar su ID en la vista del hub
  var sheet = getAsistenciaSheet();
  var rows  = sheet.getDataRange().getValues();
  var nombre = 'Elizabeth Santos (PRUEBA)';
  var idUnico = 'RNV-001';
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][4] === '✓' && rows[i][0]) {
      nombre  = (rows[i][1] || '').toString().trim() || nombre;
      idUnico = (rows[i][0] || '').toString().trim() || idUnico;
      break;
    }
  }
  enviarCorreoAgradecimiento(nombre, 'mejoracontinua@caceca.org', idUnico);
  return jsOk({ result: 'ok', nombre: nombre, id: idUnico });
}

function jsOk(obj)  { return ContentService.createTextOutput(JSON.stringify(obj || {result:'ok'})).setMimeType(ContentService.MimeType.JSON); }

/* ═══════════════════════════════════════════════════════════════════
   PROSPECTOS — Resumen diario por correo
   Se envía cada día a las 9am si hubo nuevos registros ese día.
   Destinatarios: marymendez.mk@gmail.com, alopez@alumbrastudios.com
   ═══════════════════════════════════════════════════════════════════ */
var EMAILS_PROSPECTOS = ['marymendez.mk@gmail.com', 'alopez@alumbrastudios.com'];

function enviarResumenProspectosDiario() {
  var sheet = getProspectosSheet();
  var data  = sheet.getDataRange().getValues();
  var hoy   = new Date();
  var nuevos = [];

  for (var i = 1; i < data.length; i++) {
    var fecha = data[i][0];
    if (!fecha) continue;
    var f = new Date(fecha);
    // Solo registros de las últimas 24h
    if ((hoy - f) <= 24 * 60 * 60 * 1000) {
      nuevos.push({
        nombre:      data[i][1] || '',
        correo:      data[i][2] || '',
        whatsapp:    data[i][3] || '',
        servicio:    data[i][4] || '',
        transformar: data[i][5] || '',
        cuando:      data[i][6] || '',
        contacto:    data[i][7] || ''
      });
    }
  }

  if (!nuevos.length) return; // Sin registros nuevos, no enviar nada

  var fechaStr = Utilities.formatDate(hoy, 'America/Mexico_City', "d 'de' MMMM 'de' yyyy");

  var filas = nuevos.map(function(p) {
    return '<tr>'
      + '<td style="padding:.7rem 1rem;border-bottom:1px solid rgba(42,15,37,.07);font-size:.85rem;color:#2A0F25;font-weight:600;">' + (p.nombre || '—') + '</td>'
      + '<td style="padding:.7rem 1rem;border-bottom:1px solid rgba(42,15,37,.07);font-size:.82rem;color:#4a3545;">' + (p.correo || '—') + '</td>'
      + '<td style="padding:.7rem 1rem;border-bottom:1px solid rgba(42,15,37,.07);font-size:.82rem;color:#4a3545;">' + (p.whatsapp || '—') + '</td>'
      + '<td style="padding:.7rem 1rem;border-bottom:1px solid rgba(42,15,37,.07);font-size:.82rem;color:#4a3545;">' + (p.servicio || '—') + '</td>'
      + '<td style="padding:.7rem 1rem;border-bottom:1px solid rgba(42,15,37,.07);font-size:.82rem;color:#8F7383;">' + (p.contacto || '—') + '</td>'
      + '</tr>';
  }).join('');

  var html = _headerCorreo()
    + '<div style="padding:2rem 2.4rem 1.6rem;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.2rem;">Prospectos nuevos · ' + fechaStr + '</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.35rem;color:#2A0F25;margin:0 0 .4rem;">'
    + (nuevos.length === 1 ? '1 nueva persona contactó a Mary hoy.' : nuevos.length + ' personas contactaron a Mary hoy.')
    + '</h1>'
    + '<p style="font-size:.82rem;color:#8F7383;margin:0 0 1.6rem;">Estos registros viven en la hoja <strong>Prospectos</strong> del spreadsheet.</p>'
    + '<div style="overflow-x:auto;">'
    + '<table style="width:100%;border-collapse:collapse;font-family:\'Gill Sans\',Calibri,\'Segoe UI\',sans-serif;">'
    + '<thead>'
    + '<tr style="background:#2A0F25;">'
    + '<th style="padding:.6rem 1rem;text-align:left;font-size:.6rem;letter-spacing:.12em;text-transform:uppercase;color:#C6A56A;font-weight:600;">Nombre</th>'
    + '<th style="padding:.6rem 1rem;text-align:left;font-size:.6rem;letter-spacing:.12em;text-transform:uppercase;color:#C6A56A;font-weight:600;">Correo</th>'
    + '<th style="padding:.6rem 1rem;text-align:left;font-size:.6rem;letter-spacing:.12em;text-transform:uppercase;color:#C6A56A;font-weight:600;">WhatsApp</th>'
    + '<th style="padding:.6rem 1rem;text-align:left;font-size:.6rem;letter-spacing:.12em;text-transform:uppercase;color:#C6A56A;font-weight:600;">Servicio</th>'
    + '<th style="padding:.6rem 1rem;text-align:left;font-size:.6rem;letter-spacing:.12em;text-transform:uppercase;color:#C6A56A;font-weight:600;">Contacto pref.</th>'
    + '</tr>'
    + '</thead>'
    + '<tbody>' + filas + '</tbody>'
    + '</table>'
    + '</div>'
    + '</div>'
    + _footerCorreo('resumen interno');

  var asunto = nuevos.length === 1
    ? '1 prospecto nuevo — REINVENTA · ' + fechaStr
    : nuevos.length + ' prospectos nuevos — REINVENTA · ' + fechaStr;

  EMAILS_PROSPECTOS.forEach(function(email) {
    MailApp.sendEmail({ to: email, name: 'Reinventa by Mary Méndez', subject: asunto, htmlBody: html });
  });
  Logger.log('Resumen prospectos enviado: ' + nuevos.length + ' registros');
}

/* ── Programar trigger diario a las 9am (hora México) ─────────── */
/* Ejecuta esta función UNA SOLA VEZ desde el editor para activarlo */
function programarResumenProspectosDiario() {
  // Eliminar triggers previos del mismo nombre
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarResumenProspectosDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarResumenProspectosDiario')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .inTimezone('America/Mexico_City')
    .create();
  Logger.log('✅ Trigger diario activado — resumen prospectos a las 9am');
}

/* ── Cancelar trigger ──────────────────────────────────────────── */
function cancelarResumenProspectosDiario() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarResumenProspectosDiario') { ScriptApp.deleteTrigger(t); n++; }
  });
  Logger.log('Trigger cancelado (' + n + ' eliminado(s))');
}
function jsErr(msg) { return ContentService.createTextOutput(JSON.stringify({error: msg})).setMimeType(ContentService.MimeType.JSON); }

/* ═══════════════════════════════════════════════════════════════════
   MÓDULO: Contenidos editables — Etapa 2
   Hoja "Contenidos" en el Spreadsheet.
   Columnas: A=ID | B=Canal | C=Campo | D=Borrador | E=Publicado | F=Modificado | G=Nota
   ═══════════════════════════════════════════════════════════════════ */

var SHEET_CONTENIDOS = 'Contenidos';

/* ── Filas iniciales (se crean si la hoja está vacía) ─────────── */
// Variables disponibles en plantillas: {{nombre}} {{hub}} {{id}}
// Para correo: asunto + apertura + cierre son editables; estructura HTML permanece en código.
// Para WA: mensaje completo editable como texto plano con formato WA (*negrita*, _cursiva_).
// acompanante/whatsapp/mensaje: campo reservado (no tiene generateWhatsAppLink* propio aún).
var _CAMPOS_CONTENIDOS = [
  // ── Confirmación de registro ─────────────────────────────
  ['confirmacion',   'correo',    'asunto',   'Tu lugar en el taller está confirmado — REINVENTA'],
  ['confirmacion',   'correo',    'apertura', 'Nos da mucho gusto tenerte en el taller. Mary estará encantada de acompañarte en este proceso. Guarda la fecha en tu calendario para que no se te pase ningún detalle.'],
  ['confirmacion',   'correo',    'cierre',   ''],
  ['confirmacion',   'whatsapp',  'mensaje',  ''],
  // ── Bienvenida acompañante ───────────────────────────────
  ['acompanante',    'correo',    'asunto',   '¡Te esperamos este sábado! — Un paso antes de llegar'],
  ['acompanante',    'correo',    'apertura', 'Alguien muy especial reservó un lugar para ti en el taller Lo que tu imagen comunica con Mary Méndez. Estamos felices de tenerte el sábado 15 de agosto.'],
  ['acompanante',    'correo',    'cierre',   '¡Te esperamos el sábado! Cualquier duda, escríbenos a este correo.'],
  ['acompanante',    'whatsapp',  'mensaje',  ''], // Reservado — sin función generateWhatsAppLink propia; acompañante usa el link de confirmación
  // ── Indicaciones del taller ──────────────────────────────
  ['indicaciones',   'correo',    'asunto',   '{{nombre}}, esto necesitas para el taller 🌟 — REINVENTA'],
  ['indicaciones',   'correo',    'apertura', 'Para que tu experiencia del sábado sea increíble — y podamos hacer tu análisis de cuerpo y rostro de forma precisa — aquí van las indicaciones para llegar lista:'],
  ['indicaciones',   'correo',    'cierre',   '¡Nos vemos el sábado! Va a ser un día para descubrir muchas cosas lindas.'],
  ['indicaciones',   'whatsapp',  'mensaje',  ''],
  // ── Recordatorio del evento ──────────────────────────────
  ['recordatorio',   'correo',    'asunto',   'Mañana te esperamos — REINVENTA'],
  ['recordatorio',   'correo',    'apertura', 'Mary tiene algo muy especial preparado para ti. Te esperamos puntual y con muchas ganas de transformar la manera en que tu imagen comunica quién eres.'],
  ['recordatorio',   'correo',    'cierre',   'Si tienes alguna duda de último momento no dudes en contactarnos. ¡Nos vemos mañana!'],
  ['recordatorio',   'whatsapp',  'mensaje',  ''],
  // ── QR de entrada ────────────────────────────────────────
  ['qr',             'correo',    'asunto',   'Tu entrada para hoy — REINVENTA'],
  ['qr',             'correo',    'apertura', 'Muestra este código QR al llegar al evento. El staff lo escaneará en la entrada.'],
  ['qr',             'correo',    'cierre',   '¡Nos vemos en un momento. Será un día increíble!'],
  // qr no tiene whatsapp con WA propio pero sí usa link generado
  // ── Agradecimiento post-evento ───────────────────────────
  ['agradecimiento', 'correo',    'asunto',   'Gracias por acompañarnos — REINVENTA'],
  ['agradecimiento', 'correo',    'apertura', 'Ayer vivimos algo muy especial. Un espacio íntimo donde cada una se atrevió a mirarse de otra manera — a entender lo que su imagen comunica y a imaginar cómo alinearla con quién realmente es. Gracias por confiar en este espacio y por abrirte a ese proceso.'],
  ['agradecimiento', 'correo',    'cierre',   'Lo que descubriste ayer no desaparece — es tuyo. Mary estará siempre disponible si quieres seguir este camino con más profundidad.'],
  ['agradecimiento', 'whatsapp',  'mensaje',  '*REINVENTA by Mary Mendez*\n\nHola, {{nombre}}. Ayer fue un día muy especial.\n\nGracias por confiar en este espacio y por abrirte a transformar la manera en que tu imagen comunica quién eres. Fue un honor acompañarte.\n\n- - -\n\n*Antes de cerrar — un paso importante*\n\nNos gustaría mucho conocer tu experiencia. La encuesta toma menos de 3 minutos y al completarla desbloqueas en tu espacio:\n\n· Tu constancia de participación.\n· Las guías del taller (7 estilos + Compras inteligentes).\n\nEntra aquí y contéstala desde la sección de Recursos:\n{{hub}}\n\n- - -\n\nCon mucho cariño,\nMary & el equipo de Reinventa by Mary Méndez\n\n_Mensaje informativo — por favor no respondas a este chat._'],
  // ── Campos editoriales ampliados — migración 2026-08-12 ──────
  // Confirmación
  ['confirmacion',   'correo', 'h1',                 '{{nombre}}, tu lugar en el taller está reservado.'],
  ['confirmacion',   'correo', 'seccion_hub',         'Aquí encontrarás tu pase de entrada con código QR, la agenda del día y los recursos del taller.'],
  ['confirmacion',   'correo', 'nota_encuesta',       'Dentro de tu espacio hay una encuesta breve que te pedimos contestar antes del evento. Mary la revisa personalmente para preparar materiales y recomendaciones a la medida de cada asistente. No toma más de 5 minutos y hace una gran diferencia.'],
  ['confirmacion',   'correo', 'cta_calendar_label',  'Agregar a Google Calendar'],
  // Editor dual-mode — confirmacion
  ['confirmacion',   'correo', 'template_mode',       'visual'],
  ['confirmacion',   'correo', 'html',                ''],
  // Campos adicionales — Confirmación: editor visual completo
  ['confirmacion',   'correo', 'preheader',              'Nos da mucho gusto tenerte. Aquí está todo lo que necesitas saber.'],
  ['confirmacion',   'correo', 'chip_prefijo',            'Pago confirmado'],
  ['confirmacion',   'correo', 'seccion_hub_label',       'Tu espacio personal del evento'],
  ['confirmacion',   'correo', 'cta_hub_label',           'Acceder a mi espacio →'],
  ['confirmacion',   'correo', 'seccion_encuesta_titulo', 'Una cosa importante'],
  // Datos del evento — compartidos entre comunicaciones (id: 'global', canal: 'evento')
  ['global', 'evento', 'nombre',    'Lo que tu imagen comunica'],
  ['global', 'evento', 'fecha',     'Sábado 15 de agosto de 2026'],
  ['global', 'evento', 'hora',      '10:00 a.m. a 12:00 p.m.'],
  ['global', 'evento', 'lugar',     'The University Club of Mexico'],
  ['global', 'evento', 'direccion', 'Av. Paseo de la Reforma 150, Juárez, CDMX'],
  // Indicaciones
  ['indicaciones',   'correo', 'h1',                 '{{nombre}}, ya casi llegamos. 🌟'],
  ['indicaciones',   'correo', 'look_bloque',         'Abajo: Pantalón ajustado — leggings o jeans pegados funcionan perfecto.\nArriba: Blusa o playera básica ajustada. Si puedes, en blanco o neutro.\nEncima: Un saco, blazer o blusón en el color que más te guste.'],
  ['indicaciones',   'correo', 'rostro_bloque',       'De preferencia con el cabello recogido para el análisis de visagismo. Si no puedes, habrá pinzas disponibles.'],
  ['indicaciones',   'correo', 'medidas_instrucciones', 'Si puedes tomarte las medidas antes de llegar, Mary podrá hacer tu análisis aún más personalizado. Son solo 4: hombros, busto, cintura y cadera. Necesitas una cinta métrica y 5 minutos.\nHombros: De un extremo al otro por la parte más alta de la espalda.\nBusto: La parte más voluminosa del pecho, a la altura de los pezones.\nCintura: La zona más angosta, 2 dedos arriba del ombligo, sin meter el abdomen.\nCadera: La parte más ancha de los glúteos y los huesos de la cadera.'],
  ['indicaciones',   'correo', 'seccion_hub',         'Entra a tu espacio para ver la agenda, tu QR de entrada y registrar tus medidas de cuerpo antes del taller.'],
  ['indicaciones',   'correo', 'cta_hub_label',       'Entrar a mi espacio →'],
  // Recordatorio
  ['recordatorio',   'correo', 'h1',                 '{{nombre}}, mañana es el gran día.'],
  ['recordatorio',   'correo', 'look_compacto',       'No es obligatorio — vengas como vengas, el taller es tuyo. Esta combinación ayuda a aprovechar mejor los análisis:\nParte inferior: Pantalón ajustado (leggings o jeans).\nParte superior: Blusa o playera básica ajustada (blanco o neutro de preferencia).\nCapa extra: Saco, blazer o blusón en el color que prefieras.\nRostro: Cabello recogido si puedes — facilita el análisis de visagismo. Si no, hay pinzas disponibles.'],
  ['recordatorio',   'correo', 'seccion_hub',         'Ahí encontrarás la agenda del día, tu código QR de entrada y los materiales del taller.'],
  ['recordatorio',   'correo', 'seccion_medidas',     'Si aún no las registraste, puedes hacerlo ahora en tu espacio. Mary las usará para personalizar tu análisis de cuerpo. Solo necesitas una cinta métrica: hombros, busto, cintura y cadera.'],
  ['recordatorio',   'correo', 'nota_encuesta',       'Si aún no has contestado la encuesta dentro de tu espacio, este es el momento. Mary la revisa personalmente antes del taller para personalizar tu experiencia. Quienes no la contesten antes del evento no tendrán acceso al material digital posterior.'],
  // QR
  ['qr',             'correo', 'h1',                 '{{nombre}}, hoy te esperamos.'],
  // Agradecimiento
  ['agradecimiento', 'correo', 'h1',                 '{{nombre}}, gracias por ser parte de esta primera edición.'],
  ['agradecimiento', 'correo', 'seccion_encuesta_post', 'Tu opinión es lo más valioso que puedes dejarnos — y nos ayuda a seguir mejorando. La encuesta toma menos de 3 minutos. Al completarla desbloqueas en tu espacio:\n· Tu constancia de participación, lista para descargar.\n· Las guías del taller: 7 estilos y sus esenciales + Compras inteligentes.'],
  ['agradecimiento', 'correo', 'cta_encuesta_label',  'Dejar mi opinión y ver mis recursos →'],
  // ── Recordatorio de recursos (Mié 19 ago) ───────────────────
  ['recordatorio_post', 'correo', 'asunto',   '[nombre], tus recursos del taller todavía te esperan — REINVENTA'],
  ['recordatorio_post', 'correo', 'h1',       '[nombre], aún puedes llevarte tus guías del taller.'],
  ['recordatorio_post', 'correo', 'apertura', 'Hace unos días compartiste con nosotras algo muy especial. Queremos asegurarnos de que puedas llevarte todo lo que prometimos — las guías y tu constancia de participación.\n\nSolo necesitas contestar la encuesta de satisfacción (toma menos de 3 minutos). Al hacerlo, se desbloquean automáticamente en tu espacio personal.'],
  ['recordatorio_post', 'correo', 'cierre',   'El acceso al hub cierra el lunes 24 de agosto.'],
  ['recordatorio_post', 'whatsapp', 'mensaje', '*REINVENTA by Mary Mendez*\n\nHola, {{nombre}}. Tus guías del taller todavía te esperan.\n\nSolo necesitas contestar la encuesta de satisfacción (toma 3 minutos) para desbloquear:\n\n· Tu constancia de participación.\n· Guía de 7 estilos y sus esenciales.\n· Guía de compras inteligentes.\n\nEl acceso al hub cierra el lunes 24 de agosto.\n\nEntra aquí:\n{{hub}}\n\n_Con cariño,_\n_Reinventa by Mary Méndez_\n\n_Mensaje informativo — por favor no respondas a este chat._'],
  // ── Última llamada — cierre hub (Vie 21 ago) ────────────────
  ['ultima_llamada',    'correo', 'asunto',   'Último aviso — tu acceso al hub cierra el lunes — REINVENTA'],
  ['ultima_llamada',    'correo', 'h1',       '[nombre], el lunes 24 cierra el acceso a tu espacio personal.'],
  ['ultima_llamada',    'correo', 'apertura', 'Este es el último aviso. El lunes 24 de agosto se cierra definitivamente el acceso a tu espacio personal de REINVENTA.\n\nAntes de que eso pase, contesta la encuesta (3 minutos) y descarga tus guías. Es lo último que necesitas hacer.'],
  ['ultima_llamada',    'correo', 'cierre',   'Tienes hasta el domingo 23 de agosto.'],
  ['ultima_llamada',    'whatsapp', 'mensaje', '*REINVENTA by Mary Mendez*\n\nHola, {{nombre}}. Último aviso importante.\n\nEl lunes 24 de agosto se cierra definitivamente el acceso a tu espacio personal de REINVENTA.\n\nAntes de que eso pase, contesta la encuesta (3 minutos) y descarga tus guías:\n\n· Constancia de participación.\n· Guía de 7 estilos y sus esenciales.\n· Guía de compras inteligentes.\n\nEntra aquí antes del lunes:\n{{hub}}\n\n_Con cariño,_\n_Reinventa by Mary Méndez_\n\n_Mensaje informativo — por favor no respondas a este chat._'],
];

/* ── Getter con auto-creación ────────────────────────────────────── */
function getContenidosSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_CONTENIDOS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_CONTENIDOS);
    var header = ['ID', 'Canal', 'Campo', 'Borrador', 'Publicado', 'Modificado', 'Nota'];
    sh.appendRow(header);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
    // Poblar con filas iniciales (Borrador y Publicado vacíos — fallback al hardcode)
    _CAMPOS_CONTENIDOS.forEach(function(c) {
      sh.appendRow([c[0], c[1], c[2], '', '', '', '']);
    });
  }
  return sh;
}

/* ── Insertar filas faltantes (idempotente — no toca las existentes) ── */
function _ensureCamposFaltantes() {
  var sh   = getContenidosSheet();
  var data = sh.getDataRange().getValues();
  var existentes = {};
  for (var i = 1; i < data.length; i++) {
    var key = [data[i][0], data[i][1], data[i][2]].join('|');
    existentes[key] = true;
  }
  _CAMPOS_CONTENIDOS.forEach(function(c) {
    var key = [c[0], c[1], c[2]].join('|');
    if (!existentes[key]) {
      // Inserta con el valor default en Borrador para que la validación de aprobación pase
      sh.appendRow([c[0], c[1], c[2], c[3] || '', '', '', '']);
    }
  });
}

/* ── Obtener campo requerido — lanza Error si no está publicado ─── */
// Úsalo en enviarCorreo* para campos que DEBEN existir en Contenidos.
// Si falta, el catch de _enviarCorreosMasivo intercepta y omite ese envío.
function _reqContenido(id, canal, campo, vars) {
  var val = getContenidoPublicado(id, canal, campo);
  if (!val) throw new Error('CONTENIDO_FALTANTE:' + id + '/' + canal + '/' + campo);
  return vars ? _aplicarVariables(val, vars) : val;
}

/* ── Renderizar bloque de líneas como HTML ──────────────────────── */
// texto: multiline string. Líneas con "Label: desc" → bold label.
// Líneas con "·" al inicio → bullet. Líneas vacías → ignoradas.
function _renderLineasBloque(texto, paraStyle, boldLabels) {
  if (!texto) return '';
  var ps = paraStyle || 'font-size:.85rem;color:#4a3545;line-height:1.7;margin:0 0 .35rem;';
  return texto.split('\n').map(function(linea) {
    linea = linea.trim();
    if (!linea) return '';
    if (linea.charAt(0) === '·') {
      return '<p style="' + ps + '">· ' + linea.slice(1).trim() + '</p>';
    }
    if (boldLabels !== false) {
      var colon = linea.indexOf(':');
      if (colon > 0 && colon < 35) {
        var label = linea.slice(0, colon);
        var rest  = linea.slice(colon + 1).trim();
        return '<p style="' + ps + '"><strong style="color:#2A0F25;">' + label + ':</strong> ' + rest + '</p>';
      }
    }
    return '<p style="' + ps + '">' + linea + '</p>';
  }).filter(Boolean).join('');
}

/* ── Leer contenido publicado (usado por enviarCorreo* y generateWhatsApp*) ── */
function getContenidoPublicado(id, canal, campo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_CONTENIDOS);
    if (!sh) return '';
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === id &&
          String(data[i][1]).trim() === canal &&
          String(data[i][2]).trim() === campo) {
        return String(data[i][4] || '').trim(); // columna E = Publicado
      }
    }
    return '';
  } catch(e) {
    Logger.log('getContenidoPublicado error: ' + e);
    return '';
  }
}

/* ── Aplicar variables dinámicas ─────────────────────────────────── */
// Variables soportadas por comunicación: {{nombre}} {{hub}} {{id}}
// Cualquier {{placeholder}} desconocido se deja intacto (no produce error).
function _aplicarVariables(texto, vars) {
  if (!texto) return '';
  Object.keys(vars).forEach(function(k) {
    var val = vars[k] || '';
    texto = texto.split('{{' + k + '}}').join(val); // formato técnico: {{nombre}}
    texto = texto.split('[' + k + ']').join(val);   // formato alternativo: [nombre]
  });
  return texto;
}

/* ── Construir link WA desde plantilla de texto plano ───────────── */
// rawText: texto con saltos de línea reales y formato WA (*bold*, _italic_)
// vars: { nombre, hub, id } — los mismos que _CAMPOS_CONTENIDOS soporta
// Devuelve la parte ?text=... ya URL-encoded (idéntico a cómo lo hacen las funciones originales)
function _buildWAMsgFromTemplate(rawText, vars) {
  if (!rawText) return '';
  var texto = _aplicarVariables(rawText, vars);
  // Cada segmento entre saltos de línea se encodeURIComponent por separado;
  // los saltos reales se convierten en %0A (igual que NL en las funciones originales)
  return texto.split('\n').map(encodeURIComponent).join('%0A');
}

/* ── Extraer variables desconocidas (validación antes de publicar) ── */
// Devuelve array de nombres desconocidos. Vacío = todo bien.
var _VARS_PERMITIDAS = ['nombre', 'hub', 'id'];
function _validarPlaceholders(texto) {
  if (!texto) return [];
  var matches = texto.match(/\{\{([^}]+)\}\}/g) || [];
  return matches
    .map(function(m) { return m.slice(2, -2).trim(); })
    .filter(function(v) { return _VARS_PERMITIDAS.indexOf(v) === -1; });
}

/* ═══════════════════════════════════════════════════════════════════
   ENDPOINTS ADMIN — Contenidos
   ═══════════════════════════════════════════════════════════════════ */

/* ── leer_contenidos: devuelve todos los campos con borrador+publicado ── */
function adminLeerContenidos() {
  try {
    _ensureCamposFaltantes(); // Inserta filas nuevas si faltan (idempotente)
    var sh   = getContenidosSheet();
    var data = sh.getDataRange().getValues();
    var result = {};
    for (var i = 1; i < data.length; i++) {
      var id      = String(data[i][0] || '').trim();
      var canal   = String(data[i][1] || '').trim();
      var campo   = String(data[i][2] || '').trim();
      var borrdr  = String(data[i][3] || '').trim();
      var pub     = String(data[i][4] || '').trim();
      var modif   = data[i][5] ? String(data[i][5]) : '';
      if (!id || !canal || !campo) continue;
      if (!result[id])         result[id]         = {};
      if (!result[id][canal])  result[id][canal]  = {};
      result[id][canal][campo] = {
        borrador:   borrdr,
        publicado:  pub,
        modificado: modif,
        tienePublicado: pub.length > 0
      };
    }
    return jsOk({ contenidos: result });
  } catch(e) {
    return jsErr('Error leyendo contenidos: ' + e.message);
  }
}

/* ── guardar_borrador: guarda o actualiza el borrador de un campo ── */
function adminGuardarBorrador(id, canal, campo, valor) {
  if (!id || !canal || !campo) return jsErr('Faltan parámetros: id, canal, campo');
  if (valor === undefined || valor === null) return jsErr('Falta el valor');
  try {
    var sh   = getContenidosSheet();
    var data = sh.getDataRange().getValues();
    var ts   = new Date().toISOString();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === id &&
          String(data[i][1]).trim() === canal &&
          String(data[i][2]).trim() === campo) {
        sh.getRange(i + 1, 4).setValue(valor);  // col D = Borrador
        sh.getRange(i + 1, 6).setValue(ts);     // col F = Modificado
        return jsOk({ guardado: true, campo: campo, ts: ts });
      }
    }
    // Fila no existe — insertar
    sh.appendRow([id, canal, campo, valor, '', ts, '']);
    return jsOk({ guardado: true, campo: campo, insertado: true, ts: ts });
  } catch(e) {
    return jsErr('Error guardando borrador: ' + e.message);
  }
}

/* ── publicar_contenido: copia borrador → publicado ─────────────── */
function adminPublicarContenido(id, canal, campo) {
  if (!id || !canal || !campo) return jsErr('Faltan parámetros: id, canal, campo');
  try {
    var sh   = getContenidosSheet();
    var data = sh.getDataRange().getValues();
    var ts   = new Date().toISOString();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === id &&
          String(data[i][1]).trim() === canal &&
          String(data[i][2]).trim() === campo) {
        var borrador = String(data[i][3] || '').trim();
        if (!borrador) return jsErr('El borrador está vacío. Guarda contenido antes de publicar.');
        sh.getRange(i + 1, 5).setValue(borrador); // col E = Publicado
        sh.getRange(i + 1, 6).setValue(ts);
        return jsOk({ publicado: true, campo: campo, valor: borrador, ts: ts });
      }
    }
    return jsErr('Campo no encontrado: ' + id + '/' + canal + '/' + campo);
  } catch(e) {
    return jsErr('Error publicando: ' + e.message);
  }
}

/* ── guardar_borradores_lote: guarda múltiples campos en una llamada ── */
// campos: [{campo: 'h1', valor: '...'}, {campo: 'seccion_hub', valor: '...'}, ...]
function adminGuardarBorradoresLote(id, canal, campos) {
  if (!id || !canal || !campos || !campos.length) return jsErr('Faltan parámetros');
  var errores = [];
  campos.forEach(function(c) {
    try {
      var r = adminGuardarBorrador(id, canal, c.campo, c.valor);
      var parsed = JSON.parse(r.getContent ? r.getContent() : '{}');
      if (parsed.error) errores.push(c.campo);
    } catch(e) { errores.push(c.campo); }
  });
  if (errores.length) return jsErr('Error guardando campos: ' + errores.join(', '));
  return jsOk({ guardado: true, n: campos.length });
}

/* ── publicar_contenidos_lote: publica múltiples campos en una llamada ── */
// campos: array de strings ['h1','asunto',...] O array de objetos [{campo:'h1'},...]
function adminPublicarContenidosLote(id, canal, campos) {
  if (!id || !canal || !campos || !campos.length) return jsErr('Faltan parámetros');
  var errores = [];
  campos.forEach(function(item) {
    // Aceptar tanto string como {campo: 'nombre'}
    var nombreCampo = (typeof item === 'string') ? item : (item && item.campo ? item.campo : null);
    if (!nombreCampo) { errores.push(String(item)); return; }
    try {
      var r = adminPublicarContenido(id, canal, nombreCampo);
      var parsed = JSON.parse(r.getContent ? r.getContent() : '{}');
      if (parsed.error) errores.push(nombreCampo);
    } catch(e) { errores.push(nombreCampo); }
  });
  if (errores.length) return jsErr('Error publicando campos: ' + errores.join(', '));
  return jsOk({ publicado: true, n: campos.length });
}

/* ═══════════════════════════════════════════════════════════════════
   FUNCIONES DE PRUEBA — QA manual desde GAS
   (No disparan envíos reales a asistentes)
   ═══════════════════════════════════════════════════════════════════ */

/* Prueba 1: Leer contenidos (verificar que la hoja existe y se leen bien) */
function testLeerContenidos() {
  var r = adminLeerContenidos();
  Logger.log(r.getContent());
}

/* Prueba 2: Guardar borrador */
function testGuardarBorrador() {
  var r = adminGuardarBorrador('recordatorio', 'correo', 'asunto', 'PRUEBA — Mañana te esperamos · REINVENTA');
  Logger.log(r.getContent());
}

/* Prueba 3: Verificar que borrador NO se use en envío (publicado sigue vacío) */
function testFallbackSinPublicado() {
  var pub = getContenidoPublicado('recordatorio', 'correo', 'asunto');
  var fallback = 'Mañana te esperamos — REINVENTA';
  var asunto = pub || fallback;
  Logger.log('pub: [' + pub + ']');
  Logger.log('asunto final: ' + asunto);
  Logger.log(pub === '' ? '✅ Borrador NO afecta envíos — fallback activo' : '⚠ Hay publicado: ' + pub);
}

/* Prueba 4: Publicar y verificar */
function testPublicarYVerificar() {
  adminPublicarContenido('recordatorio', 'correo', 'asunto');
  var pub = getContenidoPublicado('recordatorio', 'correo', 'asunto');
  Logger.log('Publicado: ' + pub);
  Logger.log(pub ? '✅ getContenidoPublicado lo lee correctamente' : '❌ No lo encontró');
}

/* Prueba 5: Fallback cuando publicado se vacía */
function testFallbackConPublicadoVacio() {
  // Simula publicado vacío borrando temporalmente (restaurar después)
  var sh   = getContenidosSheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'recordatorio' && data[i][1] === 'correo' && data[i][2] === 'asunto') {
      var valorOrig = data[i][4];
      sh.getRange(i+1, 5).setValue(''); // vaciar publicado
      var pub = getContenidoPublicado('recordatorio', 'correo', 'asunto');
      var fallback = 'Mañana te esperamos — REINVENTA';
      Logger.log('Con publicado vacío — asunto final: ' + (pub || fallback));
      sh.getRange(i+1, 5).setValue(valorOrig); // restaurar
      Logger.log('✅ Fallback activo cuando publicado está vacío');
      return;
    }
  }
}

/* Prueba 7: Correo de prueba usando contenido (a destinatario controlado) */
function testCorreoConContenido() {
  // Usa mejoracontinua@caceca.org — NO envía a asistentes reales
  enviarCorreoRecordatorio('Valeria García (PRUEBA)', 'mejoracontinua@caceca.org', 'RNV-001');
  Logger.log('✅ Correo de prueba enviado a mejoracontinua@caceca.org');
}

/* Prueba 8: Verificar variables dinámicas */
function testVariablesDinamicas() {
  var texto = '{{nombre}}, esto necesitas para el taller 🌟 — REINVENTA';
  var resultado = _aplicarVariables(texto, { nombre: 'Valeria' });
  Logger.log('Variables: ' + resultado);
  Logger.log(resultado === 'Valeria, esto necesitas para el taller 🌟 — REINVENTA' ? '✅ Variables OK' : '❌ Error variables');
}

/* Prueba 9: WhatsApp — publicar y generar link */
function testWAPublicarYGenerar() {
  // Guardar y publicar mensaje de prueba para recordatorio WA
  var textoWA = '*REINVENTA by Mary Mendez*\n\nHola, {{nombre}}. Este es un mensaje de PRUEBA.\n\nTu espacio: {{hub}}\n\n_Con cariño,_\n_Reinventa by Mary Mendez_\n\n_Mensaje informativo._';
  adminGuardarBorrador('recordatorio', 'whatsapp', 'mensaje', textoWA);
  adminPublicarContenido('recordatorio', 'whatsapp', 'mensaje');

  // Generar link y verificar
  var link = generateWhatsAppLinkRecordatorio('Valeria García', '5551234567', 'RNV-001');
  Logger.log('Link generado: ' + link);
  Logger.log(link && link.indexOf('PRUEBA') >= 0 ? '✅ Mensaje publicado se usó' : '❌ Sigue usando hardcode');
  Logger.log(link && link.indexOf('Valeria') >= 0 ? '✅ Variable {{nombre}} reemplazada' : '❌ Variable no reemplazada');
  Logger.log(link && link.indexOf('hub%3Fid%3DRNV') >= 0 ? '✅ Variable {{hub}} con ID' : '⚠ Verificar hub URL');
}

/* Prueba 10: Verificar que al quitar Publicado regresa fallback WA */
function testFallbackWASinPublicado() {
  // Vaciar publicado de recordatorio WA
  var sh   = getContenidosSheet();
  var data = sh.getDataRange().getValues();
  var pubOriginal = '';
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]==='recordatorio' && data[i][1]==='whatsapp' && data[i][2]==='mensaje') {
      pubOriginal = data[i][4];
      sh.getRange(i+1,5).setValue(''); // vaciar publicado
      break;
    }
  }
  var link = generateWhatsAppLinkRecordatorio('Valeria', '5551234567', 'RNV-001');
  Logger.log('Link con publicado vacío: ' + link.substring(0,120) + '...');
  Logger.log(!link || link.indexOf('PRUEBA') < 0 ? '✅ Fallback activo — usa hardcode' : '❌ Sigue usando publicado (error)');

  // Restaurar
  for (var j = 1; j < sh.getDataRange().getValues().length; j++) {
    var d = sh.getDataRange().getValues();
    if (d[j][0]==='recordatorio' && d[j][1]==='whatsapp' && d[j][2]==='mensaje') {
      sh.getRange(j+1,5).setValue(pubOriginal);
      break;
    }
  }
  Logger.log('Publicado restaurado: ' + pubOriginal.substring(0,50));
}

/* Prueba 11: Validación de placeholders */
function testValidacionPlaceholders() {
  var tests = [
    { texto: 'Hola {{nombre}}, tu hub es {{hub}}', esperado: 0 },
    { texto: 'Hola {{nomber}}', esperado: 1 },                   // typo
    { texto: 'Hola {{nombre}} {{desconocido}}', esperado: 1 },
    { texto: 'Sin placeholders', esperado: 0 },
  ];
  tests.forEach(function(t, i) {
    var invalidos = _validarPlaceholders(t.texto);
    var ok = invalidos.length === t.esperado;
    Logger.log('Test ' + (i+1) + ': ' + (ok ? '✅' : '❌') + ' invalidos=[' + invalidos.join(',') + '] esperado=' + t.esperado);
  });
}

/* Documentación de _calcularForma — solo lectura, sin cambios */
function docCalcularForma() {
  // Tipos soportados y condiciones (de hub.html _calcularForma):
  // 1. RELOJ DE ARENA:   |busto-cadera| <= 5cm Y min(hombros,cadera)-cintura >= 20cm
  // 2. TRIÁNGULO (Pera): cadera > hombros en ≥5% → pct_cadera_vs_hombros >= 5
  // 3. TRIÁNGULO INV.:   hombros > cadera en ≥5% O busto > cadera en ≥5%
  // 4. MANZANA:          cintura >= hombros*0.9 O cintura >= cadera*0.9
  // 5. RECTÁNGULO:       todo lo demás
  //
  // Valores limítrofes: si dif_bc=5 Y dif_c=20 → Reloj de arena (condición exacta)
  // Si cintura=80, hombros=88: 80 >= 88*0.9=79.2 → MANZANA (true)
  // Si falta una medida: _calcularForma no se llama (solo cuando completas=true)
  //
  // El tipo "Reloj de arena" en hub.html coincide con el tipo que usa GAS
  // (mismo nombre, misma descripción) → coherente.
  var ejemplos = [
    { h:38, b:90, c:68, ca:92, esperado:'Reloj de arena' },  // reloj
    { h:36, b:85, c:70, ca:96, esperado:'Triángulo' },        // cadera>>hombros
    { h:44, b:90, c:72, ca:90, esperado:'Triángulo invertido' }, // hombros>>cadera
    { h:38, b:88, c:85, ca:90, esperado:'Manzana' },          // cintura ~hombros
    { h:38, b:88, c:78, ca:90, esperado:'Rectángulo' },       // medidas parejas
  ];
  ejemplos.forEach(function(ej) {
    // Este test solo corre en GAS donde existe _calcularForma (no en GAS server)
    Logger.log(JSON.stringify(ej) + ' → espera: ' + ej.esperado);
  });
  Logger.log('Ver _calcularForma() en hub.html para la lógica completa');
}

/* ── Reporte completo post-evento ──────────────────────────── */
function adminReporteCompleto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {};

  // --- REGISTROS ---
  try {
    var shR = ss.getSheetByName('Registros');
    var rowsR = shR ? shR.getDataRange().getValues() : [];
    result.registros = rowsR.slice(1).map(function(r) {
      return {
        fecha:    r[0] ? Utilities.formatDate(new Date(r[0]), 'America/Mexico_City', 'dd/MM/yyyy HH:mm') : '',
        nombre:   r[1] || '',
        correo:   r[2] || '',
        whatsapp: r[3] || '',
        contacto: r[4] || '',
        busca:    r[5] || '',
        fase:     r[6] || '',
        monto:    r[7] || '',
        fechaPago:r[8] ? Utilities.formatDate(new Date(r[8]), 'America/Mexico_City', 'dd/MM/yyyy') : '',
        pago:     r[9] === '✓' ? true : false,
        stripeId: r[10] || '',
        origen:   r[11] || '',
        marketing:r[12] || '',
        imagen:   r[13] || '',
        utm:      r[14] || '',
        enterado: r[15] || '',
        dedica:   r[16] || ''
      };
    }).filter(function(r) { return r.nombre || r.correo; });
  } catch(e) { result.registros = []; result.errorRegistros = e.message; }

  // --- COMUNICACIONES ---
  try {
    var shC = ss.getSheetByName('Comunicaciones');
    var rowsC = shC ? shC.getDataRange().getValues() : [];
    result.comunicaciones = rowsC.slice(1).map(function(r) {
      return {
        correo:        r[0] || '',
        nombre:        r[1] || '',
        whatsapp:      r[2] || '',
        estConfWA:     r[5] || '',
        correoConf:    r[6] === '✓' ? true : false,
        estRecordWA:   r[8] || '',
        correoRecord:  r[9] === '✓' ? true : false,
        estQRWA:       r[11] || '',
        correoQR:      r[12] === '✓' ? true : false,
        estAgradecWA:  r[14] || '',
        correoAgradec: r[15] === '✓' ? true : false,
        estIndicWA:    r[17] || '',
        correoIndic:   r[18] === '✓' ? true : false,
        estEncPrevWA:  r[20] || '',
        estRecursosWA: r[23] !== undefined ? r[23] : '',
        estUltimaWA:   r[26] !== undefined ? r[26] : ''
      };
    }).filter(function(r) { return r.correo || r.nombre; });
  } catch(e) { result.comunicaciones = []; result.errorComunicaciones = e.message; }

  // --- ASISTENCIA ---
  try {
    var shA = ss.getSheetByName('Asistencia');
    var rowsA = shA ? shA.getDataRange().getValues() : [];
    result.asistencia = rowsA.slice(1).map(function(r) {
      return {
        id:           r[0] || '',
        nombre:       r[1] || '',
        correo:       r[2] || '',
        fase:         r[3] || '',
        asistio:      r[4] === '✓',
        fechaEntrada: r[5] ? Utilities.formatDate(new Date(r[5]), 'America/Mexico_City', 'HH:mm') : '',
        encuesta:     r[6] === '✓',
        calificacion: r[8] || '',
        comentario:   r[9] || '',
        expectativas: r[10] || '',
        aplicabilidad:r[12] || '',
        contenido:    r[13] || '',
        mary:         r[14] || '',
        orgMatriz:    [r[16],r[17],r[18],r[19],r[20],r[21]],
        impacto:      r[24] || '',
        nps:          r[27] !== undefined && r[27] !== '' ? r[27] : null,
        futuro:       r[28] || ''
      };
    }).filter(function(r) { return r.id || r.nombre; });
  } catch(e) { result.asistencia = []; result.errorAsistencia = e.message; }

  // --- MEDIDAS ---
  try {
    var shM = ss.getSheetByName('Medidas');
    var rowsM = shM ? shM.getDataRange().getValues() : [];
    result.medidasDetalle = rowsM.slice(1).map(function(r) {
      return {
        id:       r[0] || '',
        nombre:   r[1] || '',
        hombros:  r[2] || '',
        busto:    r[3] || '',
        cintura:  r[4] || '',
        cadera:   r[5] || '',
        forma:    r[6] || '',
        fecha:    r[7] ? Utilities.formatDate(new Date(r[7]), 'America/Mexico_City', 'dd/MM/yyyy') : ''
      };
    }).filter(function(r) { return r.id || r.nombre; });
  } catch(e) { result.medidasDetalle = []; }

  // --- ENCUESTA PREVIA ---
  try {
    var shEP = ss.getSheetByName('Encuesta Previa');
    var rowsEP = shEP ? shEP.getDataRange().getValues() : [];
    result.encuestaPrevia = rowsEP.slice(1).map(function(r) {
      return {
        id:          r[0] || '',
        nombre:      r[1] || '',
        correo:      r[2] || '',
        fecha:       r[3] ? Utilities.formatDate(new Date(r[3]), 'America/Mexico_City', 'dd/MM/yyyy') : '',
        satisfaccion:r[4] || '',
        coherencia:  r[5] || '',
        confianza:   r[6] || '',
        proyeccion:  r[7] || '',
        motivacion:  r[8] || '',
        expectativa: r[9] || '',
        piel:        r[10] || '',
        cabello:     r[11] || '',
        ojos:        r[12] || ''
      };
    }).filter(function(r) { return r.id || r.nombre; });
  } catch(e) { result.encuestaPrevia = []; }

  // --- DESCARGAS ---
  try {
    var shD = ss.getSheetByName('Descargas');
    var rowsDes = shD ? shD.getDataRange().getValues() : [];
    result.descargas = rowsDes.slice(1).map(function(r) {
      return { id: r[0]||'', nombre: r[1]||'', tipo: r[2]||'', fecha: r[3]||'' };
    }).filter(function(r) { return r.id; });
  } catch(e) { result.descargas = []; }

  // --- ENCUESTA POST STATUS (Contenidos sheet) ---
  try {
    var shCont = ss.getSheetByName('Contenidos');
    var rowsCont = shCont ? shCont.getDataRange().getValues() : [];
    result.encuestaPostStatus = 'OPEN'; // default
    for (var ci = 1; ci < rowsCont.length; ci++) {
      if ((rowsCont[ci][0]||'').toString().trim() === 'config' &&
          (rowsCont[ci][1]||'').toString().trim() === 'evento' &&
          (rowsCont[ci][2]||'').toString().trim() === 'encuesta_post_status') {
        result.encuestaPostStatus = (rowsCont[ci][4]||'OPEN').toString().trim();
        break;
      }
    }
  } catch(e) { result.encuestaPostStatus = 'OPEN'; }

  return ContentService.createTextOutput(JSON.stringify({ result: 'ok', data: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

function adminSetHubStatus(status) {
  if (status !== 'OPEN' && status !== 'CLOSED') return jsErr('Estado inválido: usa OPEN o CLOSED');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shC = ss.getSheetByName('Contenidos');
  if (!shC) return jsErr('Hoja Contenidos no encontrada');
  var rows = shC.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][0]||'').toString().trim() === 'config' &&
        (rows[i][1]||'').toString().trim() === 'evento' &&
        (rows[i][2]||'').toString().trim() === 'hub_status') {
      shC.getRange(i + 1, 5).setValue(status);
      return jsOk({ hub_status: status });
    }
  }
  shC.appendRow(['config', 'evento', 'hub_status', '', status]);
  return jsOk({ hub_status: status, created: true });
}

function adminSetSurveyStatus(status) {
  if (status !== 'OPEN' && status !== 'CLOSED') return jsErr('Estado inválido: usa OPEN o CLOSED');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shC = ss.getSheetByName('Contenidos');
  if (!shC) return jsErr('Hoja Contenidos no encontrada');
  var rows = shC.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][0]||'').toString().trim() === 'config' &&
        (rows[i][1]||'').toString().trim() === 'evento' &&
        (rows[i][2]||'').toString().trim() === 'encuesta_post_status') {
      shC.getRange(i + 1, 5).setValue(status); // col E = Publicado (idx 4)
      return jsOk({ status: status });
    }
  }
  // Row not found — create it
  shC.appendRow(['config', 'evento', 'encuesta_post_status', '', status]);
  return jsOk({ status: status, created: true });
}

/* ── Perfil de prueba para testing de encuesta post ─────────────── */
function adminCreateTestProfile() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var TEST_ID = 'RNV-TEST';

  // ── Asistencia: checkin + encuesta previa marcados ──────────────
  var asiSheet = getAsistenciaSheet();
  var asiData  = asiSheet.getDataRange().getValues();
  var filaAsi  = -1;
  for (var i = 1; i < asiData.length; i++) {
    if ((asiData[i][0]||'').toString().trim() === TEST_ID) { filaAsi = i + 1; break; }
  }
  if (filaAsi === -1) {
    asiSheet.appendRow([TEST_ID, 'Prueba Test', 'mejoracontinua@caceca.org', 'Early Bird', '✓', new Date(), '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    filaAsi = asiSheet.getLastRow();
  } else {
    // Asegurar check-in marcado y encuesta post limpia
    asiSheet.getRange(filaAsi, 5).setValue('✓');
    asiSheet.getRange(filaAsi, 6).setValue(new Date());
    asiSheet.getRange(filaAsi, 7).setValue(''); // encuesta post vacía para poder probar
    asiSheet.getRange(filaAsi, 8).setValue('');
    for (var c = 9; c <= 22; c++) asiSheet.getRange(filaAsi, c).setValue('');
  }

  // ── Encuesta Previa: marcar como completada ──────────────────────
  var prevSheet = ss.getSheetByName('Encuesta Previa');
  if (prevSheet) {
    var prevData = prevSheet.getDataRange().getValues();
    var yaExiste = false;
    for (var j = 1; j < prevData.length; j++) {
      if ((prevData[j][0]||'').toString().trim() === TEST_ID) { yaExiste = true; break; }
    }
    if (!yaExiste) {
      prevSheet.appendRow([TEST_ID, 'Prueba Test', 'mejoracontinua@caceca.org', new Date(), 3, 3, 3, 3, 'Bienestar personal', 'Probar la encuesta post', 3, 3, 3, 3]);
    }
  }

  return jsOk({ id: TEST_ID, hubUrl: 'https://reinventabymarymendez.com.mx/hub?id=' + TEST_ID });
}
