// ================================================================
// Research Assessor — Apps Script v7.7
// Novedades v7.7 (privacidad):
//   • verificar (público) solo confirma existencia, estado, firmantes (cliente con iniciales) y huellas.
//     Ya no expone valor, título, plan de pagos, detalles, cédulas ni enlaces a los documentos.
//   • historialContrato (solo administradores): registro completo con detalles y enlaces.
// ----------------------------------------------------------------
// Research Assessor — Apps Script v7.6.1
// Novedades v7.6.1: autorizarPermisos pide TODOS los permisos (Drive con escritura incluido)
//   y comprueba creando las carpetas de contratos. Marca todas las casillas en la pantalla de Google.
// Novedades v7.6:
//   • autorizarPermisos(): ejecútala UNA VEZ desde el editor (botón ▶ Ejecutar) para
//     conceder a la aplicación el acceso a Google Drive y al correo. Sin ese permiso
//     no se pueden guardar ni firmar contratos ("No tienes permiso para llamar a DriveApp…").
//   • ping informa si los permisos de Drive y correo están concedidos.
// ----------------------------------------------------------------
// Research Assessor — Apps Script v7.5
// Novedades v7.5: el contrato se emite como PDF firmado con el certificado .p12 del prestador
//   (emitirContrato con formato:'pdf' y firmado:true registra emisión y firma en un solo paso)
// ----------------------------------------------------------------
// Research Assessor — Apps Script v7.4
// Novedades v7.4 (auditoría de contratos):
//   • Hojas protegidas Contratos y AuditoriaContratos (no editables desde la API genérica)
//   • emitirContrato / registrarFirmado / anularContrato (requieren usuario administrador)
//   • verificar (público): registro oficial + cadena de auditoría con HMAC-SHA256
//   • Huellas SHA-256 calculadas en el servidor; hora oficial del servidor; comprobantes por correo
// ----------------------------------------------------------------
// Research Assessor — Apps Script v7.1
// Novedades v7.1:
//   • Nueva hoja: Propuestas (proformas con varios artículos, JSON)
//   • Clientes: columna area · Cuotas: columna hitoEtapa
//   • guardarContrato también guarda proformas (propuestaId, carpeta)
// ----------------------------------------------------------------
// Research Assessor — Apps Script v7
// Novedades v7:
//   • Nueva hoja: Seguimientos (avances del proceso de publicación)
//   • Nueva hoja: Config (datos del prestador para los contratos)
//   • Trabajos: columnas detalles, etapa, contratoUrl, contratoFecha
//   • Acción guardarContrato: guarda el .docx en Google Drive y devuelve
//     un enlace de solo lectura para compartir por WhatsApp
//   • insertar/actualizar crean hojas y columnas faltantes automáticamente
// ----------------------------------------------------------------
// Research Assesor — Apps Script v6
// Novedades v6:
//   • Nueva hoja: Colaboradores (usuarios internos del sistema)
//   • Nueva hoja: Asignaciones (trabajos asignados a colaboradores + valor/pago)
//   • getData incluye colaboradores y asignaciones
//   • repararHojas crea ambas hojas automáticamente
// Novedades v6.1:
//   • Nueva hoja: Usuarios (administradores del sistema, persistente)
//   • getData incluye usuarios
// ================================================================

var SS_ID = '1ILRbM7sLA3Tsk54PNbGbDrwBMMZqUSts6V-0-v-dai8';

var ESQUEMA = {
  Clientes:       ['id','nombre','cedula','telefono','email','direccion','institucion','ciudad','createdAt','notas','area'],
  Trabajos:       ['id','clienteId','clienteNombre','titulo','tipo','estado','total','progreso','notas','fechaInicio','fechaFin','carpeta','createdAt','detalles','etapa','contratoUrl','contratoFecha'],
  Cuotas:         ['id','trabajoId','clienteId','clienteNombre','trabajoTitulo','label','fechaVencimiento','acordado','pagado','estado','hitoEtapa'],
  Abonos:         ['id','cuotaId','trabajoId','fecha','monto','nota','comprobante'],
  Reuniones:      ['id','clienteId','trabajoId','titulo','fecha','hora','plataforma','link','notas'],
  UsuariosCliente:['id','clienteId','nombre','usuario','password','role','activo'],
  // ── NUEVO v6 ──
  Colaboradores:  ['id','nombre','email','telefono','especialidad','usuario','password','activo','createdAt','notas'],
  Asignaciones:   ['id','colaboradorId','colaboradorNombre','trabajoId','trabajoTitulo','clienteNombre','descripcion','valorAsignado','estado','fechaAsignacion','fechaLimite','fechaPago','pagado','mes','notas'],
  // ── NUEVO v6.1 ──
  Usuarios:       ['id','usuario','password','nombre','role','activo'],
  // ── NUEVO v7 ──
  Seguimientos:   ['id','trabajoId','etapa','fecha','nota','responsable','revista','visibleCliente','createdAt'],
  Config:         ['id','valor'],
  // ── NUEVO v7.1 ──
  Propuestas:     ['id','clienteId','clienteNombre','titulo','area','fecha','validez','estado','articulos','hitos','notas','docUrl','docFecha','createdAt'],
  // ── NUEVO v7.4 (auditoría) ──
  Contratos:      ['id','trabajoId','clienteId','clienteNombre','titulo','tipo','valor','estado','version','fechaEmision','hashEmitido','urlEmitido',
                   'fechaFirmaPrestador','hashFirmaPrestador','urlFirmaPrestador','firmasPrestador',
                   'fechaFirmaAmbos','hashFirmaAmbos','urlFirmaAmbos','firmasAmbos','resumen','reemplazadoPor','createdAt'],
  AuditoriaContratos: ['id','contratoId','fecha','evento','actor','detalle','hashDocumento','hashPrevio','hashEvento']
};

var CARPETA_CONTRATOS = 'Research Assessor — Contratos';
var CARPETA_EMITIDOS  = 'Research Assessor — Contratos emitidos';
var CARPETA_FIRMADOS  = 'Research Assessor — Contratos firmados';
// Hojas que solo se escriben mediante las acciones de auditoría (nunca con insertar/actualizar/eliminar)
var PROTEGIDAS = ['Contratos', 'AuditoriaContratos'];
// Si la hoja Usuarios está vacía se aceptan los administradores por defecto de la app.
// Crea tus propios usuarios en la app (Usuarios) para reemplazarlos.
var ADMINS_POR_DEFECTO = [['admin','research2024'],['edison','asesor2024']];

function verificarTablaEditable(tabla) {
  if (PROTEGIDAS.indexOf(tabla) > -1) throw new Error('La hoja ' + tabla + ' está protegida (auditoría)');
}

// ── Estado de Asignaciones ────────────────────────────────────────
// pendiente  → asignado pero no completado
// completado → el colaborador terminó su parte
// pagado     → el admin ya pagó al colaborador

// ── PERMISOS ─────────────────────────────────────────────────────
// Selecciona "autorizarPermisos" en la barra superior del editor y pulsa ▶ Ejecutar.
// Google pedirá permiso para Drive, correo y la hoja: acepta con tu cuenta.
// Google concede los permisos uno por uno: en la pantalla de autorización marca
// "Seleccionar todo" (o todas las casillas) antes de pulsar Continuar/Permitir.
function autorizarPermisos() {
  // Pide de una vez TODOS los permisos que usa el proyecto (Drive completo, correo, hoja, calendario)
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  SpreadsheetApp.openById(SS_ID).getName();
  // Prueba real de escritura en Drive: crea las carpetas de contratos y un archivo temporal
  [CARPETA_CONTRATOS, CARPETA_EMITIDOS, CARPETA_FIRMADOS].forEach(function(n) { carpetaContratos(n); });
  var prueba = carpetaContratos(CARPETA_CONTRATOS).createFile('prueba-permisos.txt', 'Research Assessor: prueba de permisos');
  prueba.setTrashed(true);
  MailApp.getRemainingDailyQuota();
  var p = revisarPermisos_();
  Logger.log(p.completo
    ? 'Permisos concedidos: Drive (lectura y escritura), correo y hoja de cálculo. Ya puedes firmar contratos.'
    : 'Aún faltan permisos. Vuelve a ejecutar y marca TODAS las casillas en la pantalla de Google.');
  return p;
}
function revisarPermisos_() {
  var p = { drive: true, mail: true, completo: true };
  try {
    var info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    p.completo = info.getAuthorizationStatus() !== ScriptApp.AuthorizationStatus.REQUIRED;
  } catch (err) { p.completo = false; }
  // Pruebas que no modifican nada (la lectura de Drive no garantiza la escritura: eso lo dice "completo")
  try { DriveApp.getRootFolder().getId(); } catch (err) { p.drive = false; }
  try { MailApp.getRemainingDailyQuota(); } catch (err) { p.mail = false; }
  return p;
}

function doGet(e) {
  var action   = e.parameter.action || 'getData';
  var callback = e.parameter.callback || '';
  var result;
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    if (action === 'getData') {
      result = {
        ok: true,
        clientes:        leer(ss, 'Clientes'),
        trabajos:        leer(ss, 'Trabajos'),
        cuotas:          leer(ss, 'Cuotas'),
        abonos:          leer(ss, 'Abonos'),
        reuniones:       leer(ss, 'Reuniones'),
        usuariosCliente: leer(ss, 'UsuariosCliente'),
        colaboradores:   leer(ss, 'Colaboradores'),
        asignaciones:    leer(ss, 'Asignaciones'),
        usuarios:        leer(ss, 'Usuarios'),
        seguimientos:    leer(ss, 'Seguimientos'),
        config:          leer(ss, 'Config'),
        propuestas:      leer(ss, 'Propuestas'),
        contratos:       leer(ss, 'Contratos')
      };
    } else if (action === 'verificar') {
      result = verificarContrato(ss, e.parameter.codigo || '');
    } else if (action === 'write') {
      var body = JSON.parse(e.parameter.payload || '{}');
      verificarTablaEditable(body.tabla);
      if      (body.action === 'insertar')   { insertar(ss, body.tabla, body.fila); }
      else if (body.action === 'actualizar') { actualizar(ss, body.tabla, body.id, body.fila); }
      else if (body.action === 'eliminar')   { eliminar(ss, body.tabla, body.id); }
      else { throw new Error('Op desconocida: ' + body.action); }
      result = { ok: true };
    } else if (action === 'ping') {
      result = { ok: true, msg: 'OK', script: 'v7.7', permisos: revisarPermisos_() };
    } else if (action === 'createMeet') {
      var title    = (e.parameter.title    || 'Reunión Research Assessor');
      var date     = (e.parameter.date     || '');
      var time     = (e.parameter.time     || '10:00');
      var duration = parseInt(e.parameter.duration || '60');
      result = crearGoogleMeet(title, date, time, duration);
    } else if (action === 'repararHojas') {
      result = { ok: true, reparado: repararHojas(ss) };
    } else if (action === 'resumenColaborador') {
      // Resumen mensual de un colaborador: asignaciones + totales
      var colId = e.parameter.colaboradorId || '';
      var mes   = e.parameter.mes || '';       // formato YYYY-MM
      result = resumenColaborador(ss, colId, mes);
    } else {
      result = { ok: false, error: 'Accion desconocida: ' + action };
    }
  } catch(err) {
    result = { ok: false, error: err.toString() };
  }
  var json = JSON.stringify(result);
  return ContentService
    .createTextOutput(callback ? callback + '(' + json + ')' : json)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var b  = JSON.parse(e.postData.contents);
    if (b.action === 'guardarContrato') { return respJson(guardarContrato(ss, b)); }
    if (b.action === 'emitirContrato')   { return respJson(emitirContrato(ss, b)); }
    if (b.action === 'registrarFirmado') { return respJson(registrarFirmado(ss, b)); }
    if (b.action === 'anularContrato')   { return respJson(anularContrato(ss, b)); }
    if (b.action === 'historialContrato') { return respJson(historialContrato(ss, b)); }
    verificarTablaEditable(b.tabla);
    if      (b.action === 'insertar')   { insertar(ss, b.tabla, b.fila); }
    else if (b.action === 'actualizar') { actualizar(ss, b.tabla, b.id, b.fila); }
    else if (b.action === 'eliminar')   { eliminar(ss, b.tabla, b.id); }
    else { throw new Error('Accion no reconocida: ' + b.action); }
    return respJson({ ok: true });
  } catch(err) {
    return respJson({ ok: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════
// CRUD BASE
// ════════════════════════════════════════════════════════════════
function leer(ss, nombre) {
  var hoja = ss.getSheetByName(nombre);
  if (!hoja || hoja.getLastRow() < 2) return [];
  var datos = hoja.getDataRange().getValues();
  var cab   = datos[0];
  return datos.slice(1)
    .map(function(f) {
      var o = {};
      cab.forEach(function(h, i) {
        var val = f[i];
        if (val instanceof Date) {
          var y = val.getFullYear();
          if (y === 1899 || y === 1900) {
            // Sheets time-only cell → HH:MM
            var hh = String(val.getHours()).padStart(2, '0');
            var mn = String(val.getMinutes()).padStart(2, '0');
            o[String(h)] = hh + ':' + mn;
          } else {
            // Normal date → YYYY-MM-DD
            var m = String(val.getMonth() + 1).padStart(2, '0');
            var d = String(val.getDate()).padStart(2, '0');
            o[String(h)] = y + '-' + m + '-' + d;
          }
        } else {
          o[String(h)] = val != null ? String(val) : '';
        }
      });
      return o;
    })
    .filter(function(o) { return o.id && o.id.trim() !== ''; });
}

// Devuelve la hoja lista para escribir: la crea si falta y agrega las
// columnas del ESQUEMA que aún no existan (nunca borra ni reordena).
function hojaParaEscribir(ss, nombre) {
  var hoja = ss.getSheetByName(nombre);
  var cols = ESQUEMA[nombre];
  if (!hoja) {
    if (!cols) throw new Error('Hoja no existe: ' + nombre);
    hoja = ss.insertSheet(nombre);
    hoja.getRange(1, 1, 1, cols.length).setValues([cols])
        .setFontWeight('bold').setBackground('#0B2545').setFontColor('#fff');
    hoja.setFrozenRows(1);
    return hoja;
  }
  if (cols) {
    var lastCol = hoja.getLastColumn();
    var actuales = lastCol ? hoja.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
    var faltan = cols.filter(function(c) { return actuales.indexOf(c) === -1; });
    if (faltan.length) {
      hoja.getRange(1, lastCol + 1, 1, faltan.length).setValues([faltan])
          .setFontWeight('bold').setBackground('#0B2545').setFontColor('#fff');
    }
  }
  return hoja;
}

function insertar(ss, nombre, fila) {
  var hoja = hojaParaEscribir(ss, nombre);
  var cab = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  hoja.appendRow(cab.map(function(h) {
    return fila[h] !== undefined ? String(fila[h]) : '';
  }));
}

function actualizar(ss, nombre, id, fila) {
  var hoja = hojaParaEscribir(ss, nombre);
  var datos = hoja.getDataRange().getValues();
  var cab   = datos[0];
  var colId = cab.indexOf('id');
  if (colId === -1) throw new Error(nombre + ' sin columna id');
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][colId]) === String(id)) {
      hoja.getRange(i + 1, 1, 1, cab.length).setValues([
        cab.map(function(h, j) {
          return fila[h] !== undefined
            ? String(fila[h])
            : String(datos[i][j] != null ? datos[i][j] : '');
        })
      ]);
      return;
    }
  }
  throw new Error('No encontrado id=' + id + ' en ' + nombre);
}

function eliminar(ss, nombre, id) {
  var hoja = ss.getSheetByName(nombre);
  if (!hoja) throw new Error('Hoja no existe: ' + nombre);
  var datos = hoja.getDataRange().getValues();
  var colId = datos[0].indexOf('id');
  if (colId === -1) throw new Error(nombre + ' sin columna id');
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][colId]) === String(id)) {
      hoja.deleteRow(i + 1);
      return;
    }
  }
}

// ════════════════════════════════════════════════════════════════
// RESUMEN MENSUAL DE COLABORADOR
// Devuelve: asignaciones del mes, total a pagar, estado de pago
// ════════════════════════════════════════════════════════════════
function resumenColaborador(ss, colaboradorId, mes) {
  var asignaciones = leer(ss, 'Asignaciones');
  var filtradas = asignaciones.filter(function(a) {
    var matchCol = !colaboradorId || a.colaboradorId === colaboradorId;
    var matchMes = !mes || a.mes === mes || (a.fechaAsignacion && a.fechaAsignacion.slice(0, 7) === mes);
    return matchCol && matchMes;
  });

  var totalAsignado  = filtradas.reduce(function(s, a) { return s + parseFloat(a.valorAsignado || 0); }, 0);
  var totalPagado    = filtradas.filter(function(a) { return a.estado === 'pagado'; })
                                .reduce(function(s, a) { return s + parseFloat(a.valorAsignado || 0); }, 0);
  var totalPendiente = totalAsignado - totalPagado;

  return {
    ok: true,
    colaboradorId: colaboradorId,
    mes: mes,
    asignaciones: filtradas,
    totalAsignado: totalAsignado,
    totalPagado: totalPagado,
    totalPendiente: totalPendiente,
    cantidadTrabajos: filtradas.length
  };
}

// ════════════════════════════════════════════════════════════════
// repararHojas — NUNCA borra datos
// Solo crea hojas faltantes y agrega columnas nuevas al final
// ════════════════════════════════════════════════════════════════
function repararHojas(ss) {
  var resultado = [];

  Object.keys(ESQUEMA).forEach(function(nombre) {
    var colsEsperadas = ESQUEMA[nombre];
    var hoja = ss.getSheetByName(nombre);

    // Crear hoja si no existe
    if (!hoja) {
      hoja = ss.insertSheet(nombre);
      hoja.getRange(1, 1, 1, colsEsperadas.length)
          .setValues([colsEsperadas])
          .setFontWeight('bold')
          .setBackground('#0B2545')
          .setFontColor('#fff');
      hoja.setFrozenRows(1);
      resultado.push('CREADA: ' + nombre);
      return;
    }

    // Hoja vacía — solo poner headers
    var lastCol = hoja.getLastColumn();
    if (lastCol === 0) {
      hoja.getRange(1, 1, 1, colsEsperadas.length)
          .setValues([colsEsperadas])
          .setFontWeight('bold')
          .setBackground('#0B2545')
          .setFontColor('#fff');
      hoja.setFrozenRows(1);
      resultado.push('HEADERS AÑADIDOS: ' + nombre);
      return;
    }

    var colsActuales = hoja.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

    // Agregar solo columnas que faltan (nunca modificar existentes)
    var faltantes = colsEsperadas.filter(function(c) {
      return colsActuales.indexOf(c) === -1;
    });

    if (faltantes.length === 0) {
      resultado.push('OK: ' + nombre + ' (' + colsActuales.length + ' cols)');
      return;
    }

    faltantes.forEach(function(colNombre) {
      var nuevaCol = hoja.getLastColumn() + 1;
      hoja.getRange(1, nuevaCol)
          .setValue(colNombre)
          .setFontWeight('bold')
          .setBackground('#0B2545')
          .setFontColor('#fff');
    });

    resultado.push('COLUMNAS AGREGADAS a ' + nombre + ': ' + faltantes.join(', '));
  });

  return resultado;
}

// ── inicializar ───────────────────────────────────────────────────
function inicializar() {
  var ss  = SpreadsheetApp.openById(SS_ID);
  var res = repararHojas(ss);
  Logger.log(res.join('\n'));
  Logger.log('✅ v7.1 — datos existentes preservados');
}

// ════════════════════════════════════════════════════════════════
// AUDITORÍA DE CONTRATOS
// • Cada contrato emitido recibe un código único y su huella SHA-256
//   (calculada aquí, en el servidor, sobre los bytes exactos del archivo).
// • Cada evento (emisión, firmas, anulación) se encadena con el anterior
//   mediante HMAC-SHA256 con una clave secreta guardada en las propiedades
//   del script: si alguien edita o borra una fila, la verificación lo detecta.
// • La hora de cada evento es la del servidor de Google (no la del equipo).
// ════════════════════════════════════════════════════════════════
function hex_(bytes) {
  return bytes.map(function(b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}
function sha256Hex_(bytes) {
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes));
}
function claveAuditoria_() {
  var props = PropertiesService.getScriptProperties();
  var k = props.getProperty('AUDIT_KEY');
  if (!k) { k = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('AUDIT_KEY', k); }
  return k;
}
function hmacEvento_(ev) {
  var base = [ev.hashPrevio, ev.contratoId, ev.fecha, ev.evento, ev.actor, ev.detalle, ev.hashDocumento].join('|');
  return hex_(Utilities.computeHmacSha256Signature(base, claveAuditoria_()));
}
function ahoraISO_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
}
function validarAdmin_(ss, auth) {
  auth = auth || {};
  var u = String(auth.usuario || ''), p = String(auth.password || '');
  if (!u || !p) throw new Error('Se requiere un usuario administrador');
  var usuarios = leer(ss, 'Usuarios');
  var ok;
  if (usuarios.length) {
    ok = usuarios.some(function(x) {
      return x.usuario === u && x.password === p && String(x.activo) !== '0' && (x.role || 'admin') === 'admin';
    });
  } else {
    ok = ADMINS_POR_DEFECTO.some(function(x) { return x[0] === u && x[1] === p; });
  }
  if (!ok) throw new Error('Usuario o contraseña de administrador no válidos');
  return u;
}
function filaContrato_(ss, codigo) {
  var lista = leer(ss, 'Contratos');
  for (var i = 0; i < lista.length; i++) if (lista[i].id === codigo) return lista[i];
  return null;
}
// Añade un evento encadenado al historial del contrato
function registrarEvento_(ss, contratoId, evento, actor, detalle, hashDocumento) {
  var hoja = hojaParaEscribir(ss, 'AuditoriaContratos');
  var previos = leer(ss, 'AuditoriaContratos').filter(function(e) { return e.contratoId === contratoId; });
  var ev = {
    id: Utilities.getUuid(), contratoId: contratoId, fecha: ahoraISO_(), evento: evento,
    actor: actor || '', detalle: detalle || '', hashDocumento: hashDocumento || '',
    hashPrevio: previos.length ? previos[previos.length - 1].hashEvento : 'INICIO'
  };
  ev.hashEvento = hmacEvento_(ev);
  // Texto plano ("'") para que Sheets no convierta fechas ni números
  var cab = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  hoja.appendRow(cab.map(function(h) { return ev[h] !== undefined ? "'" + String(ev[h]) : ''; }));
  return ev;
}
function guardarArchivo_(carpeta, nombre, base64, mime) {
  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mime, String(nombre).replace(/[\\\/:*?"<>|]/g, '_'));
  var file = carpetaContratos(carpeta).createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { file: file, hash: sha256Hex_(bytes), bytes: bytes };
}
function enviarComprobante_(para, asunto, html) {
  if (!para) return;
  try { MailApp.sendEmail({ to: para, subject: asunto, htmlBody: html, name: 'Research Assessor' }); } catch (e) {}
}
function htmlComprobante_(titulo, filas, urlVerificar) {
  return '<div style="font-family:Arial,sans-serif;max-width:560px;color:#0B2545">'
    + '<h2 style="color:#0F4C81;margin:0 0 12px">' + titulo + '</h2>'
    + '<table style="border-collapse:collapse;width:100%;font-size:13px">'
    + filas.map(function(f) { return '<tr><td style="padding:6px 8px;border:1px solid #E6E3DC;background:#FAFAF7;font-weight:bold;width:38%">' + f[0] + '</td><td style="padding:6px 8px;border:1px solid #E6E3DC;word-break:break-all">' + f[1] + '</td></tr>'; }).join('')
    + '</table>'
    + (urlVerificar ? '<p style="font-size:13px">Verifique la autenticidad del documento en:<br><a href="' + urlVerificar + '">' + urlVerificar + '</a></p>' : '')
    + '<p style="font-size:11px;color:#8C94A3">Guarde este correo: es un comprobante independiente del registro del contrato (código, huella SHA-256 y hora oficial).</p></div>';
}

// Emite (registra) la versión revisada del contrato, con su código y QR ya incluidos en el archivo
function emitirContrato(ss, b) {
  try {
    var actor = validarAdmin_(ss, b.auth);
    var codigo = String(b.codigo || '');
    if (!/^RA-\d{4}-[A-Z0-9]{8}$/.test(codigo)) throw new Error('Código de verificación no válido');
    if (filaContrato_(ss, codigo)) throw new Error('El código ya existe; vuelve a intentarlo');
    // v7.5: el contrato llega como PDF ya firmado con el certificado .p12 del prestador
    var esPdf = b.formato === 'pdf';
    if (esPdf) {
      var texto = Utilities.newBlob(Utilities.base64Decode(b.base64)).getDataAsString('ISO-8859-1');
      if (texto.slice(0, 5) !== '%PDF-') throw new Error('El archivo no es un PDF');
      if (b.firmado && !/\/ByteRange\s*\[/.test(texto)) throw new Error('El PDF no contiene la firma electrónica');
    }
    var r = esPdf
      ? guardarArchivo_(CARPETA_FIRMADOS, b.nombre || (codigo + '_firmado.pdf'), b.base64, 'application/pdf')
      : guardarArchivo_(CARPETA_EMITIDOS, b.nombre || (codigo + '.docx'), b.base64, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    var firmado = esPdf && !!b.firmado;
    var resumen = b.resumen || {};
    var anteriores = leer(ss, 'Contratos').filter(function(c) {
      return c.trabajoId === b.trabajoId && c.estado !== 'anulado' && c.estado !== 'reemplazado';
    });
    var fecha = ahoraISO_();
    var fila = {
      id: codigo, trabajoId: b.trabajoId || '', clienteId: b.clienteId || '', clienteNombre: resumen.cliente ? resumen.cliente.nombre : '',
      titulo: resumen.titulo || '', tipo: resumen.tipo || '', valor: resumen.valor || '', estado: firmado ? 'firmado_prestador' : 'emitido',
      version: String(leer(ss, 'Contratos').filter(function(c) { return c.trabajoId === b.trabajoId; }).length + 1),
      fechaEmision: fecha, hashEmitido: r.hash, urlEmitido: r.file.getUrl(), resumen: JSON.stringify(resumen), createdAt: fecha
    };
    if (firmado) {
      fila.fechaFirmaPrestador = fecha; fila.hashFirmaPrestador = r.hash; fila.urlFirmaPrestador = r.file.getUrl();
      fila.firmasPrestador = String(b.firmas || '');
    }
    insertarInterno_(ss, 'Contratos', fila);
    registrarEvento_(ss, codigo, 'EMITIDO', actor, 'Versión ' + fila.version + ' revisada y emitida' + (anteriores.length ? ' (reemplaza a ' + anteriores.map(function(a) { return a.id; }).join(', ') + ')' : ''), r.hash);
    if (firmado) registrarEvento_(ss, codigo, 'FIRMADO_PRESTADOR', actor, 'Firmado electrónicamente en el sistema con certificado .p12' + (b.resumenFirmas ? ' · ' + b.resumenFirmas : ''), r.hash);
    anteriores.forEach(function(a) {
      actualizarInterno_(ss, 'Contratos', a.id, { estado: 'reemplazado', reemplazadoPor: codigo });
      registrarEvento_(ss, a.id, 'REEMPLAZADO', actor, 'Reemplazado por ' + codigo, '');
    });
    if (b.trabajoId) actualizar(ss, 'Trabajos', b.trabajoId, { contratoUrl: r.file.getUrl(), contratoFecha: fecha.slice(0, 10) });
    if (b.notificar) {
      var html = htmlComprobante_(firmado ? 'Contrato emitido y firmado por el prestador' : 'Contrato emitido', [['Código', codigo], ['Contrato', fila.titulo], ['Cliente', fila.clienteNombre],
        ['Fecha (servidor)', fecha], ['Huella SHA-256', r.hash]].concat(firmado ? [['Firma', b.resumenFirmas || 'Firma electrónica']] : []), b.urlVerificar);
      enviarComprobante_(b.emailCliente, 'Contrato ' + codigo + (firmado ? ' firmado' : ' emitido') + ' — Research Assessor', html);
      enviarComprobante_(Session.getEffectiveUser().getEmail(), '[Copia] Contrato ' + codigo + (firmado ? ' firmado' : ' emitido'), html);
    }
    return { ok: true, codigo: codigo, hash: r.hash, url: r.file.getUrl(), fecha: fecha, version: fila.version };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// Registra el PDF firmado electrónicamente (por el prestador o por ambas partes)
function registrarFirmado(ss, b) {
  try {
    var actor = validarAdmin_(ss, b.auth);
    var c = filaContrato_(ss, b.codigo);
    if (!c) throw new Error('Contrato no encontrado: ' + b.codigo);
    if (c.estado === 'anulado' || c.estado === 'reemplazado') throw new Error('El contrato está ' + c.estado + '; emite una nueva versión');
    var etapa = b.etapa === 'ambos' ? 'ambos' : 'prestador';
    var bytes = Utilities.base64Decode(b.base64);
    var texto = Utilities.newBlob(bytes).getDataAsString('ISO-8859-1');
    if (texto.slice(0, 5) !== '%PDF-') throw new Error('El archivo no es un PDF');
    var tieneFirma = /\/ByteRange\s*\[/.test(texto);
    if (etapa === 'prestador' && !tieneFirma) throw new Error('El PDF no contiene una firma electrónica. Fírmalo con un certificado de firma electrónica y vuelve a subirlo.');
    var r = guardarArchivo_(CARPETA_FIRMADOS, b.nombre || (b.codigo + '_firmado.pdf'), b.base64, 'application/pdf');
    var fecha = ahoraISO_();
    var firmas = String(b.firmas || '');
    var cambios = etapa === 'ambos'
      ? { estado: 'firmado', fechaFirmaAmbos: fecha, hashFirmaAmbos: r.hash, urlFirmaAmbos: r.file.getUrl(), firmasAmbos: firmas }
      : { estado: 'firmado_prestador', fechaFirmaPrestador: fecha, hashFirmaPrestador: r.hash, urlFirmaPrestador: r.file.getUrl(), firmasPrestador: firmas };
    actualizarInterno_(ss, 'Contratos', c.id, cambios);
    // El portal del cliente muestra siempre la última versión firmada
    if (c.trabajoId) { try { actualizar(ss, 'Trabajos', c.trabajoId, { contratoUrl: r.file.getUrl(), contratoFecha: fecha.slice(0, 10) }); } catch (e) {} }
    var detalle = (etapa === 'ambos' ? 'PDF firmado por ambas partes' : 'PDF firmado electrónicamente por el prestador')
      + (tieneFirma ? '' : ' (sin firma digital: firma manuscrita escaneada)') + (b.resumenFirmas ? ' · ' + b.resumenFirmas : '');
    registrarEvento_(ss, c.id, etapa === 'ambos' ? 'FIRMADO_AMBAS_PARTES' : 'FIRMADO_PRESTADOR', actor, detalle, r.hash);
    if (b.notificar) {
      var html = htmlComprobante_(etapa === 'ambos' ? 'Contrato firmado por ambas partes' : 'Contrato firmado por el prestador',
        [['Código', c.id], ['Contrato', c.titulo], ['Fecha (servidor)', fecha], ['Huella SHA-256 del PDF', r.hash], ['Firmas', b.resumenFirmas || (tieneFirma ? 'Firma digital' : 'Manuscrita')]], b.urlVerificar);
      enviarComprobante_(b.emailCliente, 'Contrato ' + c.id + (etapa === 'ambos' ? ' firmado' : ' firmado por el prestador') + ' — Research Assessor', html);
      enviarComprobante_(Session.getEffectiveUser().getEmail(), '[Copia] Contrato ' + c.id + ' firmado (' + etapa + ')', html);
    }
    return { ok: true, hash: r.hash, url: r.file.getUrl(), fecha: fecha, tieneFirma: tieneFirma };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

function anularContrato(ss, b) {
  try {
    var actor = validarAdmin_(ss, b.auth);
    var c = filaContrato_(ss, b.codigo);
    if (!c) throw new Error('Contrato no encontrado');
    if (c.estado === 'anulado') return { ok: true };
    actualizarInterno_(ss, 'Contratos', c.id, { estado: 'anulado' });
    registrarEvento_(ss, c.id, 'ANULADO', actor, String(b.motivo || 'Sin motivo indicado').slice(0, 300), '');
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// Consulta pública (código QR): datos del registro y validación de la cadena de auditoría
function enmascarar_(ced) {
  ced = String(ced || '');
  return ced.length > 4 ? ced.slice(0, 2) + new Array(ced.length - 3).join('•') + ced.slice(-2) : ced;
}
// Iniciales de una persona: "María José Pérez" → "M. J. P."
function iniciales_(nombre) {
  return String(nombre || '').trim().split(/\s+/).filter(Boolean).map(function(p) { return p.charAt(0).toUpperCase() + '.'; }).join(' ');
}
// Cadena de auditoría del contrato: eventos en orden y si están íntegros (HMAC encadenado)
function cadenaContrato_(ss, c) {
  var eventos = leer(ss, 'AuditoriaContratos').filter(function(e) { return e.contratoId === c.id; });
  var integra = eventos.length > 0, previo = 'INICIO';
  eventos.forEach(function(e) {
    if (e.hashPrevio !== previo || hmacEvento_(e) !== e.hashEvento) integra = false;
    previo = e.hashEvento;
  });
  // La última huella registrada debe coincidir con el estado actual del contrato
  var ultimoDoc = eventos.filter(function(e) { return e.hashDocumento; }).pop();
  var hashActual = c.hashFirmaAmbos || c.hashFirmaPrestador || c.hashEmitido;
  if (ultimoDoc && ultimoDoc.hashDocumento !== hashActual) integra = false;
  return { integra: integra, eventos: eventos };
}
function _firmasPublicas_(json, prestadorNombre) {
  var fs = [];
  try { fs = JSON.parse(json || '[]') || []; } catch (e) {}
  return fs.map(function(f, i) {
    // La firma del prestador se muestra con su nombre; la del cliente solo con iniciales
    var esPrestador = i === 0 || (prestadorNombre && String(f.nombre || '').toUpperCase() === String(prestadorNombre).toUpperCase());
    return { orden: f.orden || (i + 1), parte: esPrestador ? 'prestador' : 'cliente',
      nombre: esPrestador ? (f.nombre || '') : iniciales_(f.nombre), emisor: f.emisor || '', fecha: f.fecha || '', valida: !!f.valida };
  });
}

// Consulta PÚBLICA (código QR): solo confirma que el contrato existe, su estado, quién firmó
// y las huellas para comprobar una copia. NO expone valores, plan de pagos, títulos, detalles
// del servicio, cédulas ni enlaces a los documentos (eso es solo para las partes).
function verificarContrato(ss, codigo) {
  var c = filaContrato_(ss, String(codigo || '').trim().toUpperCase());
  if (!c) return { ok: true, encontrado: false };
  var cad = cadenaContrato_(ss, c);
  var res = {};
  try { res = JSON.parse(c.resumen || '{}'); } catch (e) {}
  var prNombre = res.prestador && res.prestador.nombre || '';
  return {
    ok: true, encontrado: true, publico: true, cadenaIntegra: cad.integra,
    contrato: {
      codigo: c.id, estado: c.estado, version: c.version, reemplazadoPor: c.reemplazadoPor,
      fechaEmision: c.fechaEmision, hashEmitido: c.hashEmitido,
      fechaFirmaPrestador: c.fechaFirmaPrestador, hashFirmaPrestador: c.hashFirmaPrestador,
      fechaFirmaAmbos: c.fechaFirmaAmbos, hashFirmaAmbos: c.hashFirmaAmbos,
      firmasPrestador: _firmasPublicas_(c.firmasPrestador, prNombre),
      firmasAmbos: _firmasPublicas_(c.firmasAmbos, prNombre),
      partes: { prestador: prNombre, cliente: iniciales_(res.cliente && res.cliente.nombre) }
    },
    eventos: cad.eventos.map(function(e) { return { fecha: e.fecha, evento: e.evento, hashDocumento: e.hashDocumento }; })
  };
}

// Consulta COMPLETA del registro (solo administradores): incluye detalles y enlaces
function historialContrato(ss, b) {
  try {
    validarAdmin_(ss, b.auth);
    var c = filaContrato_(ss, String(b.codigo || '').trim().toUpperCase());
    if (!c) return { ok: true, encontrado: false };
    return historialCompleto_(ss, c);
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}
function historialCompleto_(ss, c) {
  var cad = cadenaContrato_(ss, c), integra = cad.integra, eventos = cad.eventos;
  var res = {};
  try { res = JSON.parse(c.resumen || '{}'); } catch (e) {}
  if (res.cliente) res.cliente.cedula = enmascarar_(res.cliente.cedula);
  if (res.prestador) res.prestador.cedula = enmascarar_(res.prestador.cedula);
  return {
    ok: true, encontrado: true, cadenaIntegra: integra,
    contrato: {
      codigo: c.id, estado: c.estado, version: c.version, titulo: c.titulo, tipo: c.tipo, valor: c.valor,
      fechaEmision: c.fechaEmision, hashEmitido: c.hashEmitido, urlEmitido: c.urlEmitido,
      fechaFirmaPrestador: c.fechaFirmaPrestador, hashFirmaPrestador: c.hashFirmaPrestador, urlFirmaPrestador: c.urlFirmaPrestador, firmasPrestador: c.firmasPrestador,
      fechaFirmaAmbos: c.fechaFirmaAmbos, hashFirmaAmbos: c.hashFirmaAmbos, urlFirmaAmbos: c.urlFirmaAmbos, firmasAmbos: c.firmasAmbos,
      reemplazadoPor: c.reemplazadoPor, resumen: res
    },
    eventos: eventos.map(function(e) { return { fecha: e.fecha, evento: e.evento, detalle: e.detalle, hashDocumento: e.hashDocumento }; })
  };
}
// Escritura para hojas protegidas (solo desde las acciones de auditoría).
// Todo se guarda como texto ("'") para que Sheets no convierta huellas ni fechas.
function insertarInterno_(ss, nombre, fila) {
  var hoja = hojaParaEscribir(ss, nombre);
  var cab = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  hoja.appendRow(cab.map(function(h) { return fila[h] !== undefined ? "'" + String(fila[h]) : ''; }));
}
function actualizarInterno_(ss, nombre, id, fila) {
  var hoja = hojaParaEscribir(ss, nombre);
  var datos = hoja.getDataRange().getValues();
  var cab = datos[0], colId = cab.indexOf('id');
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][colId]) === String(id)) {
      hoja.getRange(i + 1, 1, 1, cab.length).setValues([cab.map(function(h, j) {
        var v = fila[h] !== undefined ? fila[h] : datos[i][j];
        return v === '' || v === null ? '' : "'" + String(v);
      })]);
      return;
    }
  }
  throw new Error('No encontrado id=' + id + ' en ' + nombre);
}

// ════════════════════════════════════════════════════════════════
// CONTRATOS EN GOOGLE DRIVE
// Recibe el .docx en base64, lo guarda en la carpeta CARPETA_CONTRATOS
// y lo comparte como "cualquiera con el enlace puede ver".
// La primera vez, Apps Script pedirá autorizar el acceso a Google Drive.
// ════════════════════════════════════════════════════════════════
function carpetaContratos(nombre) {
  nombre = nombre || CARPETA_CONTRATOS;
  var it = DriveApp.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : DriveApp.createFolder(nombre);
}

function guardarContrato(ss, b) {
  try {
    if (!b.base64) throw new Error('Contrato vacío');
    var nombre = String(b.nombre || 'Contrato.docx').replace(/[\\\/:*?"<>|]/g, '_');
    var blob = Utilities.newBlob(
      Utilities.base64Decode(b.base64),
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      nombre
    );
    var carpeta = String(b.carpeta || '').indexOf('Research Assessor') === 0 ? b.carpeta : CARPETA_CONTRATOS;
    var file = carpetaContratos(carpeta).createFile(blob);
    var aviso = '';
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (se) {
      aviso = 'No se pudo compartir públicamente: ' + se.toString();
    }
    var url = file.getUrl();
    var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (b.trabajoId) {
      actualizar(ss, 'Trabajos', b.trabajoId, { contratoUrl: url, contratoFecha: fecha });
    }
    if (b.propuestaId) {
      actualizar(ss, 'Propuestas', b.propuestaId, { docUrl: url, docFecha: fecha });
    }
    return { ok: true, url: url, fileId: file.getId(), fecha: fecha, aviso: aviso };
  } catch (err) {
    return { ok: false, error: 'Error guardando contrato: ' + err.toString() };
  }
}

// ════════════════════════════════════════════════════════════════
// Google Meet via Calendar API v3
// ════════════════════════════════════════════════════════════════
function crearGoogleMeet(title, dateStr, timeStr, duration) {
  try {
    var hh = 10, mm = 0;
    if (timeStr) {
      var parts = timeStr.split(':');
      hh = parseInt(parts[0]) || 10;
      mm = parseInt(parts[1]) || 0;
    }
    var startDate = new Date();
    if (dateStr && dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
      var dp = dateStr.split('-');
      startDate = new Date(parseInt(dp[0]), parseInt(dp[1]) - 1, parseInt(dp[2]), hh, mm, 0);
    } else {
      startDate.setHours(hh, mm, 0, 0);
    }
    var endDate = new Date(startDate.getTime() + (duration || 60) * 60000);

    var calId    = CalendarApp.getDefaultCalendar().getId();
    var resource = {
      summary: title,
      start:   { dateTime: startDate.toISOString(), timeZone: Session.getScriptTimeZone() },
      end:     { dateTime: endDate.toISOString(),   timeZone: Session.getScriptTimeZone() },
      conferenceData: {
        createRequest: {
          requestId: 'ra-' + Date.now(),
          conferenceSolutionKey: { type: 'hangoutsMeet' }
        }
      },
      description: 'Reunión creada desde Research Assessor'
    };

    var created  = Calendar.Events.insert(resource, calId, { conferenceDataVersion: 1 });
    var meetLink = '';
    if (created.conferenceData && created.conferenceData.entryPoints) {
      for (var i = 0; i < created.conferenceData.entryPoints.length; i++) {
        if (created.conferenceData.entryPoints[i].entryPointType === 'video') {
          meetLink = created.conferenceData.entryPoints[i].uri;
          break;
        }
      }
    }
    if (!meetLink && created.hangoutLink) meetLink = created.hangoutLink;

    return meetLink
      ? { ok: true, meetLink: meetLink }
      : { ok: false, error: 'Evento creado sin enlace Meet. Revisa Google Calendar.' };

  } catch(err) {
    return { ok: false, error: 'Error Meet: ' + err.toString() +
      '. Verifica Google Calendar API en Servicios Avanzados.' };
  }
}

function respJson(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
