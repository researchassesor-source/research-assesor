// ================================================================
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
  Propuestas:     ['id','clienteId','clienteNombre','titulo','area','fecha','validez','estado','articulos','hitos','notas','docUrl','docFecha','createdAt']
};

var CARPETA_CONTRATOS = 'Research Assessor — Contratos';

// ── Estado de Asignaciones ────────────────────────────────────────
// pendiente  → asignado pero no completado
// completado → el colaborador terminó su parte
// pagado     → el admin ya pagó al colaborador

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
        propuestas:      leer(ss, 'Propuestas')
      };
    } else if (action === 'write') {
      var body = JSON.parse(e.parameter.payload || '{}');
      if      (body.action === 'insertar')   { insertar(ss, body.tabla, body.fila); }
      else if (body.action === 'actualizar') { actualizar(ss, body.tabla, body.id, body.fila); }
      else if (body.action === 'eliminar')   { eliminar(ss, body.tabla, body.id); }
      else { throw new Error('Op desconocida: ' + body.action); }
      result = { ok: true };
    } else if (action === 'ping') {
      result = { ok: true, msg: 'OK', script: 'v7.1' };
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
