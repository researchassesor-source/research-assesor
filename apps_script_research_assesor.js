// ============================================================
//  RESEARCH ASSESOR — Google Apps Script
//  Copia este código completo en:
//  Google Sheets → Extensiones → Apps Script
// ============================================================

// ── CONFIGURACIÓN ──────────────────────────────────────────
const CONFIG = {
  HOJA_TRABAJOS:  'Trabajos',
  HOJA_PAGOS:     'Pagos',
  HOJA_SESIONES:  'Sesiones Zoom',
  HOJA_CLIENTES:  'Clientes',
  COLOR_HEADER:   '#1a6fb5',
  COLOR_HEADER_TXT: '#ffffff',
};

// ── MENÚ PERSONALIZADO ─────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Research Assesor')
    .addItem('✅ Crear / Resetear hojas', 'crearHojas')
    .addSeparator()
    .addItem('📊 Ver resumen del mes', 'mostrarResumenMes')
    .addItem('⚠️  Ver pagos vencidos', 'mostrarPagosVencidos')
    .addItem('📅 Ver sesiones de hoy', 'mostrarSesionesHoy')
    .addSeparator()
    .addItem('📱 Enviar recordatorios WhatsApp (lista)', 'listarRecordatorios')
    .addItem('🔄 Actualizar estados automáticamente', 'actualizarEstados')
    .addToUi();
}

// ── CREAR HOJAS CON ENCABEZADOS ────────────────────────────
function crearHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Hoja: TRABAJOS
  crearHoja(ss, CONFIG.HOJA_TRABAJOS, [
    'ID', 'Cliente', 'Teléfono', 'Email',
    'Tipo de Servicio', 'Título del Trabajo',
    'Valor Total (USD)', 'Total Pagado (USD)', 'Saldo Pendiente (USD)', '% Avance',
    'Estado', 'Fecha Inicio', 'Fecha Límite', 'Notas'
  ], [30, 140, 120, 160, 160, 220, 100, 100, 120, 70, 120, 100, 100, 200]);

  // Hoja: PAGOS
  crearHoja(ss, CONFIG.HOJA_PAGOS, [
    'ID Pago', 'ID Trabajo', 'Cliente', 'Tipo de Servicio',
    'Fecha Pago', 'Monto (USD)', 'Método de Pago', 'Nota / Referencia'
  ], [60, 70, 140, 160, 100, 100, 120, 200]);

  // Hoja: SESIONES ZOOM
  crearHoja(ss, CONFIG.HOJA_SESIONES, [
    'ID', 'Cliente', 'Teléfono', 'Tipo de Sesión',
    'Fecha', 'Hora', 'Duración (min)', 'ID Meeting Zoom',
    'Link Zoom', 'Estado', 'ID Trabajo Relacionado', 'Notas'
  ], [30, 140, 120, 160, 90, 70, 90, 140, 260, 110, 120, 200]);

  // Hoja: CLIENTES
  crearHoja(ss, CONFIG.HOJA_CLIENTES, [
    'ID', 'Nombre Completo', 'Teléfono', 'Email',
    'Ciudad', 'Tipo de Cliente', 'Fecha Registro', 'Notas'
  ], [30, 160, 120, 180, 100, 130, 110, 200]);

  // Hoja: RESUMEN (dashboard)
  crearHojaResumen(ss);

  SpreadsheetApp.getUi().alert(
    '✅ Hojas creadas correctamente.\n\n' +
    'Puedes empezar a ingresar datos en:\n' +
    '• Trabajos\n• Pagos\n• Sesiones Zoom\n• Clientes\n\n' +
    'El Resumen se actualiza automáticamente.'
  );
}

function crearHoja(ss, nombre, headers, anchos) {
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
  } else {
    // Si ya existe, solo actualiza encabezados sin borrar datos
    const respuesta = SpreadsheetApp.getUi().alert(
      `La hoja "${nombre}" ya existe.\n¿Deseas resetear solo los encabezados (sin borrar datos)?`,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (respuesta !== SpreadsheetApp.getUi().Button.YES) return;
  }

  // Encabezado
  const rango = hoja.getRange(1, 1, 1, headers.length);
  rango.setValues([headers]);
  rango.setBackground(CONFIG.COLOR_HEADER);
  rango.setFontColor(CONFIG.COLOR_HEADER_TXT);
  rango.setFontWeight('bold');
  rango.setFontSize(11);
  hoja.setFrozenRows(1);

  // Anchos de columna
  anchos.forEach((w, i) => hoja.setColumnWidth(i + 1, w));

  // Formato especial por hoja
  if (nombre === CONFIG.HOJA_TRABAJOS) {
    aplicarFormatoTrabajos(hoja);
  } else if (nombre === CONFIG.HOJA_PAGOS) {
    aplicarFormatoPagos(hoja);
  } else if (nombre === CONFIG.HOJA_SESIONES) {
    aplicarFormatoSesiones(hoja);
  }

  hoja.setTabColor(CONFIG.COLOR_HEADER);
}

function aplicarFormatoTrabajos(hoja) {
  // Validación de Estado (col 11)
  const reglaEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(['En proceso', 'Pendiente pago', 'Completado', 'Cancelado'], true)
    .build();
  hoja.getRange('K2:K1000').setDataValidation(reglaEstado);

  // Formato moneda cols G, H, I
  hoja.getRange('G2:I1000').setNumberFormat('"$"#,##0.00');

  // Formato porcentaje col J
  hoja.getRange('J2:J1000').setNumberFormat('0"%"');

  // Formato fecha cols L, M
  hoja.getRange('L2:M1000').setNumberFormat('dd/mm/yyyy');

  // Color condicional: Completado = verde, Pendiente = amarillo, Cancelado = rojo
  const reglasColor = [
    { valor: 'Completado',     fondo: '#d9ead3', texto: '#274e13' },
    { valor: 'En proceso',     fondo: '#dde5f7', texto: '#1c3a6e' },
    { valor: 'Pendiente pago', fondo: '#fff2cc', texto: '#7d4e00' },
    { valor: 'Cancelado',      fondo: '#fce5cd', texto: '#7f0000' },
  ];

  reglasColor.forEach(r => {
    const regla = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(r.valor)
      .setBackground(r.fondo)
      .setFontColor(r.texto)
      .setRanges([hoja.getRange('K2:K1000')])
      .build();
    const reglas = hoja.getConditionalFormatRules();
    reglas.push(regla);
    hoja.setConditionalFormatRules(reglas);
  });
}

function aplicarFormatoPagos(hoja) {
  // Formato moneda col F
  hoja.getRange('F2:F1000').setNumberFormat('"$"#,##0.00');
  // Formato fecha col E
  hoja.getRange('E2:E1000').setNumberFormat('dd/mm/yyyy');
  // Validación método de pago col G
  const reglaPago = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Transferencia', 'Efectivo', 'PayPal', 'Nequi', 'Binance', 'Otro'], true)
    .build();
  hoja.getRange('G2:G1000').setDataValidation(reglaPago);
}

function aplicarFormatoSesiones(hoja) {
  // Formato fecha col E
  hoja.getRange('E2:E1000').setNumberFormat('dd/mm/yyyy');
  // Validación Estado col J
  const reglaEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Programada', 'Completada', 'Cancelada'], true)
    .build();
  hoja.getRange('J2:J1000').setDataValidation(reglaEstado);
  // Validación Tipo sesión col D
  const reglaTipo = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Asesoría inicial','Revisión de avance','Correcciones','Asesoría metodológica','Entrega de borradores','Entrega final','Seguimiento','Otro'], true)
    .build();
  hoja.getRange('D2:D1000').setDataValidation(reglaTipo);
  // Link zoom en azul col I
  hoja.getRange('I2:I1000').setFontColor('#1155cc');
}

// ── HOJA RESUMEN (dashboard automático) ────────────────────
function crearHojaResumen(ss) {
  let hoja = ss.getSheetByName('📊 Resumen');
  if (!hoja) hoja = ss.insertSheet('📊 Resumen');
  else hoja.clear();

  hoja.setTabColor('#1D9E75');

  const datos = [
    ['RESEARCH ASSESOR — Panel de Control', '', ''],
    ['Última actualización:', new Date(), ''],
    ['', '', ''],
    ['── TRABAJOS ─────────────────', '', ''],
    ['Total trabajos registrados',    `=COUNTA(Trabajos!B2:B)`,                                         ''],
    ['En proceso',                    `=COUNTIF(Trabajos!K2:K,"En proceso")`,                           ''],
    ['Pendiente pago',                `=COUNTIF(Trabajos!K2:K,"Pendiente pago")`,                       ''],
    ['Completados',                   `=COUNTIF(Trabajos!K2:K,"Completado")`,                           ''],
    ['Cancelados',                    `=COUNTIF(Trabajos!K2:K,"Cancelado")`,                            ''],
    ['', '', ''],
    ['── FINANZAS ────────────────', '', ''],
    ['Total facturado (USD)',          `=IFERROR(SUM(Trabajos!G2:G),0)`,                                 ''],
    ['Total cobrado (USD)',            `=IFERROR(SUM(Trabajos!H2:H),0)`,                                 ''],
    ['Total pendiente (USD)',          `=IFERROR(SUM(Trabajos!I2:I),0)`,                                 ''],
    ['% Cobrado',                     `=IFERROR(SUM(Trabajos!H2:H)/SUM(Trabajos!G2:G)*100,0)`,          '%'],
    ['', '', ''],
    ['── MES ACTUAL ─────────────', '', ''],
    ['Mes',                           `=TEXT(TODAY(),"MMMM YYYY")`,                                     ''],
    ['Pagos recibidos este mes',       `=IFERROR(SUMPRODUCT((MONTH(Pagos!E2:E1000)=MONTH(TODAY()))*(YEAR(Pagos!E2:E1000)=YEAR(TODAY()))*Pagos!F2:F1000),0)`, ''],
    ['Nro. pagos este mes',            `=IFERROR(SUMPRODUCT((MONTH(Pagos!E2:E1000)=MONTH(TODAY()))*(YEAR(Pagos!E2:E1000)=YEAR(TODAY()))*(Pagos!F2:F1000>0)),0)`, ''],
    ['', '', ''],
    ['── SESIONES ZOOM ──────────', '', ''],
    ['Total sesiones programadas',     `=COUNTIF('Sesiones Zoom'!J2:J,"Programada")`,                  ''],
    ['Sesiones completadas',           `=COUNTIF('Sesiones Zoom'!J2:J,"Completada")`,                  ''],
    ['Sesiones hoy',                   `=COUNTIF('Sesiones Zoom'!E2:E,TODAY())`,                       ''],
    ['', '', ''],
    ['── CLIENTES ───────────────', '', ''],
    ['Total clientes',                 `=COUNTA(Clientes!B2:B)`,                                       ''],
  ];

  hoja.getRange(1, 1, datos.length, 3).setValues(datos);

  // Formato título
  const titulo = hoja.getRange('A1');
  titulo.setFontSize(14).setFontWeight('bold').setFontColor(CONFIG.COLOR_HEADER);

  // Formato secciones
  [4,11,17,22,27].forEach(fila => {
    hoja.getRange(fila, 1).setFontWeight('bold').setFontColor('#555');
  });

  // Formato moneda
  [12,13,14].forEach(fila => {
    hoja.getRange(fila, 2).setNumberFormat('"$"#,##0.00');
  });
  hoja.getRange(15, 2).setNumberFormat('0.00"%"');
  hoja.getRange(19, 2).setNumberFormat('"$"#,##0.00');
  hoja.getRange(2, 2).setNumberFormat('dd/mm/yyyy hh:mm');

  hoja.setColumnWidth(1, 220);
  hoja.setColumnWidth(2, 160);
  hoja.setFrozenRows(1);
}

// ── FUNCIONES DE CONSULTA ──────────────────────────────────

function mostrarResumenMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hPagos = ss.getSheetByName(CONFIG.HOJA_PAGOS);
  if (!hPagos) { SpreadsheetApp.getUi().alert('Hoja de Pagos no encontrada. Ejecuta "Crear hojas" primero.'); return; }

  const datos = hPagos.getDataRange().getValues();
  const mesActual = new Date().getMonth();
  const anioActual = new Date().getFullYear();

  let totalMes = 0, countMes = 0;
  for (let i = 1; i < datos.length; i++) {
    if (!datos[i][4]) continue;
    const fecha = new Date(datos[i][4]);
    if (fecha.getMonth() === mesActual && fecha.getFullYear() === anioActual) {
      totalMes += parseFloat(datos[i][5]) || 0;
      countMes++;
    }
  }

  const mesNombre = new Date().toLocaleDateString('es-EC', {month:'long', year:'numeric'});
  SpreadsheetApp.getUi().alert(
    `📊 Resumen de ${mesNombre}\n\n` +
    `💰 Total cobrado: $${totalMes.toFixed(2)}\n` +
    `📝 Número de pagos: ${countMes}\n\n` +
    `Ver hoja "📊 Resumen" para más detalles.`
  );
}

function mostrarPagosVencidos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hTrabajos = ss.getSheetByName(CONFIG.HOJA_TRABAJOS);
  if (!hTrabajos) { SpreadsheetApp.getUi().alert('Hoja Trabajos no encontrada.'); return; }

  const datos = hTrabajos.getDataRange().getValues();
  const hoy = new Date();
  let lista = [];

  for (let i = 1; i < datos.length; i++) {
    if (!datos[i][1]) continue;
    const estado = datos[i][10];
    if (estado === 'Completado' || estado === 'Cancelado') continue;
    const pendiente = parseFloat(datos[i][8]) || 0;
    if (pendiente <= 0) continue;
    const fechaLimite = datos[i][12] ? new Date(datos[i][12]) : null;
    if (!fechaLimite) continue;
    const diasRestantes = Math.ceil((fechaLimite - hoy) / 86400000);
    if (diasRestantes <= 30) {
      lista.push(`• ${datos[i][1]} | $${pendiente.toFixed(2)} | Vence: ${fechaLimite.toLocaleDateString('es-EC')} (${diasRestantes >= 0 ? diasRestantes + ' días' : 'VENCIDO'})`);
    }
  }

  if (lista.length === 0) {
    SpreadsheetApp.getUi().alert('✅ No hay pagos próximos a vencer en los próximos 30 días.');
  } else {
    SpreadsheetApp.getUi().alert(`⚠️ Pagos próximos a vencer (30 días):\n\n${lista.join('\n')}`);
  }
}

function mostrarSesionesHoy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hSesiones = ss.getSheetByName(CONFIG.HOJA_SESIONES);
  if (!hSesiones) { SpreadsheetApp.getUi().alert('Hoja Sesiones no encontrada.'); return; }

  const datos = hSesiones.getDataRange().getValues();
  const hoy = new Date();
  const hoyStr = `${hoy.getDate()}/${hoy.getMonth()+1}/${hoy.getFullYear()}`;
  let lista = [];

  for (let i = 1; i < datos.length; i++) {
    if (!datos[i][1]) continue;
    if (datos[i][9] !== 'Programada') continue;
    const fecha = datos[i][4] ? new Date(datos[i][4]) : null;
    if (!fecha) continue;
    const fechaStr = `${fecha.getDate()}/${fecha.getMonth()+1}/${fecha.getFullYear()}`;
    if (fechaStr === hoyStr) {
      lista.push(`• ${datos[i][5]} — ${datos[i][1]}\n  Tipo: ${datos[i][3]}\n  Link: ${datos[i][8]}`);
    }
  }

  if (lista.length === 0) {
    SpreadsheetApp.getUi().alert('📅 No tienes sesiones Zoom programadas para hoy.');
  } else {
    SpreadsheetApp.getUi().alert(`🎥 Sesiones Zoom de hoy:\n\n${lista.join('\n\n')}`);
  }
}

function listarRecordatorios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hTrabajos = ss.getSheetByName(CONFIG.HOJA_TRABAJOS);
  if (!hTrabajos) return;

  const datos = hTrabajos.getDataRange().getValues();
  const hoy = new Date();
  let lista = [];

  for (let i = 1; i < datos.length; i++) {
    if (!datos[i][1]) continue;
    const estado = datos[i][10];
    if (estado === 'Completado' || estado === 'Cancelado') continue;
    const pendiente = parseFloat(datos[i][8]) || 0;
    const tel = datos[i][2] || '';
    if (pendiente <= 0 || !tel) continue;
    const fechaLimite = datos[i][12] ? new Date(datos[i][12]).toLocaleDateString('es-EC') : 'sin fecha';
    lista.push(`${datos[i][1]} | ${tel} | $${pendiente.toFixed(2)} | Vence: ${fechaLimite}`);
  }

  if (lista.length === 0) {
    SpreadsheetApp.getUi().alert('✅ No hay recordatorios pendientes.');
  } else {
    SpreadsheetApp.getUi().alert(
      `📱 Clientes con saldo pendiente (${lista.length}):\n\n` +
      lista.join('\n') +
      '\n\nCopia el número y envía el recordatorio desde la app.'
    );
  }
}

function actualizarEstados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hTrabajos = ss.getSheetByName(CONFIG.HOJA_TRABAJOS);
  if (!hTrabajos) return;

  const datos = hTrabajos.getDataRange().getValues();
  let actualizados = 0;

  for (let i = 1; i < datos.length; i++) {
    if (!datos[i][1]) continue;
    const total = parseFloat(datos[i][6]) || 0;
    const pagado = parseFloat(datos[i][7]) || 0;
    const estadoActual = datos[i][10];
    if (estadoActual === 'Cancelado') continue;

    let nuevoEstado = estadoActual;
    if (pagado >= total && total > 0) {
      nuevoEstado = 'Completado';
    } else if (pagado > 0 && pagado < total) {
      nuevoEstado = 'En proceso';
    } else if (pagado === 0) {
      nuevoEstado = 'Pendiente pago';
    }

    if (nuevoEstado !== estadoActual) {
      hTrabajos.getRange(i + 1, 11).setValue(nuevoEstado);
      actualizados++;
    }
  }

  SpreadsheetApp.getUi().alert(
    actualizados > 0
      ? `✅ Se actualizaron ${actualizados} estado(s) automáticamente.`
      : '✅ Todos los estados ya están correctos.'
  );
}

// ── TRIGGER AUTOMÁTICO ─────────────────────────────────────
// Ejecutar esto UNA VEZ para activar el resumen automático diario:
// function activarTriggerDiario() {
//   ScriptApp.newTrigger('actualizarEstados')
//     .timeBased().everyDays(1).atHour(8).create();
// }
