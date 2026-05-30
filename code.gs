// ============================================================
//  HoriViaje Ticket — Google Apps Script
//  Horisoes IPS
//  Deployment ID: AKfycbyzbAftjknN-xws2C8u4SE_SNYZB15XFlb3yT7_TtOiPzq8FufxqWkTI8jdhEZswV9p
// ============================================================

var SHEET_NAME = 'Tickets';

// ⚠️ SOLO necesario si el script es STANDALONE (no está dentro de la Sheet).
// Pega aquí el ID de tu Google Sheet (lo encuentras en la URL de la hoja).
// Ejemplo: https://docs.google.com/spreadsheets/d/  <--- ESTE ID --->/edit
// Si el script está dentro de la hoja (Extensiones > Apps Script), deja esto vacío.
var SPREADSHEET_ID = '113NR3E1RfLsJkVnC0y6jwL71WpKNsznnbf8EtyhMiz0';

// ------------------------------------------------------------
//  doPost: recibe el JSON enviado desde el formulario HTML
// ------------------------------------------------------------
function doPost(e) {
  try {
    var raw  = e.postData ? e.postData.contents : '{}';
    var data = JSON.parse(raw);

    // Obtener el spreadsheet (container-bound o standalone)
    var ss;
    if (SPREADSHEET_ID && SPREADSHEET_ID.length > 5) {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }

    if (!ss) {
      throw new Error('No se encontró la hoja de cálculo. Verifica SPREADSHEET_ID o que el script esté vinculado a la Sheet.');
    }

    // Obtener o crear la hoja "Tickets"
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Crear cabeceras si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Ticket ID',
        'Timestamp',
        'Nombre',
        'Cargo',
        'Sede',
        'Fecha Inicio Objetivo',
        'Fecha Fin Objetivo',
        'Objetivo',
        'Progreso'
      ]);
      var headerRange = sheet.getRange(1, 1, 1, 9);
      headerRange.setBackground('#4c1d95');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
    }

    // Añadir fila con los datos del ticket
    sheet.appendRow([
      data.ticketId    || '',
      data.timestamp   || new Date().toLocaleString('es-CO'),
      data.nombre      || '',
      data.cargo       || '',
      data.sede        || '',
      data.fechaInicio || '',
      data.fechaFin    || '',
      data.objetivo    || '',
      data.progreso    || ''
    ]);

    sheet.autoResizeColumns(1, 9);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'OK', ticketId: data.ticketId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // Registrar el error en una hoja de logs para poder depurarlo
    try {
      var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs_Error');
      if (!logSheet) logSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Logs_Error');
      logSheet.appendRow([new Date().toISOString(), err.message, JSON.stringify(e)]);
    } catch(_) {}

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ERROR', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ------------------------------------------------------------
//  doGet: responde a peticiones GET (útil para pruebas)
// ------------------------------------------------------------
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'HoriViaje API activa', version: '2.1' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
//  testConexion: ejecuta esto manualmente en Apps Script
//  para verificar que la conexión con la Sheet funciona.
// ------------------------------------------------------------
function testConexion() {
  var ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log('ERROR: No se encontró el spreadsheet.');
    return;
  }
  Logger.log('Spreadsheet OK: ' + ss.getName());
  var sheet = ss.getSheetByName(SHEET_NAME);
  Logger.log(sheet ? 'Hoja "Tickets" encontrada' : 'Hoja "Tickets" NO existe (se creará al primer envío)');
}
