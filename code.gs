// ============================================================
//  HoriViaje Ticket — Google Apps Script
//  Horisoes IPS
//  Deployment ID: AKfycbwqphjujw97m5hcQGkQxR8csCJNPujD0tSytIO7flD7fW3H1p656oCVEg9ZdqPRoaEt
// ============================================================

var SHEET_NAME = 'Tickets';

// ------------------------------------------------------------
//  doPost: recibe el JSON enviado desde el formulario HTML
// ------------------------------------------------------------
function doPost(e) {
  try {
    var raw  = e.postData ? e.postData.contents : '{}';
    var data = JSON.parse(raw);

    // getActiveSpreadsheet() funciona cuando el script está
    // vinculado (container-bound) a tu Google Sheet.
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

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
      // Estilo de cabecera
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

    // Autoajustar columnas
    sheet.autoResizeColumns(1, 9);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'OK', ticketId: data.ticketId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
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
    .createTextOutput(JSON.stringify({ status: 'HoriViaje API activa', version: '2.0' }))
    .setMimeType(ContentService.MimeType.JSON);
}