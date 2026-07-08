function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    ordenCompra(data);
    return ContentService.createTextOutput('ok');
  } catch(err) {
    logError(err, e);
    return ContentService.createTextOutput('error: ' + err.toString());
  }
}

function doGet(e) {
  // Consulta del siguiente folio disponible (usada por el formulario web)
  if (e.parameter.accion === 'consecutivo') {
    var next = getNextConsecutivo(e.parameter.sheetId);
    var payload = JSON.stringify({consecutivo: next});
    if (e.parameter.callback) {
      return ContentService
        .createTextOutput(e.parameter.callback + '(' + payload + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('ok');
}

function getNextConsecutivo(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('Órdenes de Compra');
    if (!sheet || sheet.getLastRow() <= 1) return 1;
    var lastRow = sheet.getLastRow();
    var lastOC = sheet.getRange(lastRow, 1).getValue(); // Columna A = "No. OC", ej. OC-VER-005
    var match = String(lastOC).match(/(\d+)\s*$/);
    var lastNum = match ? parseInt(match[1], 10) : 0;
    return lastNum + 1;
  } catch(err) {
    return 1;
  }
}

function logError(err, e) {
  try {
    var ss = SpreadsheetApp.openById('1uneUIPolFIXZjegQXkraHnTzDGFZ5-HU7ZCEe0V8FV8');
    var sheet = ss.getSheetByName('Errores') || ss.insertSheet('Errores');
    sheet.appendRow([
      new Date(),
      err.toString(),
      err.stack || 'sin stack',
      e && e.postData ? e.postData.contents.substring(0,500) : 'sin data'
    ]);
  } catch(e2) {}
}

function ordenCompra(data) {
  var nombreCarpeta = data.ocNum + ' - ' + data.proveedor;
  var carpetaPadre = DriveApp.getFolderById(data.folderId);
  var carpeta = carpetaPadre.createFolder(nombreCarpeta);

  if (data.pdfFinal && data.pdfFinal.data) {
    var pdfBytes = Utilities.base64Decode(data.pdfFinal.data);
    var pdfBlob = Utilities.newBlob(pdfBytes, 'application/pdf', data.pdfFinal.name);
    carpeta.createFile(pdfBlob);
  }

  if (data.documentos) {
    for (var key in data.documentos) {
      var doc = data.documentos[key];
      if (!doc) continue;
      var bytes = Utilities.base64Decode(doc.data);
      var blob = Utilities.newBlob(bytes, doc.type, doc.name);
      carpeta.createFile(blob);
    }
  }

  var ss = SpreadsheetApp.openById(data.sheetId);
  var sheet = ss.getSheetByName('Órdenes de Compra') || ss.insertSheet('Órdenes de Compra');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['No. OC','Fecha','Fecha Entrega','Proveedor','Proyecto',
      'Concepto','O.T.','Moneda','Forma de Pago',
      'Subtotal','IVA','Total MXN','Adjuntos','Carpeta Drive']);
  }
  sheet.appendRow([
    data.ocNum, data.fecha, data.fechaEntrega,
    data.proveedor, data.proyecto, data.concepto,
    data.otNum||'---', data.moneda||'MXN', data.formaPago,
    data.subtotal, data.iva, data.total,
    data.documentos ? Object.keys(data.documentos).length+' adjunto(s)' : '0',
    carpeta.getUrl()
  ]);
}

function testScript() {
  var carpeta = DriveApp.getFolderById('1OZkCuWxeL-G4xC5LAj3wsjq4a1tsi5Oo');
  var nueva = carpeta.createFolder('TEST - borrar');
  Logger.log('OK: ' + nueva.getUrl());

  var ss = SpreadsheetApp.openById('1uneUIPolFIXZjegQXkraHnTzDGFZ5-HU7ZCEe0V8FV8');
  var sheet = ss.getActiveSheet();
  sheet.appendRow(['TEST']);
  Logger.log('Sheets OK');
}
