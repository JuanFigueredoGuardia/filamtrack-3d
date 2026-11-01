function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('FilamTrack 3D')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerRollos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  var data = [];
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0]) {
      data.push({
        id_qr: rows[i][0],
        material: rows[i][1],
        marca: rows[i][2],
        color: rows[i][3],
        hex: rows[i][4],
        total: rows[i][5],
        restante: rows[i][6],
        estado: rows[i][7],
        ubicacion: rows[i][8]
      });
    }
  }
  return data;
}

function guardarRollo(params) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  var idQr = params.id_qr;
  var encontrado = false;
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().toLowerCase() === idQr.toString().toLowerCase()) {
      sheet.getRange(i + 1, 2).setValue(params.material);
      sheet.getRange(i + 1, 3).setValue(params.marca);
      sheet.getRange(i + 1, 4).setValue(params.color);
      sheet.getRange(i + 1, 5).setValue(params.hex);
      sheet.getRange(i + 1, 6).setValue(params.total);
      sheet.getRange(i + 1, 7).setValue(params.restante);
      sheet.getRange(i + 1, 8).setValue(params.estado);
      sheet.getRange(i + 1, 9).setValue(params.ubicacion);
      encontrado = true;
      break;
    }
  }
  
  if (!encontrado) {
    sheet.appendRow([
      params.id_qr,
      params.material,
      params.marca,
      params.color,
      params.hex,
      params.total,
      params.restante,
      params.estado,
      params.ubicacion
    ]);
  }
  return "success";
}

function eliminarRollo(idQr) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().toLowerCase() === idQr.toString().toLowerCase()) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return "success";
}
