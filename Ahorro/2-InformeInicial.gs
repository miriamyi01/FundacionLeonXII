function registrarSociosCondensado() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fileObj = DriveApp.getFileById(ss.getId());
  var parents = fileObj.getParents();
  var folder = parents.hasNext() ? parents.next() : null;
  if (!folder) return Logger.log('El archivo no está dentro de una carpeta en Drive.');

  var files = folder.getFiles(), file = null;
  while (files.hasNext()) {
    var f = files.next();
    if (f.getName().indexOf('03 LISTA DE INSCRIPCION') !== -1) { file = f; break; }
  }
  if (!file) return Logger.log('No se encontró el archivo 03 LISTA DE INSCRIPCION en la carpeta.');

  var fileId = file.getId();
  var fileUrl = 'https://docs.google.com/spreadsheets/d/' + fileId;
  var sheetCondensado = ss.getSheetByName('Ahorros y Retiros');
  if (!sheetCondensado) return Logger.log('No se encontró la hoja Ahorros y Retiros.');

  var inscripcionSheet = SpreadsheetApp.openById(fileId).getSheetByName('Inscripción');
  if (!inscripcionSheet) return Logger.log('No se encontró la hoja Inscripción en el archivo de origen.');

  var data = inscripcionSheet.getRange('A8:A').getValues();
  var importRows = data.findIndex(function(r) { return !r[0] && r[0] !== 0; });
  if (importRows === -1) importRows = data.length;
  if (importRows === 0) return Logger.log('No hay filas para importar.');

  // Asegura que hay suficientes filas en el condensado para los socios (omite las últimas 3 filas)
  var totalRows = sheetCondensado.getMaxRows();
  var last3Start = totalRows - 2;
  var filasActuales = last3Start - 8;
  var filasNecesarias = importRows;

  if (filasActuales < filasNecesarias) {
    // Inserta filas debajo de la primera fila de socios (después de la fila 8)
    var filasFaltantes = filasNecesarias - filasActuales;
    sheetCondensado.insertRowsAfter(8, filasFaltantes);
  } else if (filasActuales > filasNecesarias) {
    sheetCondensado.deleteRows(8, filasActuales - filasNecesarias);
  }

  sheetCondensado.getRange(8, 1, importRows, 6).clearContent();

  for (var row = 0; row < importRows; row++) {
    var inscRow = 8 + row;
    for (var col = 1; col <= 6; col++) {
      if (col === 5) continue;
      var colLetter = String.fromCharCode(64 + col);
      var formula = '=IMPORTRANGE("' + fileUrl + '", "Inscripción!' + colLetter + inscRow + '")';
      sheetCondensado.getRange(inscRow, col).setFormula(formula);
    }
  }
  Logger.log('Fórmulas IMPORTRANGE individuales insertadas automáticamente y filas ajustadas.');
}