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

  // Obtener todos los socios de inscripción
  var dataInscripcion = inscripcionSheet.getRange('A8:A').getValues();
  var importRows = dataInscripcion.findIndex(function(r) { return !r[0] && r[0] !== 0; });
  if (importRows === -1) importRows = dataInscripcion.length;
  if (importRows === 0) return Logger.log('No hay filas para importar.');

  // Obtener socios que ya existen en el condensado
  var totalRows = sheetCondensado.getMaxRows();
  var last3Start = totalRows - 2;
  var filasCondensado = last3Start - 8;
  
  var sociosExistentes = [];
  if (filasCondensado > 0) {
    var dataCondensado = sheetCondensado.getRange(8, 1, filasCondensado, 1).getValues();
    for (var i = 0; i < dataCondensado.length; i++) {
      if (dataCondensado[i][0]) {
        sociosExistentes.push(dataCondensado[i][0].toString());
      }
    }
  }

  // Identificar socios nuevos que necesitan ser agregados
  var sociosNuevos = [];
  for (var i = 0; i < importRows; i++) {
    var numeroSocio = dataInscripcion[i][0];
    if (numeroSocio && sociosExistentes.indexOf(numeroSocio.toString()) === -1) {
      sociosNuevos.push({
        fila: 8 + i,
        numeroSocio: numeroSocio
      });
    }
  }

  if (sociosNuevos.length === 0) {
    return Logger.log('No hay socios nuevos para agregar. Todos los socios ya están registrados.');
  }

  // Calcular dónde empezar a agregar los nuevos socios
  var filaInicioNuevos = 8 + sociosExistentes.length;
  
  // Asegurar que hay suficientes filas para los nuevos socios
  var filasNecesarias = sociosExistentes.length + sociosNuevos.length;
  if (filasCondensado < filasNecesarias) {
    var filasFaltantes = filasNecesarias - filasCondensado;
    sheetCondensado.insertRowsAfter(filaInicioNuevos - 1, filasFaltantes);
  }

  // Agregar solo los socios nuevos
  var sociosAgregados = 0;
  for (var i = 0; i < sociosNuevos.length; i++) {
    var socioNuevo = sociosNuevos[i];
    var filaDestino = filaInicioNuevos + i;
    
    for (var col = 1; col <= 6; col++) {
      if (col === 5) continue; // Saltar columna E
      var colLetter = String.fromCharCode(64 + col);
      var formula = '=IMPORTRANGE("' + fileUrl + '", "Inscripción!' + colLetter + socioNuevo.fila + '")';
      sheetCondensado.getRange(filaDestino, col).setFormula(formula);
    }
    sociosAgregados++;
  }
  
  Logger.log('Proceso completado exitosamente:');
  Logger.log('- Socios existentes: ' + sociosExistentes.length);
  Logger.log('- Socios nuevos agregados: ' + sociosAgregados);
  Logger.log('- Total socios en condensado: ' + (sociosExistentes.length + sociosAgregados));
}