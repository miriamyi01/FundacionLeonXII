function registrarSociosCondensado() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var anioActual = new Date().getFullYear();

  function normalizarCodigo(codigo) {
    return (codigo === null || codigo === undefined) ? '' : codigo.toString().trim();
  }

  function construirFormulaNombreCompleto(fileUrl, filaOrigen) {
    return '=TRIM(IFERROR(IMPORTRANGE("' + fileUrl + '", "Inscripción!B' + filaOrigen + '"), "")&" "&IFERROR(IMPORTRANGE("' + fileUrl + '", "Inscripción!C' + filaOrigen + '"), "")&" "&IFERROR(IMPORTRANGE("' + fileUrl + '", "Inscripción!D' + filaOrigen + '"), ""))';
  }

  function extraerFilaDesdeFormulaImportRange(formula) {
    if (!formula) return null;
    var m = formula.match(/Inscripci[oó]n![A-Z]+(\d+)/i);
    return m ? parseInt(m[1], 10) : null;
  }

  function ordenarSociosAlfabeticamente(sheet) {
    var ultimaFila = sheet.getLastRow();
    if (ultimaFila < 4) return;

    var ultimaColumna = sheet.getLastColumn();
    var rangoDatos = sheet.getRange(4, 1, ultimaFila - 3, ultimaColumna);
    rangoDatos.sort({ column: 2, ascending: true });

    var codigosOrdenados = sheet.getRange(4, 1, ultimaFila - 3, 1).getValues();
    for (var i = 0; i < codigosOrdenados.length; i++) {
      var codigo = normalizarCodigo(codigosOrdenados[i][0]);
      if (!codigo) continue;
      var fila = 4 + i;
      var formulaEsperada = '=SUM(F' + fila + ':' + fila + ')';
      var celdaFormula = sheet.getRange(fila, 5);
      if (celdaFormula.getFormula() !== formulaEsperada) {
        celdaFormula.setFormula(formulaEsperada);
      }
    }
  }

  // Nombres de carpetas para buscar
  var rootFolderName = 'GA0452 METAMORFOSIS';
  var fichaFolderName = 'GA0452-FICHA DE INSCRIPCIÓN';

  // Buscar la carpeta raíz
  var folders = DriveApp.getFolders();
  var rootFolder = null;
  while (folders.hasNext()) {
    var f = folders.next();
    if (f.getName() === rootFolderName) {
      rootFolder = f;
      break;
    }
  }
  if (!rootFolder) {
    Logger.log('No se encontró la carpeta raíz: ' + rootFolderName);
    return;
  }

  // Buscar la carpeta de fichas dentro de la carpeta raíz
  var carpetaFicha = null;
  var subFolders = rootFolder.getFolders();
  while (subFolders.hasNext()) {
    var sf = subFolders.next();
    if (sf.getName().indexOf(fichaFolderName) !== -1) {
      carpetaFicha = sf;
      break;
    }
  }
  if (!carpetaFicha) {
    Logger.log('No se encontró la carpeta ' + fichaFolderName + ' dentro de ' + rootFolderName);
    return;
  }

  var file = null;
  var files = carpetaFicha.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (f.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
    var nombre = f.getName();
    if (nombre.indexOf(anioActual + ' 01 Lista inscripcion GA Metamorfosis') !== -1) {
      file = f;
      break;
    }
  }

  var fileId = file.getId();
  var fileUrl = 'https://docs.google.com/spreadsheets/d/' + fileId;
  var inscripcionSheet = SpreadsheetApp.openById(fileId).getSheetByName('Inscripción');
  if (!inscripcionSheet) return Logger.log('No se encontró la hoja Inscripción en el archivo de origen.');

  // Obtener todos los socios de inscripción
  var dataInscripcion = inscripcionSheet.getRange('A7:A').getValues();
  var importRows = dataInscripcion.findIndex(function(r) { return !r[0] && r[0] !== 0; });
  if (importRows === -1) importRows = dataInscripcion.length;
  if (importRows === 0) return Logger.log('No hay filas para importar.');

  function procesarHojaCondensado(nombreHoja) {
    var sheetCondensado = ss.getSheetByName(nombreHoja);
    if (!sheetCondensado) {
      Logger.log('No se encontró la hoja ' + nombreHoja + '.');
      return;
    }

    var celdaTotal = sheetCondensado.getRange('E2');
    var formulaTotalEsperada = '=SUM(E4:E)';
    if (celdaTotal.getFormula() !== formulaTotalEsperada) {
      celdaTotal.setFormula(formulaTotalEsperada);
    }

    // Obtener socios que ya existen en el condensado
    var ultimaFilaHoja = sheetCondensado.getLastRow();
    var filasCondensado = Math.max(0, ultimaFilaHoja - 3);
    var ultimaFilaConSocio = 3;

    var sociosExistentes = [];
    var sociosExistentesMap = {};
    if (filasCondensado > 0) {
      var dataCondensado = sheetCondensado.getRange(4, 1, filasCondensado, 1).getValues();
      for (var i = 0; i < dataCondensado.length; i++) {
        var codigoExistente = normalizarCodigo(dataCondensado[i][0]);
        if (codigoExistente) {
          sociosExistentes.push(codigoExistente);
          sociosExistentesMap[codigoExistente] = true;
          var filaExistente = 4 + i;
          ultimaFilaConSocio = filaExistente;

          var formulaAExistente = sheetCondensado.getRange(filaExistente, 1).getFormula();
          var filaOrigenExistente = extraerFilaDesdeFormulaImportRange(formulaAExistente);
          if (filaOrigenExistente) {
            var formulaNombreCompletoExistente = construirFormulaNombreCompleto(fileUrl, filaOrigenExistente);
            var celdaNombreExistente = sheetCondensado.getRange(filaExistente, 2);
            if (celdaNombreExistente.getFormula() !== formulaNombreCompletoExistente) {
              celdaNombreExistente.setFormula(formulaNombreCompletoExistente);
            }
          }

          var celdaFormulaExistente = sheetCondensado.getRange(filaExistente, 5);
          var formulaFilaEsperada = '=SUM(F' + filaExistente + ':' + filaExistente + ')';
          if (celdaFormulaExistente.getFormula() !== formulaFilaEsperada) {
            celdaFormulaExistente.setFormula(formulaFilaEsperada);
          }
        }
      }
    }

    // Identificar socios nuevos que necesitan ser agregados
    var sociosNuevos = [];
    for (var i = 0; i < importRows; i++) {
      var numeroSocio = dataInscripcion[i][0];
      var codigoNuevo = normalizarCodigo(numeroSocio);
      if (codigoNuevo && !sociosExistentesMap[codigoNuevo]) {
        sociosNuevos.push({
          fila: 7 + i,
          numeroSocio: numeroSocio,
          codigoSocio: codigoNuevo
        });
        sociosExistentesMap[codigoNuevo] = true;
      }
    }

    if (sociosNuevos.length === 0) {
      ordenarSociosAlfabeticamente(sheetCondensado);
      Logger.log('[' + nombreHoja + '] No hay socios nuevos para agregar. Todos los socios ya están registrados.');
      return;
    }

    // Calcular dónde empezar a agregar los nuevos socios
    var filaInicioNuevos = Math.max(4, ultimaFilaConSocio + 1);
    
    // Asegurar que hay suficientes filas físicas para los nuevos socios
    var ultimaFilaNecesaria = filaInicioNuevos + sociosNuevos.length - 1;
    var maxRows = sheetCondensado.getMaxRows();
    if (maxRows < ultimaFilaNecesaria) {
      var filasFaltantes = ultimaFilaNecesaria - maxRows;
      sheetCondensado.insertRowsAfter(maxRows, filasFaltantes);
    }

    // Agregar solo los socios nuevos
    var sociosAgregados = 0;
    for (var i = 0; i < sociosNuevos.length; i++) {
      var socioNuevo = sociosNuevos[i];
      var filaDestino = filaInicioNuevos + i;

      var formulaCodigo = '=IMPORTRANGE("' + fileUrl + '", "Inscripción!A' + socioNuevo.fila + '")';
      sheetCondensado.getRange(filaDestino, 1).setFormula(formulaCodigo);

      var formulaNombreCompleto = construirFormulaNombreCompleto(fileUrl, socioNuevo.fila);
      sheetCondensado.getRange(filaDestino, 2).setFormula(formulaNombreCompleto);

      var celdaFormulaNueva = sheetCondensado.getRange(filaDestino, 5);
      var formulaNuevaEsperada = '=SUM(F' + filaDestino + ':' + filaDestino + ')';
      if (celdaFormulaNueva.getFormula() !== formulaNuevaEsperada) {
        celdaFormulaNueva.setFormula(formulaNuevaEsperada);
      }
      sociosAgregados++;
    }

    Logger.log('[' + nombreHoja + '] Proceso completado exitosamente:');
    Logger.log('[' + nombreHoja + '] - Socios existentes: ' + sociosExistentes.length);
    Logger.log('[' + nombreHoja + '] - Socios nuevos agregados: ' + sociosAgregados);
    Logger.log('[' + nombreHoja + '] - Total socios en condensado: ' + (sociosExistentes.length + sociosAgregados));

    ordenarSociosAlfabeticamente(sheetCondensado);
  }

  procesarHojaCondensado('Ahorros y Retiros');
  procesarHojaCondensado('Fondo de Emergencia');
}