function procesarAvales() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCondensado = ss.getSheetByName('Ahorros y Retiros');
  if (!sheetCondensado) {
    Logger.log('No se encontró la hoja Ahorros y Retiros.');
    return;
  }

  var parentFolder;
  try {
    var parents = DriveApp.getFileById(ss.getId()).getParents();
    parentFolder = parents.hasNext() ? parents.next() : null;
  } catch (e) {
    Logger.log('Error obteniendo carpeta principal: ' + e);
    return;
  }
  
  var baseName = sheetCondensado.getRange("A8").getValue().toString().substring(0, 4);
  var sociosFolderName = baseName + "-SOCIOS AS";
  var sociosFolders = parentFolder.getFoldersByName(sociosFolderName);
  if (!sociosFolders.hasNext()) {
    Logger.log('No se encontró la subcarpeta de socios: ' + sociosFolderName);
    return;
  }
  var sociosMainFolder = sociosFolders.next();

  var sociosMap = {};
  var socioFoldersIter = sociosMainFolder.getFolders();
  while (socioFoldersIter.hasNext()) {
    var folder = socioFoldersIter.next();
    var nombreCarpeta = folder.getName();
    var partes = nombreCarpeta.split(" ");
    if (partes.length > 1) {
      var codigo = partes[0];
      var iniciales = partes[1];
      
      var tarjetaFiles = folder.getFiles();
      while (tarjetaFiles.hasNext()) {
        var tf = tarjetaFiles.next();
        if (tf.getName().indexOf(iniciales) !== -1 && tf.getName().indexOf('TARJETA AHORRO Y PRESTAMO') !== -1) {
          sociosMap[codigo] = {
            tarjetaId: tf.getId(),
            iniciales: iniciales,
            nombreCompleto: nombreCarpeta
          };
          break;
        }
      }
    }
  }

  var contadorAvales = 0;

  for (var codigo in sociosMap) {
    var socioInfo = sociosMap[codigo];
    
    try {
      var tarjeta = SpreadsheetApp.openById(socioInfo.tarjetaId);
      var hojaAhorro = tarjeta.getSheetByName('Tarjeta Ahorro');
      
      if (!hojaAhorro) continue;

      var lastRow = hojaAhorro.getLastRow();
      var filaInicio = Math.max(lastRow - 3, 1);
      
      for (var fila = filaInicio; fila <= lastRow; fila++) {
        var numeroPrestamo = hojaAhorro.getRange(fila, 1).getValue();
        var codigoPrestatario = hojaAhorro.getRange(fila, 2).getValue();
        var nombrePrestatario = hojaAhorro.getRange(fila, 3).getValue(); // Nombre completo del prestatario
        var montoAvalado = hojaAhorro.getRange(fila, 4).getValue();
        
        if (!numeroPrestamo || !codigoPrestatario || !nombrePrestatario || !montoAvalado) continue;
        
        if (!sociosMap[codigoPrestatario]) {
          continue;
        }
        
        var prestatarioInfo = sociosMap[codigoPrestatario];
        var nombreCompletoAval = hojaAhorro.getRange("B1").getValue().toString(); // Nombre completo del aval desde B1
        var nombreAval = nombreCompletoAval.split(' ')[0]; // Primer nombre del aval
        var primerNombrePrestatario = nombrePrestatario.toString().split(' ')[0]; // Primer nombre del prestatario
        
        try {
          var tarjetaPrestatario = SpreadsheetApp.openById(prestatarioInfo.tarjetaId);
          var nombreHojaPrestamo = 'Tarjeta Prestamo #' + numeroPrestamo;
          var hojaPrestamo = tarjetaPrestatario.getSheetByName(nombreHojaPrestamo);
          
          if (!hojaPrestamo) {
            Logger.log('No se encontró hoja de préstamo: ' + nombreHojaPrestamo);
            continue;
          }
          
          var avalAgregado = false;
          
          // Primero verificar si el aval ya existe para evitar duplicados
          var avalYaExiste = false;
          for (var filaVerificar = 4; filaVerificar <= 7; filaVerificar++) {
            var nombreExistenteVerificar = hojaPrestamo.getRange(filaVerificar, 8).getValue();
            if (nombreExistenteVerificar && nombreExistenteVerificar.toString().trim() === nombreAval.trim()) {
              avalYaExiste = true;
              Logger.log('Aval ya existe: ' + nombreAval + ' para préstamo #' + numeroPrestamo + ' - Omitiendo duplicado');
              break;
            }
          }
          
          // Solo agregar el aval si no existe ya
          if (!avalYaExiste) {
            for (var filaAval = 4; filaAval <= 7; filaAval++) {
              var nombreExistente = hojaPrestamo.getRange(filaAval, 8).getValue();
              
              if (!nombreExistente) {
                hojaPrestamo.getRange(filaAval, 8).setValue(nombreAval);
                hojaPrestamo.getRange(filaAval, 9).setValue(montoAvalado);
                avalAgregado = true;
                break;
              }
            }
          }
          
          if (avalAgregado) {
            // Fórmula para obtener el último saldo pendiente del préstamo (usando ID dinámico)
            var formulaSaldoPendiente = '=IFERROR(INDEX(FILTER(IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!H13:H23"),IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!H13:H23")<>"",IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!H13:H23")<>0),COUNTA(FILTER(IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!H13:H23"),IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!H13:H23")<>"",IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!H13:H23")<>0))),0)';
            // Fórmula para obtener la fecha de compromiso del préstamo
            var formulaFechaCompromiso = '=IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!C5")';
            
            // Aplicar las fórmulas a las celdas correspondientes
            hojaAhorro.getRange(fila, 5).setFormula(formulaSaldoPendiente);
            hojaAhorro.getRange(fila, 6).setFormula(formulaFechaCompromiso);
            
            contadorAvales++;
            Logger.log('Aval procesado: ' + nombreAval + ' (' + codigo + ') por $' + montoAvalado + ' para préstamo #' + numeroPrestamo + ' de ' + codigoPrestatario + ' (' + primerNombrePrestatario + ')');
          }
          
        } catch (e) {
          Logger.log('Error procesando aval de ' + codigo + ': ' + e);
        }
      }
      
    } catch (e) {
      Logger.log('Error abriendo tarjeta de ' + codigo + ': ' + e);
    }
  }
  
  Logger.log('Procesamiento completado. Total de avales procesados: ' + contadorAvales);
}