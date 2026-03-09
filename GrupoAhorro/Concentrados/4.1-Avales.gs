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
  
  // Usar directamente la carpeta principal como carpeta de socios
  var sociosMainFolder = parentFolder;

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

  // PRIMER SENTIDO: De aval a prestatario (código existente)
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
  
  // SEGUNDO SENTIDO: De prestatario a aval (búsqueda inversa)
  Logger.log('--- Iniciando búsqueda inversa (prestatario → aval) ---');
  var contadorAvalesInversos = 0;
  
  for (var codigoPrestatario in sociosMap) {
    var prestatarioInfo = sociosMap[codigoPrestatario];
    
    try {
      var tarjetaPrestatario = SpreadsheetApp.openById(prestatarioInfo.tarjetaId);
      var hojasPrestamo = tarjetaPrestatario.getSheets();
      
      for (var i = 0; i < hojasPrestamo.length; i++) {
        var hojaPrestamo = hojasPrestamo[i];
        var nombreHoja = hojaPrestamo.getName();
        
        // Solo procesar hojas de préstamo
        if (nombreHoja.indexOf('Tarjeta Prestamo #') !== 0) continue;
        
        var numeroPrestamo = nombreHoja.replace('Tarjeta Prestamo #', '');
        var nombrePrestatario = hojaPrestamo.getRange("B1").getValue().toString();
        var primerNombrePrestatario = nombrePrestatario.split(' ')[0];
        
        // Revisar avales en filas 4-7
        for (var filaAval = 4; filaAval <= 7; filaAval++) {
          var nombreAval = hojaPrestamo.getRange(filaAval, 8).getValue();
          var montoAvalado = hojaPrestamo.getRange(filaAval, 9).getValue();
          
          if (!nombreAval || !montoAvalado) continue;
          
          // Normalizar el nombre del aval para búsqueda
          var nombreAvalNormalizado = nombreAval.toString().toUpperCase()
            .replace(/Á/g, 'A').replace(/É/g, 'E').replace(/Í/g, 'I')
            .replace(/Ó/g, 'O').replace(/Ú/g, 'U').replace(/Ñ/g, 'N');
          
          // Dividir el nombre en palabras para búsqueda más flexible
          var palabrasAval = nombreAvalNormalizado.split(' ').filter(function(p) { return p.length > 2; });
          
          // Buscar carpeta del aval por nombre
          var avalEncontrado = false;
          var socioFoldersIterInverso = sociosMainFolder.getFolders();
          
          while (socioFoldersIterInverso.hasNext()) {
            var folderAval = socioFoldersIterInverso.next();
            var nombreCarpetaAval = folderAval.getName();
            var nombreCarpetaNormalizado = nombreCarpetaAval.toUpperCase()
              .replace(/Á/g, 'A').replace(/É/g, 'E').replace(/Í/g, 'I')
              .replace(/Ó/g, 'O').replace(/Ú/g, 'U').replace(/Ñ/g, 'N');
            
            // Verificar si todas las palabras significativas del aval están en la carpeta
            var coincidencias = 0;
            for (var p = 0; p < palabrasAval.length; p++) {
              if (nombreCarpetaNormalizado.indexOf(palabrasAval[p]) !== -1) {
                coincidencias++;
              }
            }
            
            // Si al menos 2 palabras coinciden (o todas si son menos de 2)
            var umbralCoincidencias = Math.min(2, palabrasAval.length);
            if (coincidencias >= umbralCoincidencias) {
              var partesAval = nombreCarpetaAval.split(" ");
              if (partesAval.length > 1) {
                var codigoAval = partesAval[0];
                var inicialesAval = partesAval[1];
                
                // Buscar tarjeta del aval
                var tarjetaFilesAval = folderAval.getFiles();
                while (tarjetaFilesAval.hasNext()) {
                  var tfAval = tarjetaFilesAval.next();
                  if (tfAval.getName().indexOf(inicialesAval) !== -1 && 
                      tfAval.getName().indexOf('TARJETA AHORRO Y PRESTAMO') !== -1) {
                    
                    try {
                      var tarjetaAval = SpreadsheetApp.openById(tfAval.getId());
                      var hojaAhorroAval = tarjetaAval.getSheetByName('Tarjeta Ahorro');
                      
                      if (!hojaAhorroAval) break;
                      
                      // Verificar si ya existe este registro
                      var lastRowAval = hojaAhorroAval.getLastRow();
                      var yaExiste = false;
                      
                      for (var filaVerif = Math.max(lastRowAval - 10, 1); filaVerif <= lastRowAval; filaVerif++) {
                        var numPrestamoExist = hojaAhorroAval.getRange(filaVerif, 1).getValue();
                        var codPrestatarioExist = hojaAhorroAval.getRange(filaVerif, 2).getValue();
                        
                        if (numPrestamoExist == numeroPrestamo && codPrestatarioExist == codigoPrestatario) {
                          yaExiste = true;
                          break;
                        }
                      }
                      
                      // Si no existe, agregarlo
                      if (!yaExiste) {
                        var nuevaFila = lastRowAval + 1;
                        hojaAhorroAval.getRange(nuevaFila, 1).setValue(numeroPrestamo);
                        hojaAhorroAval.getRange(nuevaFila, 2).setValue(codigoPrestatario);
                        hojaAhorroAval.getRange(nuevaFila, 3).setValue(nombrePrestatario);
                        hojaAhorroAval.getRange(nuevaFila, 4).setValue(montoAvalado);
                        
                        // Fórmulas
                        var formulaSaldoPendiente = '=IFERROR(INDEX(FILTER(IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHoja + '!H13:H23"),IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHoja + '!H13:H23")<>"",IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHoja + '!H13:H23")<>0),COUNTA(FILTER(IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHoja + '!H13:H23"),IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHoja + '!H13:H23")<>"",IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHoja + '!H13:H23")<>0))),0)';
                        var formulaFechaCompromiso = '=IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHoja + '!C5")';
                        
                        hojaAhorroAval.getRange(nuevaFila, 5).setFormula(formulaSaldoPendiente);
                        hojaAhorroAval.getRange(nuevaFila, 6).setFormula(formulaFechaCompromiso);
                        
                        contadorAvalesInversos++;
                        Logger.log('Aval inverso procesado: ' + nombreAval + ' (' + codigoAval + ') por $' + montoAvalado + ' para préstamo #' + numeroPrestamo + ' de ' + codigoPrestatario + ' (' + primerNombrePrestatario + ')');
                      }
                      
                      avalEncontrado = true;
                    } catch (e) {
                      Logger.log('Error procesando aval inverso ' + nombreAval + ': ' + e);
                    }
                    
                    break;
                  }
                }
              }
              
              if (avalEncontrado) break;
            }
          }
        }
      }
      
    } catch (e) {
      Logger.log('Error en búsqueda inversa para ' + codigoPrestatario + ': ' + e);
    }
  }
  
  Logger.log('Procesamiento completado.');
  Logger.log('Total de avales procesados (aval→prestatario): ' + contadorAvales);
  Logger.log('Total de avales procesados (prestatario→aval): ' + contadorAvalesInversos);
}