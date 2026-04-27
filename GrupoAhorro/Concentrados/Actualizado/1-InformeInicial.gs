function registrarSociosCondensado() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var anioActual = new Date().getFullYear();

  function normalizarCodigo(codigo) {
    return (codigo === null || codigo === undefined) ? '' : codigo.toString().trim();
  }

  function valorFEsValido(valor) {
    if (valor === null || valor === undefined) return false;
    if (typeof valor === 'string' && valor.trim() === '') return false;

    var numero = Number(valor);
    if (!isNaN(numero)) {
      return numero !== 0;
    }

    return true;
  }


  function unirCeldasNombre(sheet) {
    var ultimaFila = sheet.getLastRow();
    if (ultimaFila < 4) return;

    var filasDatos = ultimaFila - 3;
    sheet.getRange(4, 2, filasDatos, 3).breakApart(); // B:D

    var codigos = sheet.getRange(4, 1, filasDatos, 1).getValues();
    for (var i = 0; i < codigos.length; i++) {
      var codigo = normalizarCodigo(codigos[i][0]);
      if (!codigo) continue;
      var fila = 4 + i;
      sheet.getRange(fila, 2, 1, 3).merge(); // Unir B:D en la fila del socio
    }
  }

  // Nombres de carpetas para buscar
  var rootFolderName = 'GA0452 METAMORFOSIS';
  var fichaFolderName = 'GA0452-FICHA DE INSCRIPCIÓN';
  var sociosFolderName = 'GA0452-SOCIOS AS';

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

  // Buscar las carpetas necesarias dentro de la carpeta raíz
  var carpetaFicha = null;
  var sociosMainFolder = null;
  var subFolders = rootFolder.getFolders();
  while (subFolders.hasNext()) {
    var sf = subFolders.next();
    var sfName = sf.getName();
    if (sfName.indexOf(fichaFolderName) !== -1) {
      carpetaFicha = sf;
    } else if (sfName.indexOf(sociosFolderName) !== -1) {
      sociosMainFolder = sf;
    }
  }

  if (!carpetaFicha || !sociosMainFolder) {
    Logger.log('No se encontró alguna de las carpetas necesarias (' + fichaFolderName + ' o ' + sociosFolderName + ') dentro de ' + rootFolderName);
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
  var dataInscripcion = inscripcionSheet.getRange('A7:F').getValues();
  var importRows = dataInscripcion.findIndex(function(r) { return !r[0] && r[0] !== 0; });
  if (importRows === -1) importRows = dataInscripcion.length;
  if (importRows === 0) return Logger.log('No hay filas para importar.');

  // Helper function to find socio folder and card file, and construct URL/name
  function getSocioCardInfo(socioId, dataInscripcionRowIndex) {
    var socioFolder = null;
    var socioFoldersIter = sociosMainFolder.getFolders();
    while (socioFoldersIter.hasNext()) {
      var folder = socioFoldersIter.next();
      if (folder.getName().indexOf(socioId + " ") === 0) {
        socioFolder = folder;
        break;
      }
    }

    var fullName = dataInscripcion[dataInscripcionRowIndex][1] + ' ' + dataInscripcion[dataInscripcionRowIndex][2] + ' ' + dataInscripcion[dataInscripcionRowIndex][3];
    fullName = fullName.trim();

    if (!socioFolder) {
      Logger.log('No se encontró carpeta para socio: ' + socioId + '. Usando solo el nombre.');
      return { url: null, fullName: fullName };
    }

    var carpetaNombre = socioFolder.getName();
    var partes = carpetaNombre.split(" ");
    var iniciales = partes.length > 1 ? partes[1] : "";

    var tarjetaFiles = socioFolder.getFiles();
    var tarjetaFile = null;
    while (tarjetaFiles.hasNext()) {
      var tf = tarjetaFiles.next();
      if (
        tf.getName().indexOf(iniciales) !== -1 &&
        tf.getName().indexOf('TARJETA AHORRO Y PRESTAMO') !== -1
      ) {
        tarjetaFile = tf;
        break;
      }
    }

    if (!tarjetaFile) {
      Logger.log('No se encontró archivo TARJETA AHORRO Y PRESTAMO para socio: ' + socioId + ' (' + iniciales + '). Usando solo el nombre.');
      return { url: null, fullName: fullName, gidAhorro: null, gidEmergencia: null };
    }

    var tarjetaId = tarjetaFile.getId();
    var tarjetaUrl = 'https://docs.google.com/spreadsheets/d/' + tarjetaId;
    
    // Abrimos la tarjeta una sola vez para obtener los IDs de las pestañas (GIDs)
    var gidAhorro = "";
    var gidEmergencia = "";
    try {
      var tempSs = SpreadsheetApp.openById(tarjetaId);
      var shA = tempSs.getSheetByName('Tarjeta Ahorro');
      if (shA) gidAhorro = shA.getSheetId();
      var shE = tempSs.getSheetByName('Fondo de Emergencia');
      if (shE) gidEmergencia = shE.getSheetId();
    } catch (e) {
      Logger.log('Aviso: No se pudieron obtener los GIDs de las pestañas para el socio ' + socioId);
    }

    return { url: tarjetaUrl, fullName: fullName, gidAhorro: gidAhorro, gidEmergencia: gidEmergencia };
  }

  // Pre-compute socio card info to avoid repeated Drive API calls
  var socioCardInfoMap = {};
  for (var i = 0; i < importRows; i++) {
    var numeroSocio = dataInscripcion[i][0];
    var codigo = normalizarCodigo(numeroSocio);
    var valorFInscripcion = dataInscripcion[i][5];
    if (codigo && valorFEsValido(valorFInscripcion)) {
      socioCardInfoMap[codigo] = getSocioCardInfo(codigo, i);
    }
  }

  function procesarHojaCondensado(nombreHoja) {
    var sheetCondensado = ss.getSheetByName(nombreHoja);
    if (!sheetCondensado) {
      Logger.log('No se encontró la hoja ' + nombreHoja + '.');
      return;
    }

    var esHojaAhorros = nombreHoja === 'Ahorros y Retiros';
    var esHojaFondoSocial = nombreHoja === 'Fondo Social';

    var celdaTotalE = sheetCondensado.getRange('E2');
    var formulaTotalEEsperada = '=SUM(E4:E)';
    if (celdaTotalE.getFormula() !== formulaTotalEEsperada) {
      celdaTotalE.setFormula(formulaTotalEEsperada);
    }

    if (esHojaAhorros) {
      var celdaTotalG = sheetCondensado.getRange('G2');
      var formulaTotalGEsperada = '=SUM(G4:G)';
      if (celdaTotalG.getFormula() !== formulaTotalGEsperada) {
        celdaTotalG.setFormula(formulaTotalGEsperada);
      }

      var celdaTotalH = sheetCondensado.getRange('H2');
      var formulaTotalHEsperada = '=SUM(H4:H)';
      if (celdaTotalH.getFormula() !== formulaTotalHEsperada) {
        celdaTotalH.setFormula(formulaTotalHEsperada);
      }

      var celdaTotalI = sheetCondensado.getRange('I2');
      var formulaTotalIEsperada = '=SUM(I4:I)';
      if (celdaTotalI.getFormula() !== formulaTotalIEsperada) {
        celdaTotalI.setFormula(formulaTotalIEsperada);
      }
    }

    // Obtener socios que ya existen en el condensado
    var ultimaFilaHoja = sheetCondensado.getLastRow();
    var filasCondensado = Math.max(0, ultimaFilaHoja - 3);
    var ultimaFilaConSocio = 3;

    var sociosExistentesMap = {};
    var sociosProcessedCount = 0;

    if (filasCondensado > 0) {
      var existingSocioCodes = sheetCondensado.getRange(4, 1, filasCondensado, 1).getValues();
      for (var i = 0; i < existingSocioCodes.length; i++) {
        var codigoExistente = normalizarCodigo(existingSocioCodes[i][0]);
        if (codigoExistente) {
          var filaExistente = 4 + i;
          ultimaFilaConSocio = filaExistente;

          var cardInfo = socioCardInfoMap[codigoExistente];
          if (cardInfo) {
            // Update Column A (Socio ID) to be just the value
            sheetCondensado.getRange(filaExistente, 1).setValue(codigoExistente);

            // Update Column B (Name/Link)
            var targetGid = (esHojaAhorros || esHojaFondoSocial) ? cardInfo.gidAhorro : cardInfo.gidEmergencia;
            var finalUrl = cardInfo.url + (targetGid ? '#gid=' + targetGid : '');
            var newFormulaB = cardInfo.url ? '=HYPERLINK("' + finalUrl + '", "' + cardInfo.fullName + '")' : cardInfo.fullName;
            if (sheetCondensado.getRange(filaExistente, 2).getFormula() !== newFormulaB) {
              sheetCondensado.getRange(filaExistente, 2).setFormula(newFormulaB);
            }
          }

          if (esHojaAhorros) {
            var celdaFormulaEExistente = sheetCondensado.getRange(filaExistente, 5);
            var formulaEExistenteEsperada = '=SUMIF(J' + filaExistente + ':' + filaExistente + ',">0")';
            if (celdaFormulaEExistente.getFormula() !== formulaEExistenteEsperada) {
              celdaFormulaEExistente.setFormula(formulaEExistenteEsperada);
            }

            var celdaFormulaFExistente = sheetCondensado.getRange(filaExistente, 6);
            var formulaFExistenteEsperada = '=E' + filaExistente + '*F$2';
            if (celdaFormulaFExistente.getFormula() !== formulaFExistenteEsperada) {
              celdaFormulaFExistente.setFormula(formulaFExistenteEsperada);
            }

            var celdaFormulaGExistente = sheetCondensado.getRange(filaExistente, 7);
            var formulaGExistenteEsperada = '=E' + filaExistente + '+F' + filaExistente;
            if (celdaFormulaGExistente.getFormula() !== formulaGExistenteEsperada) {
              celdaFormulaGExistente.setFormula(formulaGExistenteEsperada);
            }

            var celdaFormulaHExistente = sheetCondensado.getRange(filaExistente, 8);
            var formulaHExistenteEsperada = '=SUMIF(J' + filaExistente + ':' + filaExistente + ',"<0")';
            if (celdaFormulaHExistente.getFormula() !== formulaHExistenteEsperada) {
              celdaFormulaHExistente.setFormula(formulaHExistenteEsperada);
            }

            var celdaFormulaIExistente = sheetCondensado.getRange(filaExistente, 9);
            var formulaIExistenteEsperada = '=E' + filaExistente + '+H' + filaExistente;
            if (celdaFormulaIExistente.getFormula() !== formulaIExistenteEsperada) {
              celdaFormulaIExistente.setFormula(formulaIExistenteEsperada);
            }
          } else if (nombreHoja === 'Fondo de Emergencia') {
            var celdaFormulaEEmergenciaExistente = sheetCondensado.getRange(filaExistente, 5);
            var formulaEEmergenciaExistente = '=SUM(F' + filaExistente + ':' + filaExistente + ')';
            if (celdaFormulaEEmergenciaExistente.getFormula() !== formulaEEmergenciaExistente) {
              celdaFormulaEEmergenciaExistente.setFormula(formulaEEmergenciaExistente);
            }
          } else if (esHojaFondoSocial) {
            var celdaSocialExistente = sheetCondensado.getRange(filaExistente, 5);
            var formulaSocialExistente = cardInfo.url ? '=IFERROR(IMPORTRANGE("' + cardInfo.url + '", "\'Tarjeta Ahorro\'!F1"), 0)' : '0';
            if (celdaSocialExistente.getFormula() !== formulaSocialExistente) {
              celdaSocialExistente.setFormula(formulaSocialExistente);
            }
          }
          sociosExistentesMap[codigoExistente] = true; // Mark as processed
          sociosProcessedCount++;
        }
      }
    }

    // Identificar socios nuevos que necesitan ser agregados
    var sociosNuevosParaAgregar = [];
    for (var i = 0; i < importRows; i++) {
      var numeroSocio = dataInscripcion[i][0];
      var codigoNuevo = normalizarCodigo(numeroSocio);
      var valorFInscripcion = dataInscripcion[i][5];
      if (codigoNuevo && valorFEsValido(valorFInscripcion) && !sociosExistentesMap[codigoNuevo]) {
        sociosNuevosParaAgregar.push({
          codigoSocio: codigoNuevo,
          dataInscripcionRowIndex: i // Store index to get card info
        });
        sociosExistentesMap[codigoNuevo] = true;
      }
    } // End of loop for identifying new socios

    if (sociosNuevosParaAgregar.length === 0) {
      unirCeldasNombre(sheetCondensado);
      Logger.log('[' + nombreHoja + '] No hay socios nuevos para agregar. Todos los socios ya están registrados.');
      return;
    }

    // Calcular dónde empezar a agregar los nuevos socios
    var filaInicioNuevos = Math.max(4, ultimaFilaHoja + 1);
    
    // Asegurar que hay suficientes filas físicas para los nuevos socios
    var ultimaFilaNecesaria = filaInicioNuevos + sociosNuevosParaAgregar.length - 1;
    var maxRows = sheetCondensado.getMaxRows();
    if (maxRows < ultimaFilaNecesaria) {
      var filasFaltantes = ultimaFilaNecesaria - maxRows;
      sheetCondensado.insertRowsAfter(maxRows, filasFaltantes);
    }

    // Add new socios
    var sociosAgregados = 0;
    for (var i = 0; i < sociosNuevosParaAgregar.length; i++) {
      var socioNuevo = sociosNuevosParaAgregar[i];
      var filaDestino = filaInicioNuevos + i;

      var cardInfo = socioCardInfoMap[socioNuevo.codigoSocio];
      if (!cardInfo) {
        Logger.log('Error: No se encontró información de tarjeta para socio nuevo ' + socioNuevo.codigoSocio + '. No se agregará.');
        continue;
      }

      // Set Column A (Socio ID)
      sheetCondensado.getRange(filaDestino, 1).setValue(socioNuevo.codigoSocio);

      // Set Column B (Name/Link)
      var targetGid = (esHojaAhorros || esHojaFondoSocial) ? cardInfo.gidAhorro : cardInfo.gidEmergencia;
      var finalUrl = cardInfo.url + (targetGid ? '#gid=' + targetGid : '');
      var newFormulaB = cardInfo.url ? '=HYPERLINK("' + finalUrl + '", "' + cardInfo.fullName + '")' : cardInfo.fullName;
      sheetCondensado.getRange(filaDestino, 2).setFormula(newFormulaB);

      if (esHojaAhorros) {
        var celdaFormulaENueva = sheetCondensado.getRange(filaDestino, 5);
        var formulaENuevaEsperada = '=SUMIF(J' + filaDestino + ':' + filaDestino + ',">0")';
        if (celdaFormulaENueva.getFormula() !== formulaENuevaEsperada) {
          celdaFormulaENueva.setFormula(formulaENuevaEsperada);
        }

        var celdaFormulaFNueva = sheetCondensado.getRange(filaDestino, 6);
        var formulaFNuevaEsperada = '=E' + filaDestino + '*F$2';
        if (celdaFormulaFNueva.getFormula() !== formulaFNuevaEsperada) {
          celdaFormulaFNueva.setFormula(formulaFNuevaEsperada);
        }

        var celdaFormulaGNueva = sheetCondensado.getRange(filaDestino, 7);
        var formulaGNuevaEsperada = '=E' + filaDestino + '+F' + filaDestino;
        if (celdaFormulaGNueva.getFormula() !== formulaGNuevaEsperada) {
          celdaFormulaGNueva.setFormula(formulaGNuevaEsperada);
        }

        var celdaFormulaHNueva = sheetCondensado.getRange(filaDestino, 8);
        var formulaHNuevaEsperada = '=SUMIF(J' + filaDestino + ':' + filaDestino + ',"<0")';
        if (celdaFormulaHNueva.getFormula() !== formulaHNuevaEsperada) {
          celdaFormulaHNueva.setFormula(formulaHNuevaEsperada);
        }

        var celdaFormulaINueva = sheetCondensado.getRange(filaDestino, 9);
        var formulaINuevaEsperada = '=E' + filaDestino + '+H' + filaDestino;
        if (celdaFormulaINueva.getFormula() !== formulaINuevaEsperada) {
          celdaFormulaINueva.setFormula(formulaINuevaEsperada);
        }
      } else if (nombreHoja === 'Fondo de Emergencia') {
        var celdaFormulaEEmergenciaNueva = sheetCondensado.getRange(filaDestino, 5);
        var formulaEEmergenciaNueva = '=SUM(F' + filaDestino + ':' + filaDestino + ')';
        if (celdaFormulaEEmergenciaNueva.getFormula() !== formulaEEmergenciaNueva) {
          celdaFormulaEEmergenciaNueva.setFormula(formulaEEmergenciaNueva);
        }
      } else if (esHojaFondoSocial) {
        var celdaSocialNueva = sheetCondensado.getRange(filaDestino, 5);
        var formulaSocialNueva = cardInfo.url ? '=IFERROR(IMPORTRANGE("' + cardInfo.url + '", "\'Tarjeta Ahorro\'!F1"), 0)' : '0';
        if (celdaSocialNueva.getFormula() !== formulaSocialNueva) {
          celdaSocialNueva.setFormula(formulaSocialNueva);
        }
      }
      sociosAgregados++;
    }

    Logger.log('[' + nombreHoja + '] Proceso completado exitosamente:');
    Logger.log('[' + nombreHoja + '] - Socios existentes actualizados: ' + sociosProcessedCount);
    Logger.log('[' + nombreHoja + '] - Socios nuevos agregados: ' + sociosAgregados);
    Logger.log('[' + nombreHoja + '] - Total socios en condensado: ' + (sociosProcessedCount + sociosAgregados));

    unirCeldasNombre(sheetCondensado);
  }

  procesarHojaCondensado('Ahorros y Retiros');
  procesarHojaCondensado('Fondo de Emergencia');
  procesarHojaCondensado('Fondo Social');
}