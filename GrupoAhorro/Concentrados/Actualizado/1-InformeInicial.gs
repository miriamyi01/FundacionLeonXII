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

  function construirFormulaNombreCompleto(fileUrl, filaOrigen) {
    return '=TRIM(IFERROR(IMPORTRANGE("' + fileUrl + '", "Inscripción!B' + filaOrigen + '"), "")&" "&IFERROR(IMPORTRANGE("' + fileUrl + '", "Inscripción!C' + filaOrigen + '"), "")&" "&IFERROR(IMPORTRANGE("' + fileUrl + '", "Inscripción!D' + filaOrigen + '"), ""))';
  }

  function extraerFilaDesdeFormulaImportRange(formula) {
    if (!formula) return null;
    var m = formula.match(/Inscripci[oó]n![A-Z]+(\d+)/i);
    return m ? parseInt(m[1], 10) : null;
  }

  function ordenarSociosAlfabeticamente(sheet, aplicarFormulasAhorros) {
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
      if (!aplicarFormulasAhorros) {
        var formulaEmergenciaEsperada = '=SUM(F' + fila + ':' + fila + ')';
        var celdaFormulaEmergencia = sheet.getRange(fila, 5);
        if (celdaFormulaEmergencia.getFormula() !== formulaEmergenciaEsperada) {
          celdaFormulaEmergencia.setFormula(formulaEmergenciaEsperada);
        }
        continue;
      }
      var formulaEEsperada = '=SUMIF(J' + fila + ':' + fila + ',">0")';
      var formulaFEsperada = '=E' + fila + '*F$2';

      var celdaFormulaE = sheet.getRange(fila, 5);
      if (celdaFormulaE.getFormula() !== formulaEEsperada) {
        celdaFormulaE.setFormula(formulaEEsperada);
      }

      var celdaFormulaF = sheet.getRange(fila, 6);
      if (celdaFormulaF.getFormula() !== formulaFEsperada) {
        celdaFormulaF.setFormula(formulaFEsperada);
      }

      var formulaGEsperada = '=E' + fila + '+F' + fila;
      var celdaFormulaG = sheet.getRange(fila, 7);
      if (celdaFormulaG.getFormula() !== formulaGEsperada) {
        celdaFormulaG.setFormula(formulaGEsperada);
      }

      var formulaHEsperada = '=SUMIF(J' + fila + ':' + fila + ',"<0")';
      var celdaFormulaH = sheet.getRange(fila, 8);
      if (celdaFormulaH.getFormula() !== formulaHEsperada) {
        celdaFormulaH.setFormula(formulaHEsperada);
      }

      var formulaIEsperada = '=E' + fila + '+H' + fila;
      var celdaFormulaI = sheet.getRange(fila, 9);
      if (celdaFormulaI.getFormula() !== formulaIEsperada) {
        celdaFormulaI.setFormula(formulaIEsperada);
      }
    }

    unirCeldasNombre(sheet);
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
  var dataInscripcion = inscripcionSheet.getRange('A7:F').getValues();
  var importRows = dataInscripcion.findIndex(function(r) { return !r[0] && r[0] !== 0; });
  if (importRows === -1) importRows = dataInscripcion.length;
  if (importRows === 0) return Logger.log('No hay filas para importar.');

  function procesarHojaCondensado(nombreHoja) {
    var sheetCondensado = ss.getSheetByName(nombreHoja);
    if (!sheetCondensado) {
      Logger.log('No se encontró la hoja ' + nombreHoja + '.');
      return;
    }

    var esHojaAhorros = nombreHoja === 'Ahorros y Retiros';

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
          } else {
            var celdaFormulaEEmergenciaExistente = sheetCondensado.getRange(filaExistente, 5);
            var formulaEEmergenciaExistente = '=SUM(F' + filaExistente + ':' + filaExistente + ')';
            if (celdaFormulaEEmergenciaExistente.getFormula() !== formulaEEmergenciaExistente) {
              celdaFormulaEEmergenciaExistente.setFormula(formulaEEmergenciaExistente);
            }
          }
        }
      }
    }

    // Identificar socios nuevos que necesitan ser agregados
    var sociosNuevos = [];
    for (var i = 0; i < importRows; i++) {
      var numeroSocio = dataInscripcion[i][0];
      var codigoNuevo = normalizarCodigo(numeroSocio);
      var valorFInscripcion = dataInscripcion[i][5];
      if (codigoNuevo && valorFEsValido(valorFInscripcion) && !sociosExistentesMap[codigoNuevo]) {
        sociosNuevos.push({
          fila: 7 + i,
          numeroSocio: numeroSocio,
          codigoSocio: codigoNuevo
        });
        sociosExistentesMap[codigoNuevo] = true;
      }
    }

    if (sociosNuevos.length === 0) {
      ordenarSociosAlfabeticamente(sheetCondensado, esHojaAhorros);
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
      } else {
        var celdaFormulaEEmergenciaNueva = sheetCondensado.getRange(filaDestino, 5);
        var formulaEEmergenciaNueva = '=SUM(F' + filaDestino + ':' + filaDestino + ')';
        if (celdaFormulaEEmergenciaNueva.getFormula() !== formulaEEmergenciaNueva) {
          celdaFormulaEEmergenciaNueva.setFormula(formulaEEmergenciaNueva);
        }
      }
      sociosAgregados++;
    }

    Logger.log('[' + nombreHoja + '] Proceso completado exitosamente:');
    Logger.log('[' + nombreHoja + '] - Socios existentes: ' + sociosExistentes.length);
    Logger.log('[' + nombreHoja + '] - Socios nuevos agregados: ' + sociosAgregados);
    Logger.log('[' + nombreHoja + '] - Total socios en condensado: ' + (sociosExistentes.length + sociosAgregados));

    ordenarSociosAlfabeticamente(sheetCondensado, esHojaAhorros);
  }

  procesarHojaCondensado('Ahorros y Retiros');
  procesarHojaCondensado('Fondo de Emergencia');
}