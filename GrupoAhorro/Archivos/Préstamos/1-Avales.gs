function procesarAvales() {
  var rootFolderName = 'GA0452 METAMORFOSIS';
  var sociosFolderName = 'GA0452-SOCIOS AS';

  function normalizarNombre(nombre) {
    return nombre.toString().replace(/\s+/g, ' ').trim().toLowerCase();
  }

  // Buscar la carpeta raíz
  var folders = DriveApp.getFolders();
  var rootFolder = null;
  while (folders.hasNext()) {
    var f = folders.next();
    if (normalizarNombre(f.getName()) === normalizarNombre(rootFolderName)) {
      rootFolder = f;
      break;
    }
  }
  if (!rootFolder) {
    Logger.log('No se encontró la carpeta raíz: ' + rootFolderName);
    return;
  }

  // Buscar la carpeta de socios dentro del folder raíz
  var sociosMainFolder = null;
  var subFolders = rootFolder.getFolders();
  while (subFolders.hasNext()) {
    var sf = subFolders.next();
    if (normalizarNombre(sf.getName()) === normalizarNombre(sociosFolderName)) {
      sociosMainFolder = sf;
      break;
    }
  }
  if (!sociosMainFolder) {
    Logger.log('No se encontró la carpeta ' + sociosFolderName + ' dentro de ' + rootFolderName);
    return;
  }

  // Obtener archivo base para tomar protecciones y formato condicional
  var scriptFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var scriptFolder = scriptFile.getParents().next();
  var filesBase = scriptFolder.getFiles();
  var baseFileId = null;
  while (filesBase.hasNext()) {
    var fileBase = filesBase.next();
    if (fileBase.getName().indexOf('04 TARJETA AHORRO Y PRESTAMO') !== -1 && fileBase.getMimeType() === MimeType.GOOGLE_SHEETS) {
      baseFileId = fileBase.getId();
      break;
    }
  }
  if (!baseFileId) {
    Logger.log('No se encontró el archivo base 04 TARJETA AHORRO Y PRESTAMO en la carpeta del script.');
    return;
  }

  var baseSS = SpreadsheetApp.openById(baseFileId);
  var hojaBasePrestamo = baseSS.getSheetByName('Tarjeta Prestamo #1');
  if (!hojaBasePrestamo) {
    Logger.log('No existe la hoja Tarjeta Prestamo #1 en el archivo base.');
    return;
  }

  var proteccionesRef = hojaBasePrestamo.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var reglasCondRef = hojaBasePrestamo.getConditionalFormatRules();

  var proteccionesRefMapGlobal = [];
  for (var prb = 0; prb < proteccionesRef.length; prb++) {
    var protRefBase = proteccionesRef[prb];
    proteccionesRefMapGlobal.push({
      rangeA1: protRefBase.getRange().getA1Notation(),
      editors: protRefBase.getEditors().map(function(u) { return u.getEmail(); }),
      warningOnly: protRefBase.isWarningOnly(),
      description: protRefBase.getDescription()
    });
  }

  // Buscar todos los socios y sus archivos
  var sociosMap = {};
  var socioFoldersIter = sociosMainFolder.getFolders();
  while (socioFoldersIter.hasNext()) {
    var folder = socioFoldersIter.next();
    var nombreCarpeta = folder.getName();
    var partes = nombreCarpeta.split(' ');
    if (partes.length > 1) {
      var codigo = partes[0];
      var iniciales = partes[1];
      var archivoNombre = iniciales + ' - TARJETA AHORRO Y PRESTAMO';
      var archivosExistentes = folder.getFilesByName(archivoNombre);
      var destFile = null;
      if (archivosExistentes.hasNext()) {
        destFile = archivosExistentes.next();
      }
      if (destFile) {
        sociosMap[codigo] = {
          tarjetaId: destFile.getId(),
          iniciales: iniciales,
          nombreCompleto: nombreCarpeta,
          folder: folder
        };
      }
    }
  }



  // -------- NUEVO: Validar hojas de préstamo con permisos y celdas clave --------
  for (var codigoSocio in sociosMap) {
    var socioInfo = sociosMap[codigoSocio];
    try {
      var tarjeta = SpreadsheetApp.openById(socioInfo.tarjetaId);

      function obtenerNumeroPrestamoDesdeNombre(nombreHoja) {
        var texto = nombreHoja.toString().trim();
        var match = texto.match(/^tarjeta\s+pr[eé]stamo\s*#?\s*(\d+)$/i);
        return match ? parseInt(match[1], 10) : null;
      }

      function esNombreRelacionadoAPrestamo(nombreHoja) {
        return /\bpr[eé]stamo\b/i.test(nombreHoja.toString());
      }

      // Buscar todas las hojas de préstamo existentes
      var hojas = tarjeta.getSheets();
      var hojasPrestamo = [];
      var maxN = 0;
      for (var i = 0; i < hojas.length; i++) {
        var nombre = hojas[i].getName();
        var n = obtenerNumeroPrestamoDesdeNombre(nombre);
        if (n !== null) {
          hojasPrestamo.push({n: n, sheet: hojas[i]});
          if (n > maxN) maxN = n;
        }
      }

      // Renombrar hojas de préstamo con variaciones al formato Tarjeta Prestamo #n
      for (var k = 0; k < hojas.length; k++) {
        var hojaActual = hojas[k];
        var nombreActual = hojaActual.getName();
        if (!esNombreRelacionadoAPrestamo(nombreActual)) continue;

        var numeroCorrecto = obtenerNumeroPrestamoDesdeNombre(nombreActual);
        if (numeroCorrecto !== null && nombreActual === ('Tarjeta Prestamo #' + numeroCorrecto)) {
          continue;
        }

        var siguienteN = maxN + 1;
        var nombreNuevo = 'Tarjeta Prestamo #' + siguienteN;
        while (tarjeta.getSheetByName(nombreNuevo)) {
          siguienteN++;
          nombreNuevo = 'Tarjeta Prestamo #' + siguienteN;
        }

        hojaActual.setName(nombreNuevo);
        hojasPrestamo.push({n: siguienteN, sheet: hojaActual});
        maxN = siguienteN;
        Logger.log('Hoja renombrada: ' + nombreActual + ' → ' + nombreNuevo + ' para socio ' + codigoSocio);
      }


      // Regla: si C30 de un préstamo no dice "Completado", borrar el siguiente préstamo
      hojasPrestamo.sort(function(a, b) { return a.n - b.n; });
      for (var hp = 0; hp < hojasPrestamo.length - 1; hp++) {
        var hojaActualPrestamo = hojasPrestamo[hp].sheet;
        var estadoPrestamo = hojaActualPrestamo.getRange('C30').getValue();
        var estadoNormalizado = estadoPrestamo ? estadoPrestamo.toString().trim().toLowerCase() : '';

        if (estadoNormalizado !== 'completado') {
          var siguientePrestamo = hojasPrestamo[hp + 1];
          try {
            tarjeta.deleteSheet(siguientePrestamo.sheet);
            Logger.log('Préstamo eliminado por existir otro activo: Tarjeta Prestamo #' + siguientePrestamo.n + ' (socio ' + codigoSocio + ')');
            hojasPrestamo.splice(hp + 1, 1);
          } catch (errBorrarPrestamo) {
            Logger.log('No se pudo borrar Tarjeta Prestamo #' + siguientePrestamo.n + ' para socio ' + codigoSocio + ': ' + errBorrarPrestamo);
          }
          break;
        }
      }


      // Para cada hoja de préstamo existente, asegurar permisos, celdas y formato
      for (var j = 0; j < hojasPrestamo.length; j++) {
        var n = hojasPrestamo[j].n;
        var hoja = hojasPrestamo[j].sheet;
        // Sincronizar permisos desde archivo base (faltantes y diferentes)
        var proteccionesActuales = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        var mapaActual = {};
        for (var pa = 0; pa < proteccionesActuales.length; pa++) {
          var protActual = proteccionesActuales[pa];
          mapaActual[protActual.getRange().getA1Notation()] = protActual;
        }

        function arraysEqual(a, b) {
          if (a.length !== b.length) return false;
          var aa = a.slice().sort();
          var bb = b.slice().sort();
          for (var ai = 0; ai < aa.length; ai++) {
            if (aa[ai] !== bb[ai]) return false;
          }
          return true;
        }

        for (var pb = 0; pb < proteccionesRefMapGlobal.length; pb++) {
          var base = proteccionesRefMapGlobal[pb];
          var match = mapaActual[base.rangeA1];
          var requiereActualizacion = true;

          if (match) {
            var editorsMatch = match.getEditors().map(function(u) { return u.getEmail(); });
            if (
              match.getDescription() === base.description &&
              match.isWarningOnly() === base.warningOnly &&
              arraysEqual(editorsMatch, base.editors)
            ) {
              requiereActualizacion = false;
            } else {
              match.remove();
            }
          }

          if (requiereActualizacion) {
            try {
              var rangoDestino = hoja.getRange(base.rangeA1);
              var nuevaProteccion = rangoDestino.protect();
              nuevaProteccion.setDescription(base.description);
              nuevaProteccion.setWarningOnly(base.warningOnly);
              if (base.editors.length > 0) {
                nuevaProteccion.removeEditors(nuevaProteccion.getEditors());
                nuevaProteccion.addEditors(base.editors);
              }
            } catch (errProteccion) {
              Logger.log('Error copiando protección en hoja ' + hoja.getName() + ' (' + base.rangeA1 + '): ' + errProteccion);
            }
          }
        }

        var proteccionesRevisar = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        for (var pe = 0; pe < proteccionesRevisar.length; pe++) {
          var protExtra = proteccionesRevisar[pe];
          var keyExtra = protExtra.getRange().getA1Notation();
          var existeEnBase = false;
          for (var pbr = 0; pbr < proteccionesRefMapGlobal.length; pbr++) {
            if (proteccionesRefMapGlobal[pbr].rangeA1 === keyExtra) {
              existeEnBase = true;
              break;
            }
          }
          if (!existeEnBase) {
            protExtra.remove();
          }
        }

        // Copiar reglas de formato condicional solo si faltan
        var reglasActuales = hoja.getConditionalFormatRules();
        if (!reglasActuales || reglasActuales.length === 0) {
          hoja.setConditionalFormatRules(reglasCondRef);
        }
        
        // Celdas clave: solo rellenar si el valor es incorrecto o está vacío
        var valorNombre = hoja.getRange('C47').getValue();
        var valorNombreEsperado = tarjeta.getSheetByName('Tarjeta Ahorro').getRange('B1').getValue();
        if (valorNombre !== valorNombreEsperado) {
          hoja.getRange('C47').setValue(valorNombreEsperado);
        }
        var valorCodigo = hoja.getRange('G47').getValue();
        if (valorCodigo != codigoSocio) {
          hoja.getRange('G47').setValue(codigoSocio);
        }
      }
    } catch (e) {
      Logger.log('Error validando hojas de préstamo para socio ' + codigoSocio + ': ' + e);
    }
  }
  // ------------------------------------------------------------------------------



  var contadorAvales = 0;

  // PRIMER SENTIDO: De aval a prestatario
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
        var nombrePrestatario = hojaAhorro.getRange(fila, 3).getValue();
        var montoAvalado = hojaAhorro.getRange(fila, 4).getValue();
        if (!numeroPrestamo || !codigoPrestatario || !nombrePrestatario || !montoAvalado) continue;
        if (!sociosMap[codigoPrestatario]) continue;
        var prestatarioInfo = sociosMap[codigoPrestatario];
        var nombreCompletoAval = hojaAhorro.getRange("B1").getValue().toString();
        var nombreAval = nombreCompletoAval.split(' ')[0];
        var primerNombrePrestatario = nombrePrestatario.toString().split(' ')[0];
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
            var formulaSaldoPendiente = '=IFERROR(INDEX(FILTER(IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!H13:H23"),IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!H13:H23")<>"",IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!H13:H23")<>0),COUNTA(FILTER(IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!H13:H23"),IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!H13:H23")<>"",IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!H13:H23")<>0))),0)';
            var formulaFechaCompromiso = '=IMPORTRANGE("' + prestatarioInfo.tarjetaId + '","' + nombreHojaPrestamo + '!C5")';
            hojaAhorro.getRange(fila, 5).setFormula(formulaSaldoPendiente);
            hojaAhorro.getRange(fila, 6).setFormula(formulaFechaCompromiso);
            contadorAvales++;
            Logger.log('Aval procesado: ' + nombreAval + ' (' + codigo + ') por $' + montoAvalado + ' para préstamo #' + numeroPrestamo + ' de ' + codigoPrestatario + ' (' + primerNombrePrestatario + ')');
          }


          // Procesar avales solidarios en las últimas 5 filas de la hoja de préstamo
          var lastRowPrestamo = hojaPrestamo.getLastRow();
          for (var filaSolidario = lastRowPrestamo - 4; filaSolidario <= lastRowPrestamo; filaSolidario++) {
            var nombreSolidario = hojaPrestamo.getRange(filaSolidario, 2).getValue(); // B
            var codigoSolidario = hojaPrestamo.getRange(filaSolidario, 4).getValue(); // D
            var montoSolidario = hojaPrestamo.getRange(filaSolidario, 5).getValue(); // E
            if (nombreSolidario && codigoSolidario && montoSolidario && sociosMap[codigoSolidario]) {
              // Reflejar el monto en la tarjeta de ahorro del socio que prestó
              var tarjetaSolidario = SpreadsheetApp.openById(sociosMap[codigoSolidario].tarjetaId);
              var hojaAhorroSolidario = tarjetaSolidario.getSheetByName('Tarjeta Ahorro');
              if (hojaAhorroSolidario) {
                var lastRowSolidario = hojaAhorroSolidario.getLastRow();
                var nuevaFilaSolidario = lastRowSolidario + 1;
                hojaAhorroSolidario.getRange(nuevaFilaSolidario, 1).setValue(numeroPrestamo);
                hojaAhorroSolidario.getRange(nuevaFilaSolidario, 2).setValue(codigo);
                hojaAhorroSolidario.getRange(nuevaFilaSolidario, 3).setValue(nombrePrestatario);
                hojaAhorroSolidario.getRange(nuevaFilaSolidario, 4).setValue(montoSolidario);
                // Fórmulas
                var formulaSaldoPendienteS = '=IFERROR(INDEX(FILTER(IMPORTRANGE("' + tarjeta.getId() + '","' + nombreHojaPrestamo + '!H13:H23"),IMPORTRANGE("' + tarjeta.getId() + '","' + nombreHojaPrestamo + '!H13:H23")<>"",IMPORTRANGE("' + tarjeta.getId() + '","' + nombreHojaPrestamo + '!H13:H23")<>0),COUNTA(FILTER(IMPORTRANGE("' + tarjeta.getId() + '","' + nombreHojaPrestamo + '!H13:H23"),IMPORTRANGE("' + tarjeta.getId() + '","' + nombreHojaPrestamo + '!H13:H23")<>"",IMPORTRANGE("' + tarjeta.getId() + '","' + nombreHojaPrestamo + '!H13:H23")<>0))),0)';
                var formulaFechaCompromisoS = '=IMPORTRANGE("' + tarjeta.getId() + '","' + nombreHojaPrestamo + '!C5")';
                hojaAhorroSolidario.getRange(nuevaFilaSolidario, 5).setFormula(formulaSaldoPendienteS);
                hojaAhorroSolidario.getRange(nuevaFilaSolidario, 6).setFormula(formulaFechaCompromisoS);
                Logger.log('Solidario procesado: ' + nombreSolidario + ' (' + codigoSolidario + ') por $' + montoSolidario + ' para préstamo #' + numeroPrestamo + ' de ' + codigo + ' (' + nombrePrestatario + ')');
              }
            }
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