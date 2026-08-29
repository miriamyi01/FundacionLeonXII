function llenarCondensadoAhorros() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCondensado = ss.getSheetByName('Ahorros y Retiros');
  if (!sheetCondensado) {
    Logger.log('No se encontró la hoja Ahorros y Retiros.');
    return;
  }

  function copiarCondicionalesColumna(sheet, fromColumn, toColumn) {
    if (fromColumn === toColumn) return;

    var reglas = sheet.getConditionalFormatRules();
    var nuevasReglas = [];

    for (var i = 0; i < reglas.length; i++) {
      var regla = reglas[i];
      var rangos = regla.getRanges();
      var rangosExtendidos = rangos.slice();

      for (var r = 0; r < rangos.length; r++) {
        var rango = rangos[r];
        if (rango.getColumn() === fromColumn && rango.getNumColumns() === 1) {
          rangosExtendidos.push(
            sheet.getRange(rango.getRow(), toColumn, rango.getNumRows(), 1)
          );
        }
      }

      if (rangosExtendidos.length !== rangos.length) {
        regla = regla.copy().setRanges(rangosExtendidos).build();
      }

      nuevasReglas.push(regla);
    }

    sheet.setConditionalFormatRules(nuevasReglas);
  }

  function asegurarIfErrorEnColumna(sheet, startRow, numRows, column) {
    if (numRows <= 0) return;

    var rango = sheet.getRange(startRow, column, numRows, 1);
    var formulas = rango.getFormulas();
    var huboCambios = false;

    for (var i = 0; i < formulas.length; i++) {
      var formula = formulas[i][0];
      if (!formula) continue;

      if (/^\s*=\s*IFERROR\s*\(/i.test(formula)) continue;

      var cuerpo = formula.charAt(0) === '=' ? formula.substring(1) : formula;
      formulas[i][0] = '=IFERROR(' + cuerpo + ', 0)';
      huboCambios = true;
    }

    if (huboCambios) {
      rango.setFormulas(formulas);
    }
  }

  function convertirTextoAClaveFecha(texto) {
    var limpio = String(texto).trim();

    var m1 = limpio.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m1) {
      var dia = ('0' + m1[1]).slice(-2);
      var mes = ('0' + m1[2]).slice(-2);
      var anio = m1[3];
      return anio + '-' + mes + '-' + dia;
    }

    var m2 = limpio.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m2) {
      return m2[1] + '-' + ('0' + m2[2]).slice(-2) + '-' + ('0' + m2[3]).slice(-2);
    }

    return null;
  }

  function convertirUtcAFechaLocal(fechaUtc) {
    // No usar new Date(anio, mes, dia) aca porque toma el huso horario del runtime y puede desfasar el dia
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var fechaStr = Utilities.formatDate(fechaUtc, 'UTC', 'yyyy-MM-dd');
    return Utilities.parseDate(fechaStr, tz, 'yyyy-MM-dd');
  }

  function obtenerLunesDelaSemana(fecha) {
    var tz = Session.getScriptTimeZone();
    var ymd = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd');
    var partes = ymd.split('-');
    var fechaUtc = new Date(Date.UTC(parseInt(partes[0], 10), parseInt(partes[1], 10) - 1, parseInt(partes[2], 10)));
    var dia = fechaUtc.getUTCDay();
    var diferenciaAlunes = (dia === 0) ? -6 : 1 - dia;
    fechaUtc.setUTCDate(fechaUtc.getUTCDate() + diferenciaAlunes);
    return fechaUtc;
  }

  function convertirNumeroAColumnaLetra(columna) {
    var letra = '', temp;
    while (columna > 0) {
      temp = (columna - 1) % 26;
      letra = String.fromCharCode(65 + temp) + letra;
      columna = (columna - temp - 1) / 26;
    }
    return letra;
  }

  // Mueve las filas tardías de un socio a la siguiente semana libre en su tarjeta y en el condensado
  function moverAhorroASemanaSiguiente(socio, filasTardias, lunesOrigen, sheetCondensado, colStart, colCursorRef, semanasExistentes, columnasSemanaExistente) {
    var tarjeta = socio.tarjetaSpreadsheet || SpreadsheetApp.openById(socio.tarjetaId);
    var hojaAhorro = tarjeta.getSheetByName('Tarjeta Ahorro');

    var lunesDestino = new Date(lunesOrigen);
    var colDestino, fechaDestinoStr, celda, libre;

    do {
      lunesDestino = new Date(Date.UTC(lunesDestino.getUTCFullYear(), lunesDestino.getUTCMonth(), lunesDestino.getUTCDate() + 7));
      fechaDestinoStr = Utilities.formatDate(lunesDestino, 'UTC', 'yyyy-MM-dd');

      if (semanasExistentes.has(fechaDestinoStr)) {
        colDestino = columnasSemanaExistente[fechaDestinoStr]
      } else {
        colDestino = colCursorRef.valor
        var fechaLunesDisplay = Utilities.formatDate(lunesDestino, 'UTC', 'dd/MM/yyyy')
        sheetCondensado.getRange(3, colDestino).setValue(fechaLunesDisplay)
        var columnaLetra = convertirNumeroAColumnaLetra(colDestino)
        sheetCondensado.getRange(2, colDestino).setFormula('=SUM(' + columnaLetra + '4:' + columnaLetra + ')')
        copiarCondicionalesColumna(sheetCondensado, colStart, colDestino)
        semanasExistentes.add(fechaDestinoStr)
        columnasSemanaExistente[fechaDestinoStr] = colDestino
        colCursorRef.valor = colDestino + 1
      }

      // Semana futura: se llena con formula viva solo a quien tenga algo real, al resto se lo deja vacío
      var domingoDestinoRelleno = new Date(lunesDestino)
      domingoDestinoRelleno.setUTCDate(domingoDestinoRelleno.getUTCDate() + 6)
      var domingoDestinoRellenoStr = Utilities.formatDate(domingoDestinoRelleno, 'UTC', 'yyyy-MM-dd')
      for (var k = 0; k < sociosInfo.length; k++) {
        var otroSocio = sociosInfo[k]
        if (otroSocio.row === socio.row) continue
        var celdaOtroExistente = sheetCondensado.getRange(otroSocio.row, colDestino)
        if (celdaOtroExistente.getFormula()) continue
        var datosOtro = otroSocio.datosAhorro || []
        var sumaOtro = 0
        for (var rr = 0; rr < datosOtro.length; rr++) {
          var fOtro = datosOtro[rr][0]
          if (fOtro instanceof Date) {
            var lunesOtro = obtenerLunesDelaSemana(fOtro)
            if (Utilities.formatDate(lunesOtro, 'UTC', 'yyyy-MM-dd') === fechaDestinoStr) {
              sumaOtro += (Number(datosOtro[rr][2]) || 0) - (Number(datosOtro[rr][3]) || 0)
            }
          }
        }
        if (sumaOtro === 0) continue
        var querySumaOtro = 'select Col3 where Col1 >= date \'' + fechaDestinoStr + '\' and Col1 <= date \'' + domingoDestinoRellenoStr + '\''
        var queryRetirosOtro = 'select Col4 where Col1 >= date \'' + fechaDestinoStr + '\' and Col1 <= date \'' + domingoDestinoRellenoStr + '\''
        var formulaOtro = '(SUM(QUERY(IMPORTRANGE("' + otroSocio.tarjetaUrl + '","\'Tarjeta Ahorro\'!A12:D54"), "' + querySumaOtro + '", 0))' +
          '-SUM(QUERY(IMPORTRANGE("' + otroSocio.tarjetaUrl + '","\'Tarjeta Ahorro\'!A12:D54"), "' + queryRetirosOtro + '", 0)))'
        celdaOtroExistente.setFormula('=IFERROR(' + formulaOtro + ', 0)')
      }

      celda = sheetCondensado.getRange(socio.row, colDestino);
      libre = !celda.getFormula() && (celda.getValue() === '' || celda.getValue() === 0);
    } while (!libre);

    var lunesDestinoStr = fechaDestinoStr;
    for (var f = 0; f < filasTardias.length; f++) {
      var indexMemoria = filasTardias[f] - 12;
      var fechaDestinoLocal = convertirUtcAFechaLocal(lunesDestino);
      hojaAhorro.getRange(filasTardias[f], 1).setValue(fechaDestinoLocal);
      // Guardar la misma fecha corregida que se escribió en la celda, no la cruda en UTC
      if (socio.datosAhorro) socio.datosAhorro[indexMemoria][0] = fechaDestinoLocal;
    }

    var domingoDestinoUTC = new Date(lunesDestino);
    domingoDestinoUTC.setUTCDate(domingoDestinoUTC.getUTCDate() + 6);
    var domingoDestinoStr = Utilities.formatDate(domingoDestinoUTC, 'UTC', 'yyyy-MM-dd');
    var rangoAhorro = "'Tarjeta Ahorro'!A12:D54";
    var querySuma = 'select Col3 where Col1 >= date \'' + lunesDestinoStr + '\' and Col1 <= date \'' + domingoDestinoStr + '\'';
    var queryRetiros = 'select Col4 where Col1 >= date \'' + lunesDestinoStr + '\' and Col1 <= date \'' + domingoDestinoStr + '\'';
    var formulaBase = '(SUM(QUERY(IMPORTRANGE("' + socio.tarjetaUrl + '","' + rangoAhorro + '"), "' + querySuma + '", 0))' +
      '-SUM(QUERY(IMPORTRANGE("' + socio.tarjetaUrl + '","' + rangoAhorro + '"), "' + queryRetiros + '", 0)))';

    celda.setFormula('=IFERROR(' + formulaBase + ', 0)');
  }

  var rootFolderName = 'GA0452 METAMORFOSIS';
  var sociosFolderName = 'GA0452-SOCIOS AS';

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

  var sociosMainFolder = null;
  var subFolders = rootFolder.getFolders();
  while (subFolders.hasNext()) {
    var sf = subFolders.next();
    if (sf.getName().indexOf(sociosFolderName) !== -1) {
      sociosMainFolder = sf;
      break;
    }
  }
  if (!sociosMainFolder) {
    Logger.log('No se encontró la carpeta ' + sociosFolderName + ' dentro de ' + rootFolderName);
    return;
  }

  var lastRow = sheetCondensado.getLastRow();
  var socioStartRow = 4;
  var sociosRows = lastRow - socioStartRow + 1;
  if (sociosRows <= 0) {
    Logger.log('No hay filas de socios para procesar.');
    return;
  }

  var sociosData = sheetCondensado.getRange(socioStartRow, 1, sociosRows, 2).getValues();
  
  var foldersIter = sociosMainFolder.getFolders();
  var folderMap = {};
  while (foldersIter.hasNext()) {
    var f = foldersIter.next();
    var id = String(f.getName().split(" ")[0]).trim();
    folderMap[id] = f;
  }

  var sociosInfo = [];
  for (var sr = 0; sr < sociosData.length; sr++) {
    var socioId = String(sociosData[sr][0]).trim();
    if (!socioId || !folderMap[socioId]) continue;

    var socioFolder = folderMap[socioId];
    var iniciales = socioFolder.getName().split(" ")[1] || "";
    var tarjetaFiles = socioFolder.getFiles();
    var tarjetaFile = null;
    
    while (tarjetaFiles.hasNext()) {
      var tf = tarjetaFiles.next();
      if (tf.getName().indexOf(iniciales) !== -1 && tf.getName().indexOf('TARJETA AHORRO Y PRESTAMO') !== -1) {
        tarjetaFile = tf;
        break;
      }
    }

    if (tarjetaFile) {
      sociosInfo.push({
        row: socioStartRow + sr,
        socioId: socioId,
        nombre: String(sociosData[sr][1] || '').trim(),
        tarjetaId: tarjetaFile.getId(),
        tarjetaUrl: 'https://docs.google.com/spreadsheets/d/' + tarjetaFile.getId()
      });
    }
  }

  if (sociosInfo.length === 0) {
    Logger.log('No hay socios válidos para procesar.');
    return;
  }

  var todasLasFechas = new Set();
  var arregloFechas = [];
  var fechasConMontoReal = []; // solo fechas con ahorro o retiro real, no filas de plantilla vacías

  for (var i = 0; i < sociosInfo.length; i++) {
    try {
      var tarjeta = SpreadsheetApp.openById(sociosInfo[i].tarjetaId);
      sociosInfo[i].tarjetaSpreadsheet = tarjeta; // se reutiliza más abajo, evita reabrirla
      var tarjetaAhorro = tarjeta.getSheetByName('Tarjeta Ahorro');
      if (tarjetaAhorro) {
        var datosAhorro = tarjetaAhorro.getRange('A12:D54').getValues();
        for (var j = 0; j < datosAhorro.length; j++) {
          var fecha = datosAhorro[j][0];
          if (fecha && fecha instanceof Date) {
            var fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            if (!todasLasFechas.has(fechaStr)) {
              todasLasFechas.add(fechaStr);
              arregloFechas.push(new Date(fecha));
            }
            var montoFila = (Number(datosAhorro[j][2]) || 0) - (Number(datosAhorro[j][3]) || 0);
            if (montoFila !== 0) fechasConMontoReal.push(new Date(fecha));
          }
        }
      }
      sociosInfo[i].datosAhorro = datosAhorro;
    } catch (e) {
      Logger.log('Error extrayendo fechas de Tarjeta Ahorro para socio ' + socioId + ': ' + e);
    }
  }

  if (arregloFechas.length === 0) {
    Logger.log('No se encontraron fechas en las tarjetas de ahorro.');
    return;
  }

  arregloFechas.sort(function(a, b) { return a - b; });
  
  var hoy = new Date();
  var lunesDeEstaSemana = obtenerLunesDelaSemana(hoy);
  var lunesDeEstaSemanaStr = Utilities.formatDate(lunesDeEstaSemana, 'UTC', 'yyyy-MM-dd');

  var lastCol = sheetCondensado.getLastColumn();
  var semanasExistentes = new Set();
  var columnasSemanaExistente = {};
  var colStart = 10; // Columna J = 10
  var valorF3Inicial = sheetCondensado.getRange(3, colStart).getValue();
  var primeraEjecucionGlobal = !valorF3Inicial;
  if (lastCol >= colStart) { 
    var fila3 = sheetCondensado.getRange(3, colStart, 1, lastCol - colStart + 1).getValues()[0];
    for (var c = 0; c < fila3.length; c++) {
      var valor = fila3[c];
      if (valor instanceof Date) {
        var fechaStr = Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        semanasExistentes.add(fechaStr);
        columnasSemanaExistente[fechaStr] = colStart + c;
      } else if (typeof valor === 'string' && valor.trim() !== '') {
        var fechaStrTexto = convertirTextoAClaveFecha(valor);
        if (fechaStrTexto) {
          semanasExistentes.add(fechaStrTexto);
          columnasSemanaExistente[fechaStrTexto] = colStart + c;
        }
      }
    }
  }

  Logger.log('Semanas existentes en la hoja: ' + semanasExistentes.size);

  // Recorre semana a semana desde la más antigua hasta la actual, sin huecos
  var candidatosLunesMinimo = []
  if (arregloFechas.length > 0) candidatosLunesMinimo.push(obtenerLunesDelaSemana(arregloFechas[0]))
  semanasExistentes.forEach(function(fechaStr) {
    var partes = fechaStr.split('-')
    candidatosLunesMinimo.push(new Date(Date.UTC(parseInt(partes[0], 10), parseInt(partes[1], 10) - 1, parseInt(partes[2], 10))))
  })

  var semanas = []
  if (candidatosLunesMinimo.length > 0) {
    var lunesMinimo = candidatosLunesMinimo.reduce(function(min, d) { return d < min ? d : min })
    var lunesCursor = new Date(Date.UTC(lunesMinimo.getUTCFullYear(), lunesMinimo.getUTCMonth(), lunesMinimo.getUTCDate()))

    var limiteFinal = new Date(lunesDeEstaSemana.getTime())
    // El límite se extiende más allá de la semana actual si ya hay un ahorro adelantado o una columna futura creada
    for (var af = 0; af < fechasConMontoReal.length; af++) {
      var lunesDeEsaFecha = obtenerLunesDelaSemana(fechasConMontoReal[af])
      if (lunesDeEsaFecha > limiteFinal) limiteFinal = lunesDeEsaFecha
    }
    semanasExistentes.forEach(function(fechaStr) {
      var partes = fechaStr.split('-')
      var fechaExistente = new Date(Date.UTC(parseInt(partes[0], 10), parseInt(partes[1], 10) - 1, parseInt(partes[2], 10)))
      if (fechaExistente > limiteFinal) limiteFinal = fechaExistente
    })

    while (lunesCursor <= limiteFinal) {
      var domingoCursor = new Date(lunesCursor)
      domingoCursor.setUTCDate(domingoCursor.getUTCDate() + 6)
      semanas.push({
        lunesFecha: new Date(lunesCursor),
        domingoFecha: domingoCursor
      })
      lunesCursor = new Date(Date.UTC(lunesCursor.getUTCFullYear(), lunesCursor.getUTCMonth(), lunesCursor.getUTCDate() + 7))
    }
  }

  var allValues = [];
  var allFormulas = [];
  if (lastCol >= colStart && sociosRows > 0) {
    var dataRange = sheetCondensado.getRange(socioStartRow, colStart, sociosRows, lastCol - colStart + 1);
    allValues = dataRange.getValues();
    allFormulas = dataRange.getFormulas();
  }

  var colCursor = colStart;
  var colCursorRef = { valor: colStart }; // referencia mutable, la usa moverAhorroASemanaSiguiente

  var semanasProcesadas = 0;
  for (var s = 0; s < semanas.length; s++) {
    var semana = semanas[s];
    var fechaLunesStr = Utilities.formatDate(semana.lunesFecha, 'UTC', 'yyyy-MM-dd');
    var fechaDomingoStr = Utilities.formatDate(semana.domingoFecha, 'UTC', 'yyyy-MM-dd');
    var semanaYaExiste = semanasExistentes.has(fechaLunesStr);
    var colDestino;

    if (semanaYaExiste) {
      colDestino = columnasSemanaExistente[fechaLunesStr];
      colCursor = colDestino + 1;
    } else {
      colDestino = colCursor;
      if (colDestino <= lastCol) {
        sheetCondensado.insertColumnBefore(colDestino);
        lastCol++;
        for (var claveFecha in columnasSemanaExistente) {
          if (columnasSemanaExistente[claveFecha] >= colDestino) {
            columnasSemanaExistente[claveFecha]++;
          }
        }
        var idxInsercion = colDestino - colStart;
        for (var fIns = 0; fIns < allValues.length; fIns++) {
          allValues[fIns].splice(idxInsercion, 0, '');
          allFormulas[fIns].splice(idxInsercion, 0, '');
        }
      } else {
        lastCol = colDestino;
      }
      semanasExistentes.add(fechaLunesStr);
      columnasSemanaExistente[fechaLunesStr] = colDestino;
      colCursor = colDestino + 1;
    }
    colCursorRef.valor = colCursor;

    // Solo la semana actual es dinámica, las pasadas ya congeladas solo se revisan por atrasados
    semanasProcesadas++;

    if (!semanaYaExiste) {
      Logger.log('Creando nueva columna para la semana del ' + fechaLunesStr);
      var fechaLunesDisplay = Utilities.formatDate(semana.lunesFecha, 'UTC', 'dd/MM/yyyy');
      sheetCondensado.getRange(3, colDestino).setValue(fechaLunesDisplay);

      var columnaLetra = convertirNumeroAColumnaLetra(colDestino);
      sheetCondensado.getRange(2, colDestino).setFormula('=SUM(' + columnaLetra + '4:' + columnaLetra + ')');

      copiarCondicionalesColumna(sheetCondensado, colStart, colDestino);
    } else {
      Logger.log('Procesando semana existente: ' + fechaLunesStr);
    }

    var idxColBloque = colDestino - colStart;

    // Columna completa en memoria desde lo que ya había, evitando leer/escribir celda por celda
    var columnaSalida = [];
    for (var fila = 0; fila < sociosRows; fila++) {
      var yaTenia = semanaYaExiste && allFormulas[fila] && idxColBloque >= 0 && idxColBloque < allFormulas[fila].length;
      columnaSalida.push([yaTenia ? (allFormulas[fila][idxColBloque] || allValues[fila][idxColBloque]) : '']);
    }

    for (var i = 0; i < sociosInfo.length; i++) {
      var socio = sociosInfo[i];
      var idxFila = socio.row - socioStartRow;
      var datosAhorro = socio.datosAhorro || [];
      var sumSemana = 0;
      var entradasSemana = []; // { fila, monto }, en orden de aparición en la tarjeta
      for (var r = 0; r < datosAhorro.length; r++) {
        var fRow = datosAhorro[r][0];
        if (fRow instanceof Date) {
          var fLunesRow = obtenerLunesDelaSemana(fRow);
          if (Utilities.formatDate(fLunesRow, 'UTC', 'yyyy-MM-dd') === fechaLunesStr) {
            var ahorro = Number(datosAhorro[r][2]) || 0;
            var retiro = Number(datosAhorro[r][3]) || 0;
            var monto = ahorro - retiro;
            sumSemana += monto;
            if (monto !== 0) entradasSemana.push({ fila: r + 12, monto: monto });
          }
        }
      }

      var fechaLunesFormatted = Utilities.formatDate(semana.lunesFecha, 'UTC', 'yyyy-MM-dd');
      var fechaDomingoFormatted = Utilities.formatDate(semana.domingoFecha, 'UTC', 'yyyy-MM-dd');
      var formulaBase = '(SUM(QUERY(IMPORTRANGE("' + socio.tarjetaUrl + '","\'Tarjeta Ahorro\'!A12:D54"), "select Col3 where Col1 >= date \'' +
        fechaLunesFormatted + '\' and Col1 <= date \'' + fechaDomingoFormatted + '\'", 0))' +
        '-SUM(QUERY(IMPORTRANGE("' + socio.tarjetaUrl + '","\'Tarjeta Ahorro\'!A12:D54"), "select Col4 where Col1 >= date \'' +
        fechaLunesFormatted + '\' and Col1 <= date \'' + fechaDomingoFormatted + '\'", 0)))';

      var formulaProtegida = '=IFERROR(' + formulaBase + ', 0)';

      var formulaActual = '';
      var valorActual = '';
      if (semanaYaExiste && allFormulas[idxFila] && idxColBloque >= 0 && idxColBloque < allFormulas[idxFila].length) {
        formulaActual = allFormulas[idxFila][idxColBloque];
        valorActual = allValues[idxFila][idxColBloque];
      }

      var esUltimaSemana = (semana.lunesFecha >= lunesDeEstaSemana)
      // True para la semana actual o cualquier futura ya procesada, para no congelarla antes de tiempo
      var celdaVacia = !formulaActual && (valorActual === '' || valorActual === null || valorActual === undefined);

      // Semana futura sin movimiento real no lleva formula viva, se deja vacía para llenarse bien cuando llegue
      var esFuturaEstricta = semana.lunesFecha > lunesDeEstaSemana;
      var sinMovimientoReal = sumSemana === 0;

      var resultado;
      if (!semanaYaExiste) {
        if (esFuturaEstricta && sinMovimientoReal) {
          resultado = '';
        } else {
          // Columna nueva: siempre fórmula viva, nunca un 0 literal para no marcar la semana como cerrada
          resultado = formulaProtegida;
        }
      } else if (celdaVacia) {
        if (esFuturaEstricta && sinMovimientoReal) {
          resultado = '';
        } else {
          // Celda vacía: se rellena siempre con fórmula viva
          resultado = formulaProtegida;
        }
      } else {
        // Ya tiene algo puesto: se compara contra lo real de la tarjeta, sea semana actual o pasada
        var valorCongelado = Number(valorActual) || 0;
        if (sumSemana !== valorCongelado) {
          // Depósito tardío: solo se mueven las entradas que exceden lo ya puesto
          var acumulado = 0;
          var indexCorte;
          if (Math.abs(acumulado - valorCongelado) < 0.0001) {
            // Lo que había puesto era 0, todas las entradas son tardías
            indexCorte = 0;
          } else {
            indexCorte = entradasSemana.length;
            for (var idxEntrada = 0; idxEntrada < entradasSemana.length; idxEntrada++) {
              acumulado += entradasSemana[idxEntrada].monto;
              if (Math.abs(acumulado - valorCongelado) < 0.0001) {
                indexCorte = idxEntrada + 1;
                break;
              }
            }
          }
          var filasTardias = entradasSemana.slice(indexCorte).map(function(e) { return e.fila; });
          if (filasTardias.length > 0) {
            try {
              moverAhorroASemanaSiguiente(socio, filasTardias, semana.lunesFecha, sheetCondensado, colStart, colCursorRef, semanasExistentes, columnasSemanaExistente);
              colCursor = colCursorRef.valor;
            } catch (e) {
              Logger.log('Error moviendo fecha para socio ' + socio.socioId + ': ' + e);
            }
          }
        }
        if (formulaActual) {
          // Fórmula viva: se recalcula sola porque la fecha del depósito tardío ya cambió
          resultado = esUltimaSemana ? formulaActual : valorActual;
        } else {
          // Conserva su valor congelado, lo tardío se movió a la semana siguiente libre
          resultado = valorActual;
        }
      }

      columnaSalida[idxFila] = [resultado];
    }

    // Se escribe la columna completa de esta semana de una sola vez
    sheetCondensado.getRange(socioStartRow, colDestino, sociosRows, 1).setValues(columnaSalida);
  }

  if (!primeraEjecucionGlobal) {
    asegurarIfErrorEnColumna(sheetCondensado, socioStartRow, sociosRows, colStart);
  }

  Logger.log('Condensado de ahorros completado. Semanas procesadas: ' + semanasProcesadas);
}