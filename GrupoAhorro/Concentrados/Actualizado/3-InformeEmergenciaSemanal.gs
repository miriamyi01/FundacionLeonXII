function llenarCondensadoEmergencias() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCondensado = ss.getSheetByName('Fondo de Emergencia');
  if (!sheetCondensado) {
    Logger.log('No se encontró la hoja Fondo de Emergencia.');
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

  // Nombres de carpetas para buscar
  var rootFolderName = 'GA0452 METAMORFOSIS';
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

  // Buscar la carpeta de socios dentro de la carpeta raíz
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

  // Obtener IDs de socios desde la columna A, desde la fila 4 (después de encabezados)
  var lastRow = sheetCondensado.getLastRow();
  var socioStartRow = 4;
  var sociosRows = lastRow - socioStartRow + 1;
  if (sociosRows <= 0) {
    Logger.log('No hay filas de socios para procesar.');
    return;
  }
  var sociosData = sheetCondensado.getRange(socioStartRow, 1, sociosRows, 2).getValues();
  var sociosInfo = [];
  for (var sr = 0; sr < sociosData.length; sr++) {
    var socioIdRaw = sociosData[sr][0];
    if (!socioIdRaw) continue;
    sociosInfo.push({
      row: socioStartRow + sr,
      socioId: socioIdRaw,
      nombre: String(sociosData[sr][1] || '').trim()
    });
  }

  sociosInfo.sort(function(a, b) {
    var cmp = a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' });
    if (cmp !== 0) return cmp;
    return String(a.socioId).localeCompare(String(b.socioId), 'es', { sensitivity: 'base' });
  });

  if (sociosInfo.length === 0) {
    Logger.log('No hay socios válidos para procesar.');
    return;
  }

  // Recolectar todas las fechas de todas las tarjetas de emergencia
  var todasLasFechas = new Set();
  var arregloFechas = [];

  for (var i = 0; i < sociosInfo.length; i++) {
    var socioId = sociosInfo[i].socioId;
    if (!socioId) continue;

    // Buscar carpeta del socio
    var socioFolder = null;
    var socioFoldersIter = sociosMainFolder.getFolders();
    while (socioFoldersIter.hasNext()) {
      var folder = socioFoldersIter.next();
      if (folder.getName().indexOf(socioId + " ") === 0) {
        socioFolder = folder;
        break;
      }
    }
    if (!socioFolder) {
      Logger.log('No se encontró carpeta para socio: ' + socioId);
      continue;
    }

    // Buscar archivo TARJETA AHORRO Y PRESTAMO
    var carpetaNombre = socioFolder.getName();
    var partes = carpetaNombre.split(" ");
    var iniciales = partes.length > 1 ? partes[1] : "";
    var tarjetaFiles = socioFolder.getFiles();
    var tarjetaFile = null;
    while (tarjetaFiles.hasNext()) {
      var tf = tarjetaFiles.next();
      if (tf.getName().indexOf(iniciales) !== -1 && tf.getName().indexOf('TARJETA AHORRO Y PRESTAMO') !== -1) {
        tarjetaFile = tf;
        break;
      }
    }
    if (!tarjetaFile) {
      Logger.log('No se encontró archivo TARJETA AHORRO Y PRESTAMO para socio: ' + socioId);
      continue;
    }

    var tarjetaId = tarjetaFile.getId();
    var tarjeta;
    try {
      tarjeta = SpreadsheetApp.openById(tarjetaId);
    } catch (e) {
      Logger.log('Error abriendo archivo para socio ' + socioId + ': ' + e);
      continue;
    }

    // Obtener fechas de la hoja "Fondo de Emergencia" rango A4:A
    try {
      var tarjetaEmergencia = tarjeta.getSheetByName('Fondo de Emergencia');
      if (tarjetaEmergencia) {
        var datosAhorro = tarjetaEmergencia.getRange('A4:A').getValues();
        for (var j = 0; j < datosAhorro.length; j++) {
          var fecha = datosAhorro[j][0];
          if (fecha && fecha instanceof Date) {
            var fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            if (!todasLasFechas.has(fechaStr)) {
              todasLasFechas.add(fechaStr);
              arregloFechas.push(new Date(fecha));
            }
          }
        }
      }
    } catch (e) {
      Logger.log('Error extrayendo fechas de Fondo de Emergencia para socio ' + socioId + ': ' + e);
    }
  }

  if (arregloFechas.length === 0) {
    Logger.log('No se encontraron fechas en las tarjetas de ahorro.');
    return;
  }

  // Ordenar fechas y agrupar por semanas (lunes a domingo)
  arregloFechas.sort(function(a, b) { return a - b; });
  
  var semanas = []; // Array de {lunesFecha: Date, domingoFecha: Date}
  for (var i = 0; i < arregloFechas.length; i++) {
    var fecha = arregloFechas[i];
    var lunesFecha = obtenerLunesDelaSemana(fecha);
    var domingoFecha = new Date(lunesFecha);
    domingoFecha.setDate(domingoFecha.getDate() + 6);

    var semanaExistente = false;
    for (var s = 0; s < semanas.length; s++) {
      if (Utilities.formatDate(semanas[s].lunesFecha, Session.getScriptTimeZone(), 'yyyy-MM-dd') === 
          Utilities.formatDate(lunesFecha, Session.getScriptTimeZone(), 'yyyy-MM-dd')) {
        semanaExistente = true;
        break;
      }
    }

    if (!semanaExistente) {
      semanas.push({
        lunesFecha: lunesFecha,
        domingoFecha: domingoFecha
      });
    }
  }

  // Detectar qué semanas ya tienen fechas en la fila 3
  var lastCol = sheetCondensado.getLastColumn();
  var semanasExistentes = new Set();
  var columnasSemanaExistente = {};
  var colStart = 6; // Columna F = 6
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

  Logger.log('Semanas encontradas: ' + semanas.length + ', Semanas existentes: ' + semanasExistentes.size);

  // Llenar por semana: fila 3 (fechas del lunes), fila 2 (suma), fila 4+ (ImportRange por socio)
  var colActual = colStart;
  
  for (var s = 0; s < semanas.length; s++) {
    var semana = semanas[s];
    var fechaLunesStr = Utilities.formatDate(semana.lunesFecha, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var fechaDomingoStr = Utilities.formatDate(semana.domingoFecha, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var semanaYaExiste = semanasExistentes.has(fechaLunesStr);
    var colDestino = semanaYaExiste ? columnasSemanaExistente[fechaLunesStr] : colActual;

    if (!semanaYaExiste) {
      // Poner la fecha del lunes en fila 3
      var fechaLunesDisplay = Utilities.formatDate(semana.lunesFecha, 'UTC', 'dd/MM/yyyy');
      sheetCondensado.getRange(3, colDestino).setValue(fechaLunesDisplay);

      // Poner fórmula de suma en fila 2
      var columnaLetra = String.fromCharCode(64 + colDestino); // Convertir número de columna a letra
      sheetCondensado.getRange(2, colDestino).setFormula('=SUM(' + columnaLetra + '4:' + columnaLetra + ')');

      // Copiar reglas de formato condicional de la columna F a la nueva columna
      copiarCondicionalesColumna(sheetCondensado, colStart, colDestino);
    } else {
      Logger.log('Semana ' + fechaLunesStr + ' ya existe, rellenando faltantes.');
    }

    // Para cada socio, poner ImportRange de la tarjeta de préstamo para esa semana
    for (var i = 0; i < sociosInfo.length; i++) {
      var socio = sociosInfo[i];
      var socioId = socio.socioId;
      if (!socioId) continue;

      // Buscar carpeta del socio
      var socioFolder = null;
      var socioFoldersIter = sociosMainFolder.getFolders();
      while (socioFoldersIter.hasNext()) {
        var folder = socioFoldersIter.next();
        if (folder.getName().indexOf(socioId + " ") === 0) {
          socioFolder = folder;
          break;
        }
      }
      if (!socioFolder) continue;

      // Buscar archivo TARJETA AHORRO Y PRESTAMO
      var carpetaNombre = socioFolder.getName();
      var partes = carpetaNombre.split(" ");
      var iniciales = partes.length > 1 ? partes[1] : "";
      var tarjetaFiles = socioFolder.getFiles();
      var tarjetaFile = null;
      while (tarjetaFiles.hasNext()) {
        var tf = tarjetaFiles.next();
        if (tf.getName().indexOf(iniciales) !== -1 && tf.getName().indexOf('TARJETA AHORRO Y PRESTAMO') !== -1) {
          tarjetaFile = tf;
          break;
        }
      }
      if (!tarjetaFile) continue;

      var tarjetaId = tarjetaFile.getId();
      var tarjetaUrl = 'https://docs.google.com/spreadsheets/d/' + tarjetaId;

      // Suma de aportaciones (C) en esa semana desde Fondo de Emergencia usando QUERY
      var fechaLunesFormatted = Utilities.formatDate(semana.lunesFecha, 'UTC', 'yyyy-MM-dd');
      var fechaDomingoFormatted = Utilities.formatDate(semana.domingoFecha, 'UTC', 'yyyy-MM-dd');
      
      var formulaBase = 'SUM(QUERY(IMPORTRANGE("' + tarjetaUrl + '","\'Fondo de Emergencia\'!A4:D"), "select Col3 where Col1 >= date \'' +
        fechaLunesFormatted + '\' and Col1 <= date \'' + fechaDomingoFormatted + '\'", 0))';

      var formulaSemana = '=IFERROR(' + formulaBase + ', "")';

      var celdaDestino = sheetCondensado.getRange(socio.row, colDestino);
      if (semanaYaExiste) {
        if (!celdaDestino.getFormula() && celdaDestino.getValue() === '') {
          celdaDestino.setFormula('=IFERROR(' + formulaBase + ', "")');
        }
      } else {
        celdaDestino.setFormula(formulaSemana);
      }
    }

    colActual++;
  }

  Logger.log('Condensado de fondo de emergencia por semanas completado. Columnas procesadas: ' + (colActual - colStart));
}