function llenarCondensadoAhorros() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCondensado = ss.getSheetByName('Ahorros y Retiros');
  if (!sheetCondensado) {
    Logger.log('No se encontró la hoja Ahorros y Retiros.');
    return;
  }

  // Detectar bloques de meses y sus columnas en la fila 4
  var colStart = 12; // L = 12
  var lastCol = sheetCondensado.getLastColumn();
  var mesRow = sheetCondensado.getRange(4, colStart, 1, lastCol - colStart + 1).getValues()[0];
  var semanaRow = sheetCondensado.getRange(6, colStart, 1, lastCol - colStart + 1).getValues()[0];

  // Construir bloques de meses: [{mes, colIni, colFin, semanas}]
  var bloques = [];
  var actualMes = null, actualIni = null;
  for (var j = 0; j <= semanaRow.length; j++) {
    var semanaCelda = semanaRow[j];
    // Detecta inicio de bloque por "PRIMERA"
    if ((semanaCelda && semanaCelda.toString().toUpperCase().indexOf("PRIMERA") !== -1) || j === semanaRow.length) {
      if (actualMes !== null) {
        // Busca la última columna con "QUINTA" (o la última semana válida) en el bloque actual
        var fin = j - 1;
        for (var k = j - 1; k >= actualIni; k--) {
          if (
            semanaRow[k] &&
            semanaRow[k].toString().toUpperCase().indexOf("QUINTA") !== -1
          ) {
            fin = k;
            break;
          }
        }
        bloques.push({
          mes: actualMes,
          colIni: colStart + actualIni,
          colFin: colStart + fin,
          semanas: semanaRow.slice(actualIni, fin + 1)
        });
      }
      if (j < semanaRow.length) {
        actualIni = j;
        actualMes = mesRow[j]; // El mes es el de la fila 4 en la columna donde está "PRIMERA"
      }
    }
  }

  // Obtener IDs de socios desde la columna A, desde la fila 8, omitiendo las últimas 3 filas
  var lastRow = sheetCondensado.getLastRow();
  var sociosRows = lastRow - 7 - 3; // omitir las últimas 3 filas
  if (sociosRows <= 0) return;
  var sociosData = sheetCondensado.getRange(8, 1, sociosRows, 1).getValues();

  // Buscar carpeta principal
  var parentFolder;
  try {
    var parents = DriveApp.getFileById(ss.getId()).getParents();
    parentFolder = parents.hasNext() ? parents.next() : null;
  } catch (e) {
    Logger.log('Error obteniendo carpeta principal: ' + e);
    return;
  }
  if (!parentFolder) {
    Logger.log('No se encontró la carpeta principal.');
    return;
  }

  // La carpeta principal YA contiene las carpetas de los socios
  var sociosMainFolder = parentFolder;

  for (var i = 0; i < sociosData.length; i++) {
    var socioId = sociosData[i][0];
    if (!socioId) {
      Logger.log('Fila ' + (i + 8) + ': No hay ID de socio.');
      continue;
    }

    // Buscar carpeta del socio cuyo nombre empieza con el ID y un espacio
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

    // Buscar archivo 04 TARJETA AHORRO Y PRESTAMO en la carpeta del socio usando las iniciales
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
      Logger.log('No se encontró archivo TARJETA AHORRO Y PRESTAMO para socio: ' + socioId + ' (' + iniciales + ')');
      continue;
    }

    var tarjetaId = tarjetaFile.getId();
    var tarjetaUrl = 'https://docs.google.com/spreadsheets/d/' + tarjetaId;
    // Siempre usa la hoja llamada 'Tarjeta Ahorro'
    var hojaNombre = 'Tarjeta Ahorro';

    // AHORRO INICIAL (E9) solo en columna G (7), SIN IFERROR
    try {
      var formulaAhorroInicialG = '=IMPORTRANGE("' + tarjetaUrl + '","' + hojaNombre + '!E9")';
      sheetCondensado.getRange(i + 8, 7).setFormula(formulaAhorroInicialG);
    } catch (e) {
      Logger.log('Error poniendo fórmula de ahorro inicial en columna G para socio ' + socioId + ': ' + e);
    }

    // Para cada bloque de mes y sus semanas
    for (var b = 0; b < bloques.length; b++) {
      var bloque = bloques[b];
      var mes = bloque.mes;
      var semanas = bloque.semanas;
      for (var semanaIdx = 0; semanaIdx < semanas.length; semanaIdx++) {
        try {
          var semanaNombre = semanas[semanaIdx];
          var semanaInicio = 1 + semanaIdx * 7;
          var semanaFin = semanaInicio + 6;
          // Suma C - suma D para la semana, CON IFERROR, usando rango A9:D
          var formulaSemana = '=IFERROR((SUM(QUERY(IMPORTRANGE("' + tarjetaUrl + '","' + hojaNombre + '!A9:D"), "select Col3 where Col1 >= date \'' +
            getFechaSemanaMes(mes, semanaInicio) + '\' and Col1 <= date \'' + getFechaSemanaMes(mes, semanaFin) + '\'"))' +
            '-SUM(QUERY(IMPORTRANGE("' + tarjetaUrl + '","' + hojaNombre + '!A9:D"), "select Col4 where Col1 >= date \'' +
            getFechaSemanaMes(mes, semanaInicio) + '\' and Col1 <= date \'' + getFechaSemanaMes(mes, semanaFin) + '\'"))), "")';
          sheetCondensado.getRange(i + 8, bloque.colIni + semanaIdx).setFormula(formulaSemana);
        } catch (e) {
          Logger.log('Error poniendo fórmula de semana ' + semanaNombre + ' para socio ' + socioId + ' en mes ' + mes + ': ' + e);
        }
      }
    }
    Logger.log('Fórmulas puestas para socio: ' + socioId);
  }
  Logger.log('Condensado de ahorros y préstamos llenado con IMPORTRANGE desde las carpetas de cada socio.');
}

// Devuelve fecha YYYY-MM-DD para el mes y día dados (mes en texto)
function getFechaSemanaMes(mesTexto, dia) {
  var meses = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  // Extraer mes y año del texto
  var partes = mesTexto.toString().trim().split(' ');
  var mesNombre = partes[0].toUpperCase();
  var anio = (partes.length > 1 && /^\d{4}$/.test(partes[1])) ? parseInt(partes[1]) : new Date().getFullYear();
  var idx = meses.indexOf(mesNombre);
  var mesNum = idx !== -1 ? idx + 1 : 1;

  // Calcular el último día del mes
  var ultimoDia = new Date(anio, mesNum, 0).getDate();
  var diaValido = Math.min(dia, ultimoDia);

  mesNum = ('0' + mesNum).slice(-2);
  diaValido = ('0' + diaValido).slice(-2);
  return anio + '-' + mesNum + '-' + diaValido;
}