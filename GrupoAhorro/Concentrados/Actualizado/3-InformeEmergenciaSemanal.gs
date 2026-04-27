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

  // Obtener IDs de socios y mapear información de Drive de forma eficiente
  var lastRow = sheetCondensado.getLastRow();
  var socioStartRow = 4;
  var sociosRows = lastRow - socioStartRow + 1;
  if (sociosRows <= 0) {
    Logger.log('No hay filas de socios para procesar.');
    return;
  }

  var sociosData = sheetCondensado.getRange(socioStartRow, 1, sociosRows, 2).getValues();
  
  // Cargar todas las carpetas de socios una sola vez para evitar lentitud
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
      var tarjetaId = tarjetaFile.getId();
      var gidEmergencia = "";
      try {
        var tempSs = SpreadsheetApp.openById(tarjetaId);
        var shE = tempSs.getSheetByName('Fondo de Emergencia');
        if (shE) gidEmergencia = shE.getSheetId();
      } catch(e) {}

      sociosInfo.push({
        row: socioStartRow + sr,
        socioId: socioId,
        nombre: String(sociosData[sr][1] || '').trim(),
        tarjetaId: tarjetaId,
        tarjetaUrl: 'https://docs.google.com/spreadsheets/d/' + tarjetaId,
        gidEmergencia: gidEmergencia
      });
    }
  }

  if (sociosInfo.length === 0) return Logger.log('No se encontraron tarjetas válidas.');

  // Recolectar todas las fechas de todas las tarjetas (ahora optimizado)
  var todasLasFechas = new Set();
  var arregloFechas = [];
  for (var i = 0; i < sociosInfo.length; i++) {
    try {
      var tarjeta = SpreadsheetApp.openById(sociosInfo[i].tarjetaId);
      var tarjetaEmergencia = tarjeta.getSheetByName('Fondo de Emergencia');
      if (tarjetaEmergencia) {
        var datosAhorro = tarjetaEmergencia.getRange('A4:A25').getValues();
        for (var j = 0; j < datosAhorro.length; j++) {
          var f = datosAhorro[j][0];
          if (f && f instanceof Date) {
            var fStr = Utilities.formatDate(f, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            if (!todasLasFechas.has(fStr)) {
              todasLasFechas.add(fStr);
              arregloFechas.push(new Date(f));
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

    // Para cada socio, actualizar nombre/enlace y poner fórmulas
    for (var i = 0; i < sociosInfo.length; i++) {
      var socio = sociosInfo[i];

      // Actualizar columna B con el enlace a la pestaña de Emergencia
      var celdaNombre = sheetCondensado.getRange(socio.row, 2);
      var finalUrl = socio.tarjetaUrl + (socio.gidEmergencia ? '#gid=' + socio.gidEmergencia : '');
      var linkFormula = '=HYPERLINK("' + finalUrl + '", "' + socio.nombre + '")';
      if (celdaNombre.getFormula() !== linkFormula) {
        celdaNombre.setFormula(linkFormula);
      }

      // Construir fórmula de aportaciones
      var fechaLunesFormatted = Utilities.formatDate(semana.lunesFecha, 'UTC', 'yyyy-MM-dd');
      var fechaDomingoFormatted = Utilities.formatDate(semana.domingoFecha, 'UTC', 'yyyy-MM-dd');
      
      var formulaBase = 'SUM(QUERY(IMPORTRANGE("' + socio.tarjetaUrl + '","\'Fondo de Emergencia\'!A4:D25"), "select Col3 where Col1 >= date \'' +
        fechaLunesFormatted + '\' and Col1 <= date \'' + fechaDomingoFormatted + '\'", 0))';

      var formulaSemana = '=IFERROR(' + formulaBase + ', 0)';

      var celdaDestino = sheetCondensado.getRange(socio.row, colDestino);
      if (semanaYaExiste) {
        if (!celdaDestino.getFormula() && celdaDestino.getValue() === '') {
          celdaDestino.setFormula('=IFERROR(' + formulaBase + ', 0)');
        }
      } else {
        celdaDestino.setFormula(formulaSemana);
      }
    }

    colActual++;
  }

  // Actualizar hoja resumen: fechas en columna A y suma de ingresos por semana en columna C.
  var sheetResumen = ss.getSheetByName('Resumen Fondo de Emergencia');
  if (sheetResumen) {
    var lastColResumen = sheetCondensado.getLastColumn();

    function numeroAColumna(colNum) {
      var resultado = '';
      while (colNum > 0) {
        var residuo = (colNum - 1) % 26;
        resultado = String.fromCharCode(65 + residuo) + resultado;
        colNum = Math.floor((colNum - 1) / 26);
      }
      return resultado;
    }

    // Mapa de fechas ya existentes en resumen (columna A, desde fila 3)
    var fechasExistentesResumen = {};
    var lastRowResumen = sheetResumen.getLastRow();
    if (lastRowResumen >= 3) {
      var fechasResumen = sheetResumen.getRange(3, 1, lastRowResumen - 2, 1).getValues();
      for (var fr = 0; fr < fechasResumen.length; fr++) {
        var valorFecha = fechasResumen[fr][0];
        var claveFecha = null;
        if (valorFecha instanceof Date) {
          claveFecha = Utilities.formatDate(valorFecha, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else if (typeof valorFecha === 'string' && valorFecha.trim() !== '') {
          claveFecha = convertirTextoAClaveFecha(valorFecha);
        }
        if (claveFecha) {
          fechasExistentesResumen[claveFecha] = 3 + fr;
        }
      }
    }

    var filaSiguiente = Math.max(3, sheetResumen.getLastRow() + 1);

    if (lastColResumen >= colStart) {
      var fechasHorizontales = sheetCondensado.getRange(3, colStart, 1, lastColResumen - colStart + 1).getValues()[0];

      for (var c = 0; c < fechasHorizontales.length; c++) {
        var fechaSemana = fechasHorizontales[c];
        if (fechaSemana === '' || fechaSemana === null) continue;

        var claveSemana = null;
        if (fechaSemana instanceof Date) {
          claveSemana = Utilities.formatDate(fechaSemana, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else if (typeof fechaSemana === 'string' && fechaSemana.trim() !== '') {
          claveSemana = convertirTextoAClaveFecha(fechaSemana);
        }
        if (!claveSemana) continue;

        var colActualResumen = colStart + c;
        var colLetra = numeroAColumna(colActualResumen);

        var filaDestinoResumen = fechasExistentesResumen[claveSemana];
        if (!filaDestinoResumen) {
          filaDestinoResumen = filaSiguiente;
          filaSiguiente++;
          fechasExistentesResumen[claveSemana] = filaDestinoResumen;
        }

        sheetResumen.getRange(filaDestinoResumen, 1).setFormula("='Fondo de Emergencia'!" + colLetra + "3");
        sheetResumen.getRange(filaDestinoResumen, 3).setFormula("=SUM('Fondo de Emergencia'!" + colLetra + "4:" + colLetra + ")");
        sheetResumen.getRange(filaDestinoResumen, 5).setFormula("=C" + filaDestinoResumen + "-D" + filaDestinoResumen);
      }
    }
  } else {
    Logger.log('No se encontró la hoja Resumen Fondo de Emergencia.');
  }

  Logger.log('Condensado de fondo de emergencia por semanas completado. Columnas procesadas: ' + (colActual - colStart));
}