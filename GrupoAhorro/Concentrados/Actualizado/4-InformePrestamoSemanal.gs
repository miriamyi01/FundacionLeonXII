function llenarCondensadoPrestamos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPrestamos = ss.getSheetByName('Préstamos');
  if (!sheetPrestamos) {
    Logger.log('No se encontró la hoja Préstamos.');
    return;
  }

  function normalizarCodigo(codigo) {
    return (codigo === null || codigo === undefined) ? '' : codigo.toString().trim();
  }

  function unirCeldasNombre(sheet) {
    var ultimaFila = sheet.getLastRow();
    if (ultimaFila < 4) return;
    var filasDatos = ultimaFila - 3;
    sheet.getRange(4, 2, filasDatos, 3).breakApart(); // Columnas B:D
    var codigos = sheet.getRange(4, 1, filasDatos, 1).getValues();
    for (var i = 0; i < codigos.length; i++) {
      if (normalizarCodigo(codigos[i][0])) sheet.getRange(4 + i, 2, 1, 3).merge();
    }
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

  function construirFormulaAbonoSemanal(fileId, sheetName, lunesFecha) {
    var domingoFecha = new Date(lunesFecha.getTime());
    domingoFecha.setUTCDate(domingoFecha.getUTCDate() + 6);

    var fechaLunesFormatted = Utilities.formatDate(lunesFecha, 'UTC', 'yyyy-MM-dd');
    var fechaDomingoFormatted = Utilities.formatDate(domingoFecha, 'UTC', 'yyyy-MM-dd');
    var sheetNameEscapado = sheetName.replace(/'/g, "''");
    var rango = "'" + sheetNameEscapado + "'!B13:D23";

    var consulta = "select Col3 where Col1 >= date '" + fechaLunesFormatted +
      "' and Col1 <= date '" + fechaDomingoFormatted + "'";

    return '=IFERROR(SUM(QUERY(IMPORTRANGE("' + fileId + '","' + rango + '"),"' + consulta + '",0)),0)';
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

  function extraerNombreHojaDesdeFormulaImportRange(formula) {
    if (!formula) return null;

    var m = formula.match(/IMPORTRANGE\("[^"]+","([^"]+)!/i);
    if (!m || !m[1]) return null;

    var nombreHoja = m[1];
    if (nombreHoja.charAt(0) === "'" && nombreHoja.charAt(nombreHoja.length - 1) === "'") {
      nombreHoja = nombreHoja.substring(1, nombreHoja.length - 1);
    }

    return nombreHoja.replace(/''/g, "'");
  }

  function numeroAColumna(colNum) {
    var resultado = '';
    while (colNum > 0) {
      var residuo = (colNum - 1) % 26;
      resultado = String.fromCharCode(65 + residuo) + resultado;
      colNum = Math.floor((colNum - 1) / 26);
    }
    return resultado;
  }

  var rootFolderName = 'GA0452 METAMORFOSIS';
  var sociosFolderName = 'GA0452-SOCIOS AS';
  var fichaFolderName = 'GA0452-FICHA DE INSCRIPCIÓN';

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

  // Buscar la carpeta de Fichas de Inscripción
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

  // Buscar el archivo de Inscripción dentro de la carpeta de fichas
  var sheetInscripcion = null;
  var archivosInscripcion = carpetaFicha.getFiles();
  while (archivosInscripcion.hasNext()) {
    var archivoInsc = archivosInscripcion.next();
    if (archivoInsc.getMimeType() === MimeType.GOOGLE_SHEETS) {
      try {
        var ssInsc = SpreadsheetApp.openById(archivoInsc.getId());
        var hoja = ssInsc.getSheetByName('Inscripción');
        if (hoja) {
          sheetInscripcion = hoja;
          break;
        }
      } catch (e) {
        Logger.log('Error abriendo archivo de inscripción: ' + e);
      }
    }
  }
  if (!sheetInscripcion) {
    Logger.log("No se encontró la hoja 'Inscripción' en la carpeta " + fichaFolderName);
    return;
  }

  // Buscar la carpeta de socios
  subFolders = rootFolder.getFolders();
  var sociosMainFolder = null;
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

  var headers = [
    'CÓDIGO',
    'NOMBRES',
    'PRIMER APELLIDO',
    'SEGUNDO APELLIDO',
    'FECHA OTORGADO',
    'FECHA TÉRMINO',
    'CANTIDAD',
    'ABONO A PRÉSTAMO',
    'INTERÉS',
    'ABONO A INTERÉS',
    'SALDO PENDIENTE'
  ];

  var colSemanasInicio = 12; // L
  
  // Leer datos existentes en la hoja
  var datosExistentes = {};
  var semanasExistentes = {};
  var semanasExistentesOriginal = {};
  var filasExistentes = {};
  var lastRowActual = sheetPrestamos.getLastRow();
  
  if (lastRowActual >= 4) {
    var rangoExistente = sheetPrestamos.getRange(4, 1, lastRowActual - 3, 11).getValues();
    var formulasColEExistente = sheetPrestamos.getRange(4, 5, lastRowActual - 3, 1).getFormulas();
    for (var ex = 0; ex < rangoExistente.length; ex++) {
      var codigo = rangoExistente[ex][0];
      if (codigo) {
        // Usar clave única: codigo_nombreHoja (leído desde fórmula de columna E)
        var rowIdx = 4 + ex;
        var formulaColE = formulasColEExistente[ex][0] || '';
        var nombreHojaExistente = extraerNombreHojaDesdeFormulaImportRange(formulaColE);
        var keyExist = nombreHojaExistente ? (codigo + '_' + nombreHojaExistente) : codigo;
        datosExistentes[keyExist] = {
          fila: rowIdx,
          datos: rangoExistente[ex]
        };
        filasExistentes[keyExist] = rowIdx;
      }
    }
  }
  
  // Leer semanas existentes desde fila 3, columnas L+
  if (lastRowActual > 0 && sheetPrestamos.getLastColumn() >= colSemanasInicio) {
    var semanasExistentesData = sheetPrestamos.getRange(3, colSemanasInicio, 1, sheetPrestamos.getLastColumn() - colSemanasInicio + 1).getValues()[0];
    for (var s = 0; s < semanasExistentesData.length; s++) {
      var fecha = semanasExistentesData[s];
      if (fecha instanceof Date || (typeof fecha === 'number' && fecha > 0)) {
        var claveSemana = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        semanasExistentes[claveSemana] = {
          fecha: fecha,
          colIndex: s
        };
        semanasExistentesOriginal[claveSemana] = true;
      } else if (typeof fecha === 'string' && fecha.trim() !== '') {
        var claveSemanaTexto = convertirTextoAClaveFecha(fecha);
        if (claveSemanaTexto) {
          var partesClave = claveSemanaTexto.split('-');
          var fechaDesdeTexto = new Date(Date.UTC(
            parseInt(partesClave[0], 10),
            parseInt(partesClave[1], 10) - 1,
            parseInt(partesClave[2], 10)
          ));
          semanasExistentes[claveSemanaTexto] = {
            fecha: fechaDesdeTexto,
            colIndex: s
          };
          semanasExistentesOriginal[claveSemanaTexto] = true;
        }
      }
    }
  }

  var lastRowIns = sheetInscripcion.getLastRow();
  if (lastRowIns < 7) {
    Logger.log('No hay socios en la hoja Inscripción para procesar.');
    return;
  }

  // Optimización: Mapear carpetas de socios como en InformeAhorroSemanal
  var foldersIter = sociosMainFolder.getFolders();
  var folderMap = {};
  while (foldersIter.hasNext()) {
    var f = foldersIter.next();
    var id = f.getName().split(" ")[0].trim();
    folderMap[id] = f;
  }

  var datosIns = sheetInscripcion.getRange(7, 1, lastRowIns - 6, 4).getValues();
  var filasPrestamos = [];
  var semanasMap = {};

  for (var i = 0; i < datosIns.length; i++) {
    var socioId = normalizarCodigo(datosIns[i][0]);
    if (!socioId || !folderMap[socioId]) continue;

    var socioFolder = folderMap[socioId];
    var nombres = datosIns[i][1] || '';
    var primerApellido = datosIns[i][2] || '';
    var segundoApellido = datosIns[i][3] || '';
    var nombreCompleto = (nombres + ' ' + primerApellido + ' ' + segundoApellido).trim();

    var carpetaNombre = socioFolder.getName();
    var partes = carpetaNombre.split(' ');
    var iniciales = partes.length > 1 ? partes[1] : '';
    var tarjetaFiles = socioFolder.getFiles();
    var tarjetaFile = null;
    
    while (tarjetaFiles.hasNext()) {
      var tf = tarjetaFiles.next();
      if (tf.getName().indexOf(iniciales) !== -1 && tf.getName().indexOf('TARJETA AHORRO Y PRESTAMO') !== -1) {
        tarjetaFile = tf;
        break;
      }
    }
    
    try {
      if (!tarjetaFile) continue;
      var tarjetaId = tarjetaFile.getId();
      var tarjeta = SpreadsheetApp.openById(tarjetaId);
      var tarjetaUrl = 'https://docs.google.com/spreadsheets/d/' + tarjetaId;
      var hojas = tarjeta.getSheets();
      
      for (var h = 0; h < hojas.length; h++) {
        var hojaPrestamo = hojas[h];
        var nombreHoja = hojaPrestamo.getName();
        if (nombreHoja.indexOf('Tarjeta Prestamo #') !== 0) continue;

        var cantidad = hojaPrestamo.getRange('E12').getValue();
        if (!cantidad || cantidad === 0) continue;

        var gid = hojaPrestamo.getSheetId();

        // Recolectar fechas para el encabezado global (incluso si no hay pagos aún)
        var datosPagos = hojaPrestamo.getRange('B13:B23').getValues();
        for (var p = 0; p < datosPagos.length; p++) {
          var fechaPago = datosPagos[p][0];
          if (fechaPago instanceof Date) {
            var lunes = obtenerLunesDelaSemana(fechaPago);
            var claveSemana = Utilities.formatDate(lunes, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            if (!semanasMap[claveSemana]) semanasMap[claveSemana] = lunes;
          }
        }

        filasPrestamos.push({
          codigo: socioId,
          nombreCompleto: nombreCompleto,
          tarjetaFileId: tarjetaId,
          tarjetaUrl: tarjetaUrl,
          nombreHoja: nombreHoja,
          gid: gid
        });
      }
    } catch (e) {
      Logger.log('Error abriendo archivo para socio ' + socioId + ': ' + e);
    }
  }

  if (filasPrestamos.length > 0) {
    // Construir mapa de nuevos préstamos por código + nombreHoja (para múltiples préstamos por socio)
    var nuevosPrestamos = {};
    var indicePorCodigo = {};
    for (var np = 0; np < filasPrestamos.length; np++) {
      var prestamoId = filasPrestamos[np].codigo + '_' + filasPrestamos[np].nombreHoja;
      nuevosPrestamos[prestamoId] = np;
      indicePorCodigo[prestamoId] = np;
    }

    // Agregar nuevas semanas a semanasExistentes
    var clavesSemanas = Object.keys(semanasMap).sort();
    for (var cs = 0; cs < clavesSemanas.length; cs++) {
      var clave = clavesSemanas[cs];
      if (!semanasExistentes[clave]) {
        semanasExistentes[clave] = {
          fecha: semanasMap[clave],
          colIndex: null // Se asignará después
        };
      }
    }

    // Reconstruir el orden de semanas
    var clavesSemanasFinal = Object.keys(semanasExistentes).sort();
    var nuevaSemanasMap = {};
    for (var ncs = 0; ncs < clavesSemanasFinal.length; ncs++) {
      nuevaSemanasMap[clavesSemanasFinal[ncs]] = {
        fecha: semanasExistentes[clavesSemanasFinal[ncs]].fecha,
        colIndex: ncs
      };
    }
    semanasExistentes = nuevaSemanasMap;

    // Escribir encabezados de fechas en fila 3
    var fechasHeader = [];
    for (var fh = 0; fh < clavesSemanasFinal.length; fh++) {
      var fechaLunesHeader = semanasExistentes[clavesSemanasFinal[fh]].fecha;
      fechasHeader.push(Utilities.formatDate(fechaLunesHeader, 'UTC', 'dd/MM/yyyy'));
    }
    if (fechasHeader.length > 0) {
      sheetPrestamos.getRange(3, colSemanasInicio, 1, fechasHeader.length).setValues([fechasHeader]);
    }

    // Las fórmulas de suma en fila 2 se escribirán solo para semanas nuevas

    // Procesar cada préstamo
    var filasNuevas = [];
    var filasModificadas = [];
    
    // Identificar semanas nuevas vs existentes
    var semanasNuevas = [];
    var semanasNuevasIndices = [];
    for (var csn = 0; csn < clavesSemanasFinal.length; csn++) {
      var claveSemanaFinal = clavesSemanasFinal[csn];
      if (!semanasExistentesOriginal[claveSemanaFinal]) {
        semanasNuevas.push(claveSemanaFinal);
        semanasNuevasIndices.push(csn);
      }
    }

    // Escribir fórmulas de suma por columna en fila 2 solo para semanas nuevas
    for (var snf = 0; snf < semanasNuevasIndices.length; snf++) {
      var idxColNueva = semanasNuevasIndices[snf];
      var colSuma = colSemanasInicio + idxColNueva;
      var colLetra = numeroAColumna(colSuma);
      sheetPrestamos.getRange(2, colSuma).setFormula('=SUM(' + colLetra + '4:' + colLetra + ')');
    }
    
    for (var cod_key in nuevosPrestamos) {
      var idxNuevo = nuevosPrestamos[cod_key];
      
      // Buscar si existe este mismo préstamo
      var prestamoExistente = datosExistentes[cod_key];
      
      if (prestamoExistente) {
        // Préstamo existente - actualizar solo nuevas semanas
        var filaExistente = prestamoExistente.fila;
        var formulasSemanasActualizar = [];
        var fileIdExist = filasPrestamos[idxNuevo].tarjetaFileId;
        var sheetNameExist = filasPrestamos[idxNuevo].nombreHoja;
        
        // Actualizar enlace de nombre para socios existentes (estilo InformeInicial)
        var finalUrlExist = filasPrestamos[idxNuevo].tarjetaUrl + (filasPrestamos[idxNuevo].gid ? '#gid=' + filasPrestamos[idxNuevo].gid : '');
        var formulaNombreExist = filasPrestamos[idxNuevo].tarjetaUrl 
          ? '=HYPERLINK("' + finalUrlExist + '", "' + filasPrestamos[idxNuevo].nombreCompleto + '")' 
          : filasPrestamos[idxNuevo].nombreCompleto;
          
        if (sheetPrestamos.getRange(filaExistente, 2).getFormula() !== formulaNombreExist) {
          sheetPrestamos.getRange(filaExistente, 2).setFormula(formulaNombreExist);
        }

        for (var smn2 = 0; smn2 < semanasNuevasIndices.length; smn2++) {
          var idxSemana = semanasNuevasIndices[smn2];
          var claveSem = clavesSemanasFinal[idxSemana];
          var lunesSemana = semanasExistentes[claveSem].fecha;
          formulasSemanasActualizar.push(construirFormulaAbonoSemanal(fileIdExist, sheetNameExist, lunesSemana));
        }

        if (semanasNuevasIndices.length > 0) {
          filasModificadas.push({
            fila: filaExistente,
            formulas: formulasSemanasActualizar,
            colInicio: colSemanasInicio + semanasNuevasIndices[0],
            cantCols: semanasNuevasIndices.length
          });
        }
      } else {
        // Préstamo nuevo - agregar fila completa
        filasNuevas.push({
          codigo: filasPrestamos[idxNuevo].codigo,
          nombreCompleto: filasPrestamos[idxNuevo].nombreCompleto,
          tarjetaFileId: filasPrestamos[idxNuevo].tarjetaFileId,
          nombreHoja: filasPrestamos[idxNuevo].nombreHoja,
          tarjetaUrl: filasPrestamos[idxNuevo].tarjetaUrl,
          gid: filasPrestamos[idxNuevo].gid
        });
      }
    }

    // Escribir datos principales de nuevos préstamos
    if (filasNuevas.length > 0) {
      var primerFilaNueva = lastRowActual + 1;
      for (var fn = 0; fn < filasNuevas.length; fn++) {
        var fila = filasNuevas[fn];
        var filaDestino = primerFilaNueva + fn;
        
        // Escribir Código y Nombre con Hyperlink (estilo InformeInicial)
        var finalUrl = fila.tarjetaUrl + (fila.gid ? '#gid=' + fila.gid : '');
        sheetPrestamos.getRange(filaDestino, 1).setValue(fila.codigo);
        var formulaHyperlink = fila.tarjetaUrl ? '=HYPERLINK("' + finalUrl + '", "' + fila.nombreCompleto + '")' : fila.nombreCompleto;
        sheetPrestamos.getRange(filaDestino, 2).setFormula(formulaHyperlink);

        var formulas = [];
        var fileId = fila.tarjetaFileId;
        var sheetName = fila.nombreHoja;
        
        formulas.push('=IFERROR(IMPORTRANGE("' + fileId + '","' + sheetName + '!B12"),"")'); // E - Fecha Otorgado
        formulas.push('=IFERROR(IMPORTRANGE("' + fileId + '","' + sheetName + '!C5"),"")');  // F - Fecha Término
        formulas.push('=IFERROR(IMPORTRANGE("' + fileId + '","' + sheetName + '!E12"),"")'); // G - Cantidad
        formulas.push('=IFERROR(IMPORTRANGE("' + fileId + '","' + sheetName + '!D24"),"")'); // H - Abono Préstamo
        formulas.push('=IFERROR(IMPORTRANGE("' + fileId + '","' + sheetName + '!F5"),"")');  // I - Interés
        formulas.push('=IFERROR(IMPORTRANGE("' + fileId + '","' + sheetName + '!F24"),"")'); // J - Abono Interés
        formulas.push('=IFERROR(IMPORTRANGE("' + fileId + '","' + sheetName + '!H24"),"")'); // K - Saldo Pendiente
        
        sheetPrestamos.getRange(filaDestino, 5, 1, 7).setFormulas([formulas]);
        
        // Escribir fórmulas semanales (validando que existan columnas)
        if (clavesSemanasFinal.length > 0) {
          var formulasSemanales = [];
          for (var smnForm = 0; smnForm < clavesSemanasFinal.length; smnForm++) {
            var claveSemanaFormula = clavesSemanasFinal[smnForm];
            var lunesSemanaFormula = semanasExistentes[claveSemanaFormula].fecha;
            formulasSemanales.push(construirFormulaAbonoSemanal(fileId, sheetName, lunesSemanaFormula));
          }
          sheetPrestamos.getRange(filaDestino, colSemanasInicio, 1, clavesSemanasFinal.length).setFormulas([formulasSemanales]);
        }
      }
    }

    // Ordenar alfabéticamente por nombre y unir celdas (B:D)
    var ultimaFilaFinal = sheetPrestamos.getLastRow();
    if (ultimaFilaFinal >= 4) {
      sheetPrestamos.getRange(4, 1, ultimaFilaFinal - 3, sheetPrestamos.getLastColumn()).sort({ column: 2, ascending: true });
      unirCeldasNombre(sheetPrestamos);
    }

    // Escribir fórmulas de semanas para préstamos modificados (solo nuevas columnas)
    for (var fm = 0; fm < filasModificadas.length; fm++) {
      if (filasModificadas[fm].cantCols > 0) {
        sheetPrestamos.getRange(filasModificadas[fm].fila, filasModificadas[fm].colInicio, 1, filasModificadas[fm].cantCols).setFormulas([filasModificadas[fm].formulas]);
      }
    }

    Logger.log('Detalle de préstamos completado. Nuevos: ' + filasNuevas.length + ', Actualizados: ' + filasModificadas.length);
  } else {
    Logger.log('Detalle de préstamos completado. Filas cargadas: 0');
  }
}