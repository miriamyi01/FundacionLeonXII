function llenarCondensadoPrestamos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCondensado = ss.getSheetByName('Ahorros y Retiros');
  if (!sheetCondensado) {
    Logger.log('No se encontró la hoja Ahorros y Retiros.');
    return;
  }

  // Obtener IDs de socios desde la columna A, desde la fila 8, omitiendo las últimas 3 filas
  var lastRow = sheetCondensado.getLastRow();
  var sociosRows = lastRow - 7 - 3; // omitir las últimas 3 filas
  if (sociosRows <= 0) {
    Logger.log('No hay filas de socios para procesar.');
    return;
  }
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

  // Buscar hoja Prestamos en el condensado
  var sheetPrestamos = ss.getSheetByName('Prestamos');
  if (!sheetPrestamos) {
    Logger.log('No se encontró la hoja Prestamos.');
    return;
  }

  // Obtener préstamos existentes para evitar duplicados
  var prestamosExistentes = new Set();
  var lastRowPrestamos = sheetPrestamos.getLastRow();
  var lastColPrestamos = sheetPrestamos.getLastColumn();
  var ultimaColumnaConPrestamo = 2; // Empezar antes de la columna 3 (C)
  
  if (lastRowPrestamos >= 5 && lastColPrestamos >= 3) {
    // Leer los préstamos existentes desde la fila 5 (número de préstamo) y fila 6 (código de socio)
    for (var col = 3; col <= lastColPrestamos; col += 2) { // Avanzar de 2 en 2 columnas
      var numeroPrestamo = sheetPrestamos.getRange(5, col).getValue();
      var codigoSocio = sheetPrestamos.getRange(6, col).getValue();
      if (numeroPrestamo && codigoSocio) {
        var clavePrestamo = codigoSocio + '#' + numeroPrestamo;
        prestamosExistentes.add(clavePrestamo);
        ultimaColumnaConPrestamo = col + 1; // Actualizar la última columna que tiene préstamo real
      }
    }
  }
  
  Logger.log('Préstamos existentes encontrados: ' + prestamosExistentes.size);

  // Solo limpiar contenido de columnas que no tienen préstamos existentes
  // NO TOCAR FILA 15 - mantener contenido existente
  // Encontrar la primera columna disponible para nuevos préstamos
  var columnaPrestamo = ultimaColumnaConPrestamo + 1; // Empezar después del último préstamo real
  if (columnaPrestamo < 3) columnaPrestamo = 3; // Mínimo en columna C
  if (columnaPrestamo % 2 === 0) columnaPrestamo++; // Asegurar que empiece en columna impar (C, E, G, etc.)

  var contadorPrestamos = 0; // Contador de préstamos procesados
  var prestamosNuevos = 0; // Contador de préstamos nuevos agregados

  for (var i = 0; i < sociosData.length; i++) {
    var socioId = sociosData[i][0];
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

    // Buscar archivo del socio
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
      Logger.log('No se encontró archivo TARJETA AHORRO Y PRESTAMO para socio: ' + socioId + ' (' + iniciales + ')');
      continue;
    }

    var tarjetaId = tarjetaFile.getId();
    
    // Abrir la tarjeta de ahorro con manejo de errores
    var tarjeta;
    try {
      tarjeta = SpreadsheetApp.openById(tarjetaId);
    } catch (e) {
      Logger.log('Error abriendo archivo para socio ' + socioId + ': ' + e);
      continue;
    }
    
    // Buscar hojas de préstamos (Tarjeta Prestamo #1, #2, etc.)
    var hojas = tarjeta.getSheets();
    for (var h = 0; h < hojas.length; h++) {
      var hoja = hojas[h];
      var nombreHoja = hoja.getName();
      
      if (nombreHoja.indexOf('Tarjeta Prestamo #') !== -1) {
        // Verificar si hay cantidad en E12
        var cantidad = hoja.getRange('E12').getValue();
        if (!cantidad || cantidad === 0) continue;

        // Extraer número de préstamo del nombre de la hoja
        var numeroPrestamo = nombreHoja.match(/#(\d+)/);
        numeroPrestamo = numeroPrestamo ? numeroPrestamo[1] : '';
        
        // Obtener código del socio para verificar si ya existe
        var codigoSocio = '';
        try {
          var tarjetaAhorro = tarjeta.getSheetByName('Tarjeta Ahorro');
          if (tarjetaAhorro) {
            codigoSocio = tarjetaAhorro.getRange('D1').getValue();
          }
        } catch (e) {
          Logger.log('Error obteniendo código de socio para ' + socioId + ': ' + e);
          codigoSocio = socioId; // Usar el ID del socio como fallback
        }
        
        // Verificar si este préstamo ya existe
        var clavePrestamo = codigoSocio + '#' + numeroPrestamo;
        if (prestamosExistentes.has(clavePrestamo)) {
          Logger.log('Préstamo ya existe, omitiendo: ' + clavePrestamo);
          continue;
        }

        // Llenar datos del préstamo usando dos columnas (datos en C, más datos en D)
        // La mayoría de datos salen de la hoja "Tarjeta Prestamo #X" del socio
        sheetPrestamos.getRange(5, columnaPrestamo).setValue(numeroPrestamo); // # (valor directo)
        sheetPrestamos.getRange(6, columnaPrestamo).setFormula('=IFERROR(IMPORTRANGE("' + tarjetaId + '","Tarjeta Ahorro!D1"),"")'); // Cod. (viene de Tarjeta Ahorro)
        sheetPrestamos.getRange(7, columnaPrestamo).setValue(iniciales); // Nom. (iniciales del nombre de la carpeta)
        sheetPrestamos.getRange(8, columnaPrestamo).setFormula('=IFERROR(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!B12"),"")'); // Fecha otorgado (de Tarjeta Prestamo)
        sheetPrestamos.getRange(9, columnaPrestamo).setFormula('=IFERROR(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!E12"),"")'); // Cantidad (de Tarjeta Prestamo)
        sheetPrestamos.getRange(10, columnaPrestamo).setFormula('=IFERROR(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!H24"),"")'); // Paga (de Tarjeta Prestamo)
        
        // Pen - última celda de columna H con valor (saldo pendiente) desde H13
        var formulaPen = '=IFERROR(INDEX(FILTER(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!H13:H23"),IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!H13:H23")<>"",IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!H13:H23")<>0),COUNTA(FILTER(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!H13:H23"),IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!H13:H23")<>"",IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!H13:H23")<>0))),0)';
        sheetPrestamos.getRange(11, columnaPrestamo).setFormula(formulaPen); // Pen (de Tarjeta Prestamo)
        
        sheetPrestamos.getRange(12, columnaPrestamo).setFormula('=IFERROR(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!C6"),"")'); // Destino (de Tarjeta Prestamo)
        sheetPrestamos.getRange(13, columnaPrestamo).setFormula('=IFERROR(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!F5"),"")'); // Int. (de Tarjeta Prestamo)
        sheetPrestamos.getRange(14, columnaPrestamo).setFormula('=IFERROR(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!C5"),"")'); // Pago (de Tarjeta Prestamo)

        // Llenar pagos mensuales según estructura específica
        // Los pagos empiezan después de los datos básicos del préstamo
        // Los meses están en A16:A17 con "sem" y "mon" en columna B
        // NO TOCAR LA FILA 15 - mantener contenido existente
        var meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
        
        // Empezar los pagos mensuales en la fila 16 (donde están los meses)
        // Cada mes ocupa 2 filas: 16-17 para ENERO, 18-19 para FEBRERO, etc.
        var filaInicioMeses = 16;
        for (var m = 0; m < meses.length; m++) {
          // Usar las dos columnas del préstamo actual para cada mes
          var colPrimera = columnaPrestamo; // Primera columna del préstamo
          var colSegunda = columnaPrestamo + 1; // Segunda columna del préstamo
          var mesTexto = meses[m];
          var anio = new Date().getFullYear();
          var numeroMes = m + 1;
          
          // Calcular la fila específica para este mes (saltando fila 15)
          var filaDelMes = filaInicioMeses + (m * 2); // 16, 18, 20, 22, etc.
          
          // Fila "sem": Primera columna vacía, Segunda columna = número de semana del mes del último abono del mes
          // Buscar la fecha más reciente del mes que tenga abono y calcular su semana
          var formulaSemanaDelMes = '=IFERROR(IF(SUMPRODUCT((MONTH(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!B13:B23"))=' + numeroMes + ')*(YEAR(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!B13:B23"))=' + anio + ')*(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!D13:D23")>0))>0,INT((DAY(INDEX(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!B13:B23"),MATCH(MAX(IF((MONTH(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!B13:B23"))=' + numeroMes + ')*(YEAR(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!B13:B23"))=' + anio + ')*(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!D13:D23")>0),IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!B13:B23"))),IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!B13:B23"),0)))-1)/7)+1,0),0)';
          // Primera columna de "sem" se deja vacía (no se escribe nada)
          sheetPrestamos.getRange(filaDelMes, colSegunda).setFormula(formulaSemanaDelMes); // Número de semana del mes (fila "sem", segunda columna)
          
          // Fila "mon": Primera columna = intereses del mes, Segunda columna = abonos del mes
          var formulaInt = '=IFERROR(SUMPRODUCT((MONTH(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!B13:B23"))=' + numeroMes + ')*(YEAR(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!B13:B23"))=' + anio + ')*(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!F13:F23"))),0)';
          sheetPrestamos.getRange(filaDelMes + 1, colPrimera).setFormula(formulaInt); // Intereses del mes (fila "mon", primera columna)
          
          var formulaCap = '=IFERROR(SUMPRODUCT((MONTH(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!B13:B23"))=' + numeroMes + ')*(YEAR(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!B13:B23"))=' + anio + ')*(IMPORTRANGE("' + tarjetaId + '","' + nombreHoja + '!D13:D23"))),0)';
          sheetPrestamos.getRange(filaDelMes + 1, colSegunda).setFormula(formulaCap); // Abonos del mes (fila "mon", segunda columna)
        }

        columnaPrestamo += 2; // Avanzar 2 columnas para el próximo préstamo (datos + datos adicionales)
        contadorPrestamos++;
        prestamosNuevos++;
        Logger.log('Préstamo NUEVO #' + numeroPrestamo + ' procesado para socio: ' + socioId + ' (' + iniciales + ') en columnas ' + (columnaPrestamo - 2) + '-' + (columnaPrestamo - 1));
      }
    }
  }
  Logger.log('Condensado de préstamos completado. Total de préstamos procesados: ' + contadorPrestamos + ', Préstamos nuevos agregados: ' + prestamosNuevos);
}