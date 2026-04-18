// Ajusta la altura de filas y actualiza F1 en la hoja 'Tarjeta Ahorro' de cada socio
function actualizarTarjetaAhorroSocios() {
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


  // Buscar el archivo base '04 TARJETA AHORRO Y PRESTAMO' en la carpeta del script
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
  var ssBase = SpreadsheetApp.openById(baseFileId);
  var hojaAhorroBase = ssBase.getSheetByName('Tarjeta Ahorro');
  if (!hojaAhorroBase) {
    Logger.log('No se encontró la hoja "Tarjeta Ahorro" en el archivo base.');
    return;
  }

  // Usar directamente la hoja 'Inscripción' del archivo activo
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inscripción');
  if (!sheet) {
    Logger.log("No existe la hoja 'Inscripción' en el archivo activo.");
    return;
  }

  var mainFolder = sociosMainFolder;
  var lastRow = sheet.getLastRow();
  var sociosProcesados = 0;
  var archivosActualizados = 0;

  for (var i = 7; i <= lastRow; i++) {
    var numeroSocio = sheet.getRange("A" + i).getValue().toString().trim();
    var nombreCompleto = [
      sheet.getRange("B" + i).getValue().toString().trim(),
      sheet.getRange("C" + i).getValue().toString().trim(),
      sheet.getRange("D" + i).getValue().toString().trim()
    ].filter(Boolean).join(' ').trim();
    var valorF = sheet.getRange("F" + i).getValue();
    if (!nombreCompleto) continue;

    var numeroU = numeroSocio.toUpperCase();

    // Buscar carpeta del socio
    var carpetaSocio = null;
    var subIt = mainFolder.getFolders();
    while (subIt.hasNext()) {
      var sf = subIt.next();
      if (sf.getName().toUpperCase().indexOf(numeroU + " ") === 0) {
        carpetaSocio = sf;
        break;
      }
    }
    if (!carpetaSocio) {
      sociosProcesados++;
      continue;
    }

    // Buscar archivo por nombre: [iniciales] - TARJETA AHORRO Y PRESTAMO
    var folderNameParts = carpetaSocio.getName().split(' ');
    var inicialesCarpeta = folderNameParts[1] || '';
    var archivoNombre = inicialesCarpeta + ' - TARJETA AHORRO Y PRESTAMO';
    var archivosExistentes = carpetaSocio.getFilesByName(archivoNombre);
    var destFile = null;
    if (archivosExistentes.hasNext()) {
      destFile = archivosExistentes.next();
    }
    if (!destFile) {
      sociosProcesados++;
      continue;
    }

    // Abrir el archivo del socio
    var destSS = SpreadsheetApp.openById(destFile.getId());
    var hojaAhorro = destSS.getSheetByName('Tarjeta Ahorro');
    if (!hojaAhorro) {
      sociosProcesados++;
      continue;
    }


    // Ajustar altura de filas 2, 4 y 7 a 10px
    hojaAhorro.setRowHeight(2, 10);
    hojaAhorro.setRowHeight(4, 10);
    hojaAhorro.setRowHeight(7, 10);

    // Copiar solo color y estilo de B6 y E6 de la hoja base y aplicarlo a la hoja destino
    var formatoB6 = hojaAhorroBase.getRange('B6');
    var formatoE6 = hojaAhorroBase.getRange('E6');
    hojaAhorro.getRange('B6')
      .setFontColor(formatoB6.getFontColor())
      .setFontLine(formatoB6.getFontLine())
      .setFontStyle(formatoB6.getFontStyle())
      .setFontWeight(formatoB6.getFontWeight());
    hojaAhorro.getRange('E6')
      .setFontColor(formatoE6.getFontColor())
      .setFontLine(formatoE6.getFontLine())
      .setFontStyle(formatoE6.getFontStyle())
      .setFontWeight(formatoE6.getFontWeight());

    // Sustituir F1 con el valor correspondiente de la columna F de la hoja Inscripción
    hojaAhorro.getRange('F1').setValue(valorF);

    // Quitar todos los formatos condicionales existentes
    hojaAhorro.setConditionalFormatRules([]);

    // Copiar y ajustar los formatos condicionales de la hoja 'Tarjeta Ahorro' del archivo base
    var reglasBase = hojaAhorroBase.getConditionalFormatRules();
    if (reglasBase && reglasBase.length > 0) {
      var reglasCopiadas = reglasBase.map(function(regla) {
        var builder = regla.copy();
        // Ajustar los rangos para que sean sobre la hoja destino
        var nuevosRangos = regla.getRanges().map(function(rangoBase) {
          return hojaAhorro.getRange(rangoBase.getA1Notation());
        });
        builder.setRanges(nuevosRangos);
        return builder.build();
      });
      hojaAhorro.setConditionalFormatRules(reglasCopiadas);
    }

    archivosActualizados++;
    Logger.log("Actualizado: " + carpetaSocio.getName());
    sociosProcesados++;
  }

  Logger.log("¡Actualización de Tarjeta Ahorro completada!");
  Logger.log("Socios procesados: " + sociosProcesados);
  Logger.log("Archivos actualizados: " + archivosActualizados);
}



// Coloca el nombre del socio en C47 y el código en G47 en todas las hojas 'Tarjeta Prestamo #N' de cada socio
function actualizarDatosPrestamosSocios() {
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

  // Usar directamente la hoja 'Inscripción' del archivo activo
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inscripción');
  if (!sheet) {
    Logger.log("No existe la hoja 'Inscripción' en el archivo activo.");
    return;
  }

  var mainFolder = sociosMainFolder;
  var lastRow = sheet.getLastRow();
  var sociosProcesados = 0;
  var archivosActualizados = 0;

  for (var i = 7; i <= lastRow; i++) {
    var numeroSocio = sheet.getRange("A" + i).getValue().toString().trim();
    var nombreCompleto = [
      sheet.getRange("B" + i).getValue().toString().trim(),
      sheet.getRange("C" + i).getValue().toString().trim(),
      sheet.getRange("D" + i).getValue().toString().trim()
    ].filter(Boolean).join(' ').trim();
    if (!nombreCompleto) continue;

    var numeroU = numeroSocio.toUpperCase();

    // Buscar carpeta del socio
    var carpetaSocio = null;
    var subIt = mainFolder.getFolders();
    while (subIt.hasNext()) {
      var sf = subIt.next();
      if (sf.getName().toUpperCase().indexOf(numeroU + " ") === 0) {
        carpetaSocio = sf;
        break;
      }
    }
    if (!carpetaSocio) {
      sociosProcesados++;
      continue;
    }

    // Buscar archivo por nombre: [iniciales] - TARJETA AHORRO Y PRESTAMO
    var folderNameParts = carpetaSocio.getName().split(' ');
    var inicialesCarpeta = folderNameParts[1] || '';
    var archivoNombre = inicialesCarpeta + ' - TARJETA AHORRO Y PRESTAMO';
    var archivosExistentes = carpetaSocio.getFilesByName(archivoNombre);
    var destFile = null;
    if (archivosExistentes.hasNext()) {
      destFile = archivosExistentes.next();
    }
    if (!destFile) {
      sociosProcesados++;
      continue;
    }

    // Abrir el archivo del socio
    var destSS = SpreadsheetApp.openById(destFile.getId());
    var hojas = destSS.getSheets();
    for (var h = 0; h < hojas.length; h++) {
      var hoja = hojas[h];
      var nombreHoja = hoja.getName();
      if (nombreHoja.indexOf('Tarjeta Prestamo #') === 0) {
        hoja.getRange('C47').setValue(nombreCompleto);
        hoja.getRange('G47').setValue(numeroU);
      }
    }
    archivosActualizados++;
    sociosProcesados++;
  }

  Logger.log("¡Actualización de Tarjeta Prestamo #N completada!");
  Logger.log("Socios procesados: " + sociosProcesados);
  Logger.log("Archivos actualizados: " + archivosActualizados);
}



// Coloca la fórmula =C3 en E12 en todas las hojas 'Tarjeta Prestamo #N' de cada socio
function actualizarFormulaE12PrestamosSocios() {
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

  // Usar directamente la hoja 'Inscripción' del archivo activo
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inscripción');
  if (!sheet) {
    Logger.log("No existe la hoja 'Inscripción' en el archivo activo.");
    return;
  }

  var mainFolder = sociosMainFolder;
  var lastRow = sheet.getLastRow();
  var sociosProcesados = 0;
  var archivosActualizados = 0;

  for (var i = 7; i <= lastRow; i++) {
    var numeroSocio = sheet.getRange('A' + i).getValue().toString().trim();
    var nombreCompleto = [
      sheet.getRange('B' + i).getValue().toString().trim(),
      sheet.getRange('C' + i).getValue().toString().trim(),
      sheet.getRange('D' + i).getValue().toString().trim()
    ].filter(Boolean).join(' ').trim();
    if (!nombreCompleto) continue;

    var numeroU = numeroSocio.toUpperCase();

    // Buscar carpeta del socio
    var carpetaSocio = null;
    var subIt = mainFolder.getFolders();
    while (subIt.hasNext()) {
      var carpeta = subIt.next();
      if (carpeta.getName().toUpperCase().indexOf(numeroU + ' ') === 0) {
        carpetaSocio = carpeta;
        break;
      }
    }
    if (!carpetaSocio) {
      sociosProcesados++;
      continue;
    }

    // Buscar archivo por nombre: [iniciales] - TARJETA AHORRO Y PRESTAMO
    var folderNameParts = carpetaSocio.getName().split(' ');
    var inicialesCarpeta = folderNameParts[1] || '';
    var archivoNombre = inicialesCarpeta + ' - TARJETA AHORRO Y PRESTAMO';
    var archivosExistentes = carpetaSocio.getFilesByName(archivoNombre);
    var destFile = null;
    if (archivosExistentes.hasNext()) {
      destFile = archivosExistentes.next();
    }
    if (!destFile) {
      sociosProcesados++;
      continue;
    }

    // Abrir el archivo del socio
    var destSS = SpreadsheetApp.openById(destFile.getId());
    var hojas = destSS.getSheets();
    for (var h = 0; h < hojas.length; h++) {
      var hoja = hojas[h];
      var nombreHoja = hoja.getName();
      if (nombreHoja.indexOf('Tarjeta Prestamo #') === 0) {
        hoja.getRange('E12').setFormula('=C3');
      }
    }

    archivosActualizados++;
    sociosProcesados++;
  }

  Logger.log('¡Actualización de fórmula E12 en Tarjeta Prestamo #N completada!');
  Logger.log('Socios procesados: ' + sociosProcesados);
  Logger.log('Archivos actualizados: ' + archivosActualizados);
}



// Coloca la fórmula =0.01/30 en F6 en todas las hojas 'Tarjeta Prestamo #N' de cada socio
function actualizarFormulaF6PrestamosSocios() {
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

  // Usar directamente la hoja 'Inscripción' del archivo activo
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inscripción');
  if (!sheet) {
    Logger.log("No existe la hoja 'Inscripción' en el archivo activo.");
    return;
  }

  var mainFolder = sociosMainFolder;
  var lastRow = sheet.getLastRow();
  var sociosProcesados = 0;
  var archivosActualizados = 0;

  for (var i = 7; i <= lastRow; i++) {
    var numeroSocio = sheet.getRange('A' + i).getValue().toString().trim();
    var nombreCompleto = [
      sheet.getRange('B' + i).getValue().toString().trim(),
      sheet.getRange('C' + i).getValue().toString().trim(),
      sheet.getRange('D' + i).getValue().toString().trim()
    ].filter(Boolean).join(' ').trim();
    if (!nombreCompleto) continue;

    var numeroU = numeroSocio.toUpperCase();

    // Buscar carpeta del socio
    var carpetaSocio = null;
    var subIt = mainFolder.getFolders();
    while (subIt.hasNext()) {
      var carpeta = subIt.next();
      if (carpeta.getName().toUpperCase().indexOf(numeroU + ' ') === 0) {
        carpetaSocio = carpeta;
        break;
      }
    }
    if (!carpetaSocio) {
      sociosProcesados++;
      continue;
    }

    // Buscar archivo por nombre: [iniciales] - TARJETA AHORRO Y PRESTAMO
    var folderNameParts = carpetaSocio.getName().split(' ');
    var inicialesCarpeta = folderNameParts[1] || '';
    var archivoNombre = inicialesCarpeta + ' - TARJETA AHORRO Y PRESTAMO';
    var archivosExistentes = carpetaSocio.getFilesByName(archivoNombre);
    var destFile = null;
    if (archivosExistentes.hasNext()) {
      destFile = archivosExistentes.next();
    }
    if (!destFile) {
      sociosProcesados++;
      continue;
    }

    // Abrir el archivo del socio
    var destSS = SpreadsheetApp.openById(destFile.getId());
    var hojas = destSS.getSheets();
    for (var h = 0; h < hojas.length; h++) {
      var hoja = hojas[h];
      var nombreHoja = hoja.getName();
      if (nombreHoja.indexOf('Tarjeta Prestamo #') === 0) {
        hoja.getRange('F6').setFormula('=0.01/30');
      }
    }

    archivosActualizados++;
    sociosProcesados++;
  }

  Logger.log('¡Actualización de fórmula F6 en Tarjeta Prestamo #N completada!');
  Logger.log('Socios procesados: ' + sociosProcesados);
  Logger.log('Archivos actualizados: ' + archivosActualizados);
}



// Copia los formatos condicionales de la hoja 'Tarjeta Ahorro' del archivo base a cada socio
function copiarFormatosCondicionalesTarjetaAhorroSocios() {
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

  // Buscar el archivo base '04 TARJETA AHORRO Y PRESTAMO' en la carpeta del script
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

  var ssBase = SpreadsheetApp.openById(baseFileId);
  var hojaAhorroBase = ssBase.getSheetByName('Tarjeta Ahorro');
  if (!hojaAhorroBase) {
    Logger.log('No se encontró la hoja "Tarjeta Ahorro" en el archivo base.');
    return;
  }

  var reglasBase = hojaAhorroBase.getConditionalFormatRules();

  // Usar directamente la hoja 'Inscripción' del archivo activo para recorrer socios
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inscripción');
  if (!sheet) {
    Logger.log("No existe la hoja 'Inscripción' en el archivo activo.");
    return;
  }

  var mainFolder = sociosMainFolder;
  var lastRow = sheet.getLastRow();
  var sociosProcesados = 0;
  var archivosActualizados = 0;

  for (var i = 7; i <= lastRow; i++) {
    var numeroSocio = sheet.getRange('A' + i).getValue().toString().trim();
    var nombreCompleto = [
      sheet.getRange('B' + i).getValue().toString().trim(),
      sheet.getRange('C' + i).getValue().toString().trim(),
      sheet.getRange('D' + i).getValue().toString().trim()
    ].filter(Boolean).join(' ').trim();
    if (!nombreCompleto) continue;

    var numeroU = numeroSocio.toUpperCase();

    // Buscar carpeta del socio
    var carpetaSocio = null;
    var subIt = mainFolder.getFolders();
    while (subIt.hasNext()) {
      var carpeta = subIt.next();
      if (carpeta.getName().toUpperCase().indexOf(numeroU + ' ') === 0) {
        carpetaSocio = carpeta;
        break;
      }
    }
    if (!carpetaSocio) {
      sociosProcesados++;
      continue;
    }

    // Buscar archivo por nombre: [iniciales] - TARJETA AHORRO Y PRESTAMO
    var folderNameParts = carpetaSocio.getName().split(' ');
    var inicialesCarpeta = folderNameParts[1] || '';
    var archivoNombre = inicialesCarpeta + ' - TARJETA AHORRO Y PRESTAMO';
    var archivosExistentes = carpetaSocio.getFilesByName(archivoNombre);
    var destFile = null;
    if (archivosExistentes.hasNext()) {
      destFile = archivosExistentes.next();
    }
    if (!destFile) {
      sociosProcesados++;
      continue;
    }

    // Abrir archivo del socio y copiar reglas a la hoja Tarjeta Ahorro
    var destSS = SpreadsheetApp.openById(destFile.getId());
    var hojaAhorroDestino = destSS.getSheetByName('Tarjeta Ahorro');
    if (!hojaAhorroDestino) {
      sociosProcesados++;
      continue;
    }

    hojaAhorroDestino.setConditionalFormatRules([]);

    if (reglasBase && reglasBase.length > 0) {
      var reglasCopiadas = reglasBase.map(function(regla) {
        var builder = regla.copy();
        var nuevosRangos = regla.getRanges().map(function(rangoBase) {
          return hojaAhorroDestino.getRange(rangoBase.getA1Notation());
        });
        builder.setRanges(nuevosRangos);
        return builder.build();
      });
      hojaAhorroDestino.setConditionalFormatRules(reglasCopiadas);
    }

    archivosActualizados++;
    sociosProcesados++;
  }

  Logger.log('¡Copia de formatos condicionales de Tarjeta Ahorro completada!');
  Logger.log('Socios procesados: ' + sociosProcesados);
  Logger.log('Archivos actualizados: ' + archivosActualizados);
}