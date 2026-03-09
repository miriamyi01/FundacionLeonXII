// Reemplaza todas las hojas 'Tarjeta Prestamo #n' en cada archivo de socio por la hoja 'Tarjeta Prestamo' del archivo base
function reemplazarHojaPrestamo1Socios() {
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
  var hojaPrestamoBase = ssBase.getSheetByName('Tarjeta Prestamo #1');
  if (!hojaPrestamoBase) {
    Logger.log('No se encontró la hoja "Tarjeta Prestamo #1" en el archivo base.');
    return;
  }

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

    // Eliminar solo la hoja 'Tarjeta Prestamo #1'
    for (var h = hojas.length - 1; h >= 0; h--) {
      var hoja = hojas[h];
      var nombreHoja = hoja.getName();
      if (nombreHoja === 'Tarjeta Prestamo #1') {
        destSS.deleteSheet(hoja);
      }
    }

    // Copiar la hoja 'Tarjeta Prestamo #1' del archivo base
    var nuevaHoja = hojaPrestamoBase.copyTo(destSS);
    nuevaHoja.setName('Tarjeta Prestamo #1');

    // Reordenar las hojas: Tarjeta Ahorro, Tarjeta Prestamo #1, Fondo de Emergencia
    var hojaAhorro = destSS.getSheetByName('Tarjeta Ahorro');
    var hojaFondo = destSS.getSheetByName('Fondo de Emergencia');
    // Si existen las tres hojas, reordenarlas
    if (hojaAhorro && nuevaHoja && hojaFondo) {
      destSS.setActiveSheet(hojaAhorro);
      destSS.moveActiveSheet(1);
      destSS.setActiveSheet(nuevaHoja);
      destSS.moveActiveSheet(2);
      destSS.setActiveSheet(hojaFondo);
      destSS.moveActiveSheet(3);
    }

    archivosActualizados++;
    Logger.log("Actualizado: " + carpetaSocio.getName());
    sociosProcesados++;
  }

  Logger.log("¡Reemplazo de hoja de préstamo #1 completado!");
  Logger.log("Socios procesados: " + sociosProcesados);
  Logger.log("Archivos actualizados: " + archivosActualizados);
}