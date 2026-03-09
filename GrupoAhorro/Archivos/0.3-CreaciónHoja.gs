function renombrarYAgregarHojas() {
  // Buscar la carpeta raíz GA0452 METAMORFOSIS
  var rootFolderName = 'GA0452 METAMORFOSIS';
  var sociosFolderName = 'GA0452-SOCIOS AS';
  var baseFileName = '04 TARJETA AHORRO Y PRESTAMO';

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

  // Abrir el archivo base por nombre en la carpeta del script
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
  var baseSheets = baseSS.getSheets();

  var lastRow = sheet.getLastRow();
  var sociosProcesados = 0;
  var carpetasEncontradas = 0;
  var hojasAgregadasTotal = 0;
  var archivosRenombrados = 0;
  var carpetasNoEncontradas = 0;
  var archivosNoEncontrados = 0;

  for (var i = 7; i <= lastRow; i++) {
    // Lee el número de socio desde la columna A, fila i
    var numeroSocio = sheet.getRange("A" + i).getValue().toString().trim();
    var nombreB = sheet.getRange("B" + i).getValue().toString().trim();
    var nombreC = sheet.getRange("C" + i).getValue().toString().trim();
    var nombreD = sheet.getRange("D" + i).getValue().toString().trim();
    var nombreCompleto = [nombreB, nombreC, nombreD].filter(Boolean).join(' ').trim();
    if (!nombreCompleto) continue;

    var numeroU = numeroSocio.toUpperCase();

    // Buscar la carpeta del socio por código y nombre completo, ignorando iniciales
    var apellido1 = sheet.getRange('C' + i).getValue().toString().trim();
    var apellido2 = sheet.getRange('D' + i).getValue().toString().trim();
    var nombreCompletoSocio = (nombreB + ' ' + apellido1 + ' ' + apellido2).trim().toLowerCase();
    var carpetaSocio = null;
    var subfolders = sociosMainFolder.getFolders();
    while (subfolders.hasNext()) {
      var folder = subfolders.next();
      var folderName = folder.getName();
      var partes = folderName.split(' ');
      var folderCodigo = partes[0] || '';
      var folderNombreCompleto = partes.slice(2).join(' ').trim().toLowerCase();
      if (folderCodigo === numeroSocio && folderNombreCompleto === nombreCompletoSocio) {
        carpetaSocio = folder;
        break;
      }
    }
    if (!carpetaSocio) {
      Logger.log('No se encontró carpeta para: ' + numeroSocio + ' ' + nombreCompletoSocio);
      carpetasNoEncontradas++;
      sociosProcesados++;
      continue;
    }
    carpetasEncontradas++;

    // El archivo base ya está abierto como baseSS y sus hojas en baseSheets
    var baseFile = DriveApp.getFileById(baseFileId);

    // Usar las iniciales de la carpeta del socio para el nombre del archivo
    var folderNameParts = carpetaSocio.getName().split(' ');
    var inicialesCarpeta = folderNameParts[1] || '';

    // Verificar si el archivo ya existe en la carpeta del socio
    var archivoNombre = inicialesCarpeta + ' - TARJETA AHORRO Y PRESTAMO';
    var archivosExistentes = carpetaSocio.getFilesByName(archivoNombre);
    var destFile = null;
    if (archivosExistentes.hasNext()) {
      destFile = archivosExistentes.next();
      Logger.log('Archivo ya existía: ' + archivoNombre + ' en ' + carpetaSocio.getName());
    } else {
      destFile = baseFile.makeCopy(archivoNombre, carpetaSocio);
      Logger.log('Archivo creado: ' + archivoNombre + ' en ' + carpetaSocio.getName());
    }

    // Copiar las hojas del archivo actual al archivo del socio (evitando duplicados)
    var destSS = SpreadsheetApp.openById(destFile.getId());
    var hojasAgregadas = 0;
    for (var s = 0; s < baseSheets.length; s++) {
      var baseSh = baseSheets[s];
      var baseNameSh = baseSh.getName();
      var newSh;
      if (destSS.getSheetByName(baseNameSh)) {
        // Si la hoja ya existe y es 'Tarjeta Ahorro' o 'Fondo de Emergencia', actualiza los datos si faltan o no son correctos
        if (baseNameSh === 'Tarjeta Ahorro' || baseNameSh === 'Fondo de Emergencia') {
          newSh = destSS.getSheetByName(baseNameSh);
          if (newSh.getRange('B1').getValue() !== nombreCompleto) {
            newSh.getRange('B1').setValue(nombreCompleto);
          }
          if (!newSh.getRange('D1').getValue()) {
            newSh.getRange('D1').setValue(numeroSocio);
          }
          // Solo para 'Tarjeta Ahorro', también llena F1 con la fórmula
          if (baseNameSh === 'Tarjeta Ahorro') {
            var ssId = baseSS.getId();
            var formula = '=IMPORTRANGE("' + ssId + '", "Inscripción!F' + i + '")';
            newSh.getRange('F1').setFormula(formula);
          }
        }
        continue;
      }
      newSh = baseSh.copyTo(destSS);
      try { newSh.setName(baseNameSh); } catch (e) {}
      hojasAgregadas++;
      hojasAgregadasTotal++;
      if (baseNameSh === 'Tarjeta Ahorro' || baseNameSh === 'Fondo de Emergencia') {
        newSh.getRange('B1').setValue(nombreCompleto);
        if (!newSh.getRange('D1').getValue()) {
          newSh.getRange('D1').setValue(numeroSocio);
        }
        // Solo para 'Tarjeta Ahorro', también llena F1 con la fórmula
        if (baseNameSh === 'Tarjeta Ahorro') {
          var ssId2 = baseSS.getId();
          var formula2 = '=IMPORTRANGE("' + ssId2 + '", "Inscripción!F' + i + '")';
          newSh.getRange('F1').setFormula(formula2);
        }
      }
    }

    // Renombrar archivo: "(INICIALES) - TARJETA AHORRO Y PRESTAMO"
    var nuevoNombreArchivo = inicialesCarpeta + " - TARJETA AHORRO Y PRESTAMO";
    if (destFile.getName() !== nuevoNombreArchivo) {
      destFile.setName(nuevoNombreArchivo);
      archivosRenombrados++;
    }

    Logger.log("Actualizado: " + carpetaSocio.getName() + " | Hojas agregadas: " + hojasAgregadas + " | Nuevo nombre: " + nuevoNombreArchivo);
    sociosProcesados++;
  }

  Logger.log("¡Proceso completado!");
  Logger.log("Socios procesados: " + sociosProcesados);
  Logger.log("Carpetas encontradas: " + carpetasEncontradas + " | No encontradas: " + carpetasNoEncontradas);
  Logger.log("Archivos renombrados: " + archivosRenombrados + " | Carpetas sin archivo Sheets: " + archivosNoEncontrados);
  Logger.log("Total de hojas agregadas: " + hojasAgregadasTotal);
}