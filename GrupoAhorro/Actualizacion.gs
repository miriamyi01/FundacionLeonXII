function renombrarYAgregarHojas() {
  // Buscar el archivo llamado exactamente '03 LISTA DE INSCRIPCION' en la misma carpeta
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  var folder = null;
  while (parents.hasNext()) {
    folder = parents.next();
    break;
  }
  Logger.log("Carpeta principal: " + folder.getName());

  var archivoInscripcion = null;
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    try {
      if (
        file.getMimeType() === MimeType.GOOGLE_SHEETS &&
        file.getName().trim().toUpperCase() === '03 LISTA DE INSCRIPCION'
      ) {
        archivoInscripcion = SpreadsheetApp.openById(file.getId());
        break;
      }
    } catch (e) {}
  }
  if (!archivoInscripcion) {
    Logger.log("No se encontró ningún archivo llamado '03 LISTA DE INSCRIPCION' en la carpeta.");
    return;
  }
  var sheet = archivoInscripcion.getSheetByName('Inscripción');
  if (!sheet) {
    Logger.log("No existe la hoja 'Inscripción' en el archivo '03 LISTA DE INSCRIPCION'. Verifica el nombre exacto de la hoja.");
    return;
  }

  var mainFolder = folder;
  var baseSS = ss;
  var baseSheets = baseSS.getSheets();

  var lastRow = sheet.getLastRow();
  var sociosProcesados = 0;
  var carpetasEncontradas = 0;
  var hojasAgregadasTotal = 0;
  var archivosRenombrados = 0;
  var carpetasNoEncontradas = 0;
  var archivosNoEncontrados = 0;

  for (var i = 8; i <= lastRow; i++) {
    // Lee el número de socio desde la columna A, fila i
    var numeroSocio = sheet.getRange("A" + i).getValue().toString().trim();
    var nombreB = sheet.getRange("B" + i).getValue().toString().trim();
    var nombreC = sheet.getRange("C" + i).getValue().toString().trim();
    var nombreD = sheet.getRange("D" + i).getValue().toString().trim();
    var nombreCompleto = [nombreB, nombreC, nombreD].filter(Boolean).join(' ').trim();
    if (!nombreCompleto) continue;

    var numeroU = numeroSocio.toUpperCase();

    // Busca la carpeta cuyo nombre comienza con el número de socio seguido de espacio
    var carpetaSocio = null;
    var subIt = mainFolder.getFolders();
    while (subIt.hasNext()) {
      var sf = subIt.next();
      var sfName = sf.getName().toUpperCase();
      if (sfName.indexOf(numeroU + " ") === 0) {
        carpetaSocio = sf;
        break;
      }
    }
    if (!carpetaSocio) {
      carpetasNoEncontradas++;
      Logger.log("No se encontró carpeta para número: " + numeroU);
      sociosProcesados++;
      continue;
    }
    carpetasEncontradas++;

    // Extraer iniciales (segunda palabra del nombre de carpeta)
    // Ej: "GA0452041 VES VIDAL..." -> "VES"
    var parts = carpetaSocio.getName().trim().split(/\s+/);
    var inicialesFromFolder = (parts.length > 1 ? parts[1].replace(/[^A-Za-z]/g, '') : '');
    var inicialesUpper = inicialesFromFolder.toUpperCase();

    // Buscar el archivo de Google Sheets dentro de la carpeta del socio
    var destFile = null;
    var filesIt = carpetaSocio.getFiles();
    while (filesIt.hasNext()) {
      var f = filesIt.next();
      try {
        if (f.getMimeType && f.getMimeType() === MimeType.GOOGLE_SHEETS) {
          destFile = f;
          break;
        }
      } catch (e) {}
    }
    if (!destFile) {
      archivosNoEncontrados++;
      Logger.log("No se encontró archivo de Google Sheets en la carpeta: " + carpetaSocio.getName());
      sociosProcesados++;
      continue;
    }

    // Copiar las hojas del archivo actual al archivo del socio (evitando duplicados)
    var destSS = SpreadsheetApp.openById(destFile.getId());
    var hojasAgregadas = 0;
    for (var s = 0; s < baseSheets.length; s++) {
      var baseSh = baseSheets[s];
      var baseNameSh = baseSh.getName();
      var newSh;
      if (destSS.getSheetByName(baseNameSh)) {
        // Si la hoja ya existe y es 'Tarjeta Ahorro', actualiza los datos si faltan o no son correctos
        if (baseNameSh === 'Tarjeta Ahorro') {
          newSh = destSS.getSheetByName(baseNameSh);
          // Si B1 no es igual al nombre completo, actualízalo
          if (newSh.getRange('B1').getValue() !== nombreCompleto) {
            newSh.getRange('B1').setValue(nombreCompleto);
          }
          // Si D1 está vacío, pon el código del socio
          if (!newSh.getRange('D1').getValue()) {
            newSh.getRange('D1').setValue(numeroSocio);
          }
        }
        continue;
      }
      newSh = baseSh.copyTo(destSS);
      try { newSh.setName(baseNameSh); } catch (e) {}
      hojasAgregadas++;
      hojasAgregadasTotal++;

      // Solo para la hoja 'Tarjeta Ahorro'
      if (baseNameSh === 'Tarjeta Ahorro') {
        newSh.getRange('B1').setValue(nombreCompleto);
        if (!newSh.getRange('D1').getValue()) {
          newSh.getRange('D1').setValue(numeroSocio);
        }
      }
    }

    // Renombrar archivo: "(INICIALES) - TARJETA AHORRO Y PRESTAMO"
    var nuevoNombreArchivo = inicialesUpper + " - TARJETA AHORRO Y PRESTAMO";
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