function actualizarPermisosProteccion() {
  // El archivo actual (04 TARJETA AHORRO Y PRESTAMO) es la fuente
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  var folder = null;
  while (parents.hasNext()) {
    folder = parents.next();
    break;
  }
  Logger.log("Carpeta principal: " + folder.getName());

  // Buscar archivo de inscripción para obtener la lista de socios
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
    Logger.log("No se encontró el archivo '03 LISTA DE INSCRIPCION'");
    return;
  }
  var sheet = archivoInscripcion.getSheetByName('Inscripción');
  if (!sheet) {
    Logger.log("No existe la hoja 'Inscripción'");
    return;
  }

  // Leer rangos protegidos y permisos del archivo fuente SOLO para hojas relevantes
  var baseSheets = ss.getSheets();
  var proteccionesBase = {}; // { hoja: [ {rangeA1, editors, warningOnly, description} ] }
  for (var s = 0; s < baseSheets.length; s++) {
    var hoja = baseSheets[s];
    var nombreHoja = hoja.getName();
    if (
      nombreHoja === "Tarjeta Ahorro" ||
      /^Tarjeta Prestamo\s*#?\d*$/i.test(nombreHoja)
    ) {
      var protections = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      proteccionesBase[nombreHoja] = [];
      for (var p = 0; p < protections.length; p++) {
        var prot = protections[p];
        var range = prot.getRange();
        proteccionesBase[nombreHoja].push({
          rangeA1: range.getA1Notation(),
          editors: prot.getEditors().map(function(u) { return u.getEmail(); }),
          warningOnly: prot.isWarningOnly(),
          description: prot.getDescription()
        });
      }
    }
  }

  var mainFolder = folder;
  var lastRow = sheet.getLastRow();
  var sociosProcesados = 0;
  var archivosActualizados = 0;
  var proteccionesCopiadas = 0;

  for (var i = 8; i <= lastRow; i++) {
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

    // Buscar archivo de Google Sheets
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
      sociosProcesados++;
      continue;
    }

    // Actualizar rangos protegidos en el archivo del socio
    var destSS = SpreadsheetApp.openById(destFile.getId());
    var hojas = destSS.getSheets();
    var proteccionesEnArchivo = 0;

    for (var h = 0; h < hojas.length; h++) {
      var hoja = hojas[h];
      var nombreHoja = hoja.getName();

      // Solo procesar si el nombre coincide con los criterios
      if (
        nombreHoja === "Tarjeta Ahorro" ||
        /^Tarjeta Prestamo\s*#?\d*$/i.test(nombreHoja)
      ) {
        // Eliminar todas las protecciones de rango existentes
        var protections = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        for (var p = 0; p < protections.length; p++) {
          protections[p].remove();
        }

        // Crear nuevas protecciones igual que el archivo fuente
        var protBaseArr = proteccionesBase[nombreHoja];
        if (!protBaseArr) continue;
        for (var pb = 0; pb < protBaseArr.length; pb++) {
          var base = protBaseArr[pb];
          try {
            var rango = hoja.getRange(base.rangeA1);
            var nuevaProteccion = rango.protect();
            nuevaProteccion.setDescription(base.description);
            nuevaProteccion.setWarningOnly(base.warningOnly);
            if (!base.warningOnly && base.editors.length > 0) {
              nuevaProteccion.removeEditors(nuevaProteccion.getEditors());
              nuevaProteccion.addEditors(base.editors);
            }
            proteccionesEnArchivo++;
            proteccionesCopiadas++;
          } catch (err) {
            Logger.log("Error copiando protección en hoja " + nombreHoja + " de " + carpetaSocio.getName() + ": " + err);
          }
        }
      }
    }

    if (proteccionesEnArchivo > 0) {
      archivosActualizados++;
      Logger.log("Actualizado: " + carpetaSocio.getName() + " | Protecciones copiadas: " + proteccionesEnArchivo);
    }
    sociosProcesados++;
  }

  Logger.log("¡Actualización de protecciones completada!");
  Logger.log("Socios procesados: " + sociosProcesados);
  Logger.log("Archivos actualizados: " + archivosActualizados);
  Logger.log("Protecciones de rangos copiadas: " + proteccionesCopiadas);
}