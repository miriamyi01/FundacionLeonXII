function actualizarPermisosProteccion() {
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
  var ss = SpreadsheetApp.openById(baseFileId);
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

  // Leer rangos protegidos y permisos del archivo fuente SOLO para hojas relevantes
  var baseSheets = ss.getSheets();
  var proteccionesBase = {}; // { hoja: [ {rangeA1, editors, warningOnly, description} ] }
  for (var s = 0; s < baseSheets.length; s++) {
    var hoja = baseSheets[s];
    var nombreHoja = hoja.getName();
    if (
      nombreHoja === "Tarjeta Ahorro" ||
      nombreHoja === "Fondo de Emergencia" ||
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

  var mainFolder = sociosMainFolder;
  var lastRow = sheet.getLastRow();
  var sociosProcesados = 0;
  var archivosActualizados = 0;
  var proteccionesCopiadas = 0;

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

    // Actualizar rangos protegidos en el archivo del socio
    var destSS = SpreadsheetApp.openById(destFile.getId());
    var hojas = destSS.getSheets();
    var proteccionesEnArchivo = 0;

    for (var h = 0; h < hojas.length; h++) {
      var hoja = hojas[h];
      var nombreHoja = hoja.getName();

      if (
        nombreHoja === "Tarjeta Ahorro" ||
        nombreHoja === "Fondo de Emergencia" ||
        /^Tarjeta Prestamo\s*#?\d*$/i.test(nombreHoja)
      ) {
        var protections = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        var protBaseArr = proteccionesBase[nombreHoja];
        if (!protBaseArr) continue;

        // Crear un mapa para comparar protecciones existentes
        var existingMap = {};
        for (var p = 0; p < protections.length; p++) {
          var prot = protections[p];
          var key = prot.getRange().getA1Notation();
          existingMap[key] = prot;
        }

        for (var pb = 0; pb < protBaseArr.length; pb++) {
          var base = protBaseArr[pb];
          var match = existingMap[base.rangeA1];

          // Función para comparar arrays de editores
          function arraysEqual(a, b) {
            if (a.length !== b.length) return false;
            a = a.slice().sort(); b = b.slice().sort();
            for (var i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
            return true;
          }

          var needsUpdate = true;
          if (match) {
            // Compara descripción, warningOnly y editores
            var editors = match.getEditors().map(function(u){return u.getEmail();});
            if (
              match.getDescription() === base.description &&
              match.isWarningOnly() === base.warningOnly &&
              arraysEqual(editors, base.editors)
            ) {
              needsUpdate = false; // Ya existe igual, no hacer nada
            } else {
              match.remove(); // Si es diferente, eliminarla
            }
          }

          if (needsUpdate) {
            try {
              var rango = hoja.getRange(base.rangeA1);
              var nuevaProteccion = rango.protect();
              nuevaProteccion.setDescription(base.description);
              nuevaProteccion.setWarningOnly(false); // Quitar el warning, solo los editores pueden modificar
              if (base.editors.length > 0) {
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

        // Elimina protecciones que no están en el archivo fuente
        for (var p = 0; p < protections.length; p++) {
          var prot = protections[p];
          var key = prot.getRange().getA1Notation();
          var found = protBaseArr.some(function(b){return b.rangeA1 === key;});
          if (!found) prot.remove();
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