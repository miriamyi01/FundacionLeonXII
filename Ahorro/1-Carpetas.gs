function crearCarpetasSocios() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inscripción');
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Intenta obtener la carpeta contenedora, pero si no existe, muestra error y termina
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  var folder = null;
  while (parents.hasNext()) {
    folder = parents.next();
    break; // solo toma el primero
  }
  if (!folder) {
    Logger.log("El archivo no está dentro de una carpeta en Drive. Mueve el archivo a una carpeta y vuelve a intentar.");
    return;
  }
  Logger.log("Carpeta principal: " + folder.getName());

  // Busca el archivo base único en la carpeta
  var files = folder.getFiles();
  var baseFile = null;
  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().indexOf("04 TARJETA AHORRO Y PRESTAMO") !== -1) baseFile = file;
  }
  if (!baseFile) {
    Logger.log("No se encontró el archivo base 04 TARJETA AHORRO Y PRESTAMO en la carpeta.");
    return;
  }

  var baseCell = sheet.getRange("A8").getValue();
  if (!baseCell) {
    Logger.log("La celda A8 está vacía.");
    return;
  }
  var baseName = baseCell.toString().substring(0, 4);
  var mainFolderName = baseName + "-SOCIOS AS";

  var mainFolder;
  var folders = folder.getFoldersByName(mainFolderName);
  if (folders.hasNext()) {
    mainFolder = folders.next();
  } else {
    mainFolder = folder.createFolder(mainFolderName);
    Logger.log("Carpeta principal creada: " + mainFolderName);
  }

  var lastRow = sheet.getLastRow();
  var sociosCreados = 0;
  var carpetasCreadas = 0;
  var archivosCreados = 0;
  var sociosExistentes = 0;
  
  for (var i = 8; i <= lastRow; i++) {
    var nombre = sheet.getRange("B" + i).getValue().toString().trim();
    var apellido1 = sheet.getRange("C" + i).getValue().toString().trim();
    var apellido2 = sheet.getRange("D" + i).getValue().toString().trim();
    var numeroSocio = sheet.getRange("A" + i).getValue().toString().trim();

    if (!nombre) continue;

    // Iniciales: primer nombre, primer apellido, segundo apellido
    var nombres = nombre.split(" ");
    var inicialNombre = nombres[0][0] || "";
    var inicialApellido1 = apellido1[0] || "";
    var inicialApellido2 = apellido2[0] || "";
    var iniciales = (inicialNombre + inicialApellido1 + inicialApellido2).toUpperCase();

    var carpetaSocioNombre = numeroSocio + " " + iniciales + " " + nombre + " " + apellido1 + " " + apellido2;

    var carpetaSocio;
    var carpetaExistia = false;
    var subfolders = mainFolder.getFoldersByName(carpetaSocioNombre);
    if (subfolders.hasNext()) {
      carpetaSocio = subfolders.next();
      carpetaExistia = true;
    } else {
      carpetaSocio = mainFolder.createFolder(carpetaSocioNombre);
      carpetasCreadas++;
      Logger.log("Carpeta de socio creada: " + carpetaSocioNombre);
    }

    // Verificar si el archivo ya existe en la carpeta del socio
    var archivoNombre = iniciales + " - TARJETA AHORRO Y PRESTAMO";
    var archivosExistentes = carpetaSocio.getFilesByName(archivoNombre);
    var archivoExiste = archivosExistentes.hasNext();

    if (archivoExiste) {
      sociosExistentes++;
    } else {
      // Crear copia y actualizar B1 y D1 en la tarjeta de ahorro y préstamo
      var socioCopy = baseFile.makeCopy(archivoNombre, carpetaSocio);
      var socioCopyId = socioCopy.getId();
      var socioSheet = SpreadsheetApp.openById(socioCopyId).getSheets()[0];
      // Capitalizar cada palabra del nombre completo
      function capitalizar(str) {
        return str.replace(/\w\S*/g, function(txt) {
          return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
        });
      }
      var nombreCompleto = capitalizar(nombre + " " + apellido1 + " " + apellido2);
      socioSheet.getRange("B1").setValue(nombreCompleto);
      socioSheet.getRange("D1").setValue(numeroSocio); // Número de socio
      archivosCreados++;
      Logger.log("Archivo creado para el socio: " + carpetaSocioNombre);
    }

    sociosCreados++;
  }
  Logger.log("¡Proceso completado exitosamente!");
  Logger.log("Total socios procesados: " + sociosCreados);
  Logger.log("Carpetas nuevas creadas: " + carpetasCreadas);
  Logger.log("Archivos nuevos creados: " + archivosCreados);
  Logger.log("Socios que ya tenían archivo: " + sociosExistentes);
}