function crearCarpetasSocios() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inscripción');
  // Buscar la carpeta raíz GA0452 METAMORFOSIS
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

  var lastRow = sheet.getLastRow();
  var sociosCreados = 0;
  var carpetasCreadas = 0;
  var archivosCreados = 0;
  var sociosExistentes = 0;

  for (var i = 7; i <= lastRow; i++) {
    var nombre = sheet.getRange('B' + i).getValue().toString().trim();
    var apellido1 = sheet.getRange('C' + i).getValue().toString().trim();
    var apellido2 = sheet.getRange('D' + i).getValue().toString().trim();
    var numeroSocio = sheet.getRange('A' + i).getValue().toString().trim();

    if (!nombre) continue;

    // Buscar si ya existe una carpeta para este socio (por número y nombre completo, ignorando iniciales)
    var existeCarpeta = false;
    var subfolders = sociosMainFolder.getFolders();
    var nombreCompleto = (nombre + ' ' + apellido1 + ' ' + apellido2).trim().toLowerCase();
    while (subfolders.hasNext()) {
      var folder = subfolders.next();
      var folderName = folder.getName();
      var partes = folderName.split(' ');
      var folderNumero = partes[0] || '';
      // El nombre completo está después de las iniciales (índice 2 en adelante)
      var folderNombreCompleto = partes.slice(2).join(' ').trim().toLowerCase();
      if (folderNumero === numeroSocio && folderNombreCompleto === nombreCompleto) {
        existeCarpeta = true;
        sociosExistentes++;
        break;
      }
    }

    if (!existeCarpeta) {
      // Iniciales: primer nombre, primer apellido, segundo apellido
      var nombres = nombre.split(' ');
      var inicialNombre = nombres[0][0] || '';
      var inicialApellido1 = apellido1[0] || '';
      var inicialApellido2 = apellido2[0] || '';
      var iniciales = (inicialNombre + inicialApellido1 + inicialApellido2).toUpperCase();

      var carpetaSocioNombre = numeroSocio + ' ' + iniciales + ' ' + nombre + ' ' + apellido1 + ' ' + apellido2;
      sociosMainFolder.createFolder(carpetaSocioNombre);
      carpetasCreadas++;
      Logger.log('Carpeta de socio creada: ' + carpetaSocioNombre);
    }
    sociosCreados++;
  }
  Logger.log('¡Proceso completado exitosamente!');
  Logger.log('Total socios procesados: ' + sociosCreados);
  Logger.log('Carpetas nuevas creadas: ' + carpetasCreadas);
  Logger.log('Archivos nuevos creados: ' + archivosCreados);
  Logger.log('Socios que ya tenían archivo: ' + sociosExistentes);
}