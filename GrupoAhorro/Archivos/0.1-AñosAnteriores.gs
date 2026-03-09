function moverYOrganizarAnios() {
  // --- Movimiento de listas años anteriores ---
  var anioActual = new Date().getFullYear();
  var nombreCarpetaRaiz = "GA0452-FICHA DE INSCRIPCIÓN *";
  var nombreCarpetaPadre = "GA0452 METAMORFOSIS";
  var nombreCarpetaDestino = "01 LISTAS AÑOS ANTERIORES";
  var nombreExcluido = "04 TARJETA AHORRO Y PRESTAMO";

  var carpetasPadre = DriveApp.getFoldersByName(nombreCarpetaPadre);
  if (!carpetasPadre.hasNext()) {
    Logger.log("No se encontró la carpeta padre: " + nombreCarpetaPadre);
    return;
  }
  var carpetaPadre = carpetasPadre.next();

  var carpetasInicial = carpetaPadre.getFoldersByName(nombreCarpetaRaiz);
  if (!carpetasInicial.hasNext()) {
    Logger.log("No se encontró la carpeta inicial: " + nombreCarpetaRaiz);
    return;
  }
  var carpetaInicial = carpetasInicial.next();

  var carpetasDestino = carpetaInicial.getFoldersByName(nombreCarpetaDestino);
  if (!carpetasDestino.hasNext()) {
    Logger.log("No se encontró la carpeta destino: " + nombreCarpetaDestino);
    return;
  }
  var carpetaDestino = carpetasDestino.next();

  var archivos = carpetaInicial.getFiles();
  while (archivos.hasNext()) {
    var archivo = archivos.next();
    var nombre = archivo.getName();
    if (
      nombre.indexOf(anioActual) === -1 &&
      nombre !== nombreExcluido
    ) {
      archivo.moveTo(carpetaDestino);
      Logger.log('Archivo movido a listas años anteriores: ' + nombre);
    }
  }

  // --- Organización de ahorro por años en socios ---
  var rootFolderName = 'GA0452 METAMORFOSIS';
  var sociosFolderName = 'GA0452-SOCIOS AS';
  var ahorroAnterioresFolderName = '01 AHORRO AÑOS ANTERIORES';
  // Genera nombres de carpetas para los dos años anteriores con sufijo consecutivo
  var anios = [anioActual - 2, anioActual - 1];
  function obtenerSufijo(folder, anio) {
    var contador = 1;
    var subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      var subf = subFolders.next();
      var nombre = normalizarNombre(subf.getName());
      var regex = new RegExp(anio + ' (\d{2}) ahorro');
      var match = nombre.match(regex);
      if (match) {
        var num = parseInt(match[1], 10);
        if (num >= contador) contador = num + 1;
      }
    }
    return (contador < 10 ? '0' + contador : contador);
  }

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

  // Recorrer todas las carpetas de socios
  var socioFoldersIter = sociosMainFolder.getFolders();
  while (socioFoldersIter.hasNext()) {
    var socioFolder = socioFoldersIter.next();
    Logger.log('Procesando socio: ' + socioFolder.getName());

    // Buscar o crear carpeta 01 AHORRO AÑOS ANTERIORES
    var ahorroAnterioresFolder = null;
    var socioSubFolders = socioFolder.getFolders();
    while (socioSubFolders.hasNext()) {
      var subf = socioSubFolders.next();
      if (normalizarNombre(subf.getName()) === normalizarNombre(ahorroAnterioresFolderName)) {
        ahorroAnterioresFolder = subf;
        break;
      }
    }
    if (!ahorroAnterioresFolder) {
      ahorroAnterioresFolder = socioFolder.createFolder(ahorroAnterioresFolderName);
      Logger.log('Creada carpeta: ' + ahorroAnterioresFolderName);
    }

    var ahorroFoldersCreated = [];
    // Crear carpetas de los dos años anteriores si no existen
    anios.forEach(function(anio) {
      var sufijo = obtenerSufijo(ahorroAnterioresFolder, anio);
      var folderName = anio + ' ' + sufijo + ' AHORRO';
      var folder = null;
      var subFolders = ahorroAnterioresFolder.getFolders();
      while (subFolders.hasNext()) {
        var subf = subFolders.next();
        if (normalizarNombre(subf.getName()) === normalizarNombre(folderName)) {
          folder = subf;
          break;
        }
      }
      if (!folder) {
        folder = ahorroAnterioresFolder.createFolder(folderName);
        Logger.log('Creada carpeta: ' + folderName);
      }
      ahorroFoldersCreated.push(folder);
    });

    // Mover archivos de 01 AHORRO AÑOS ANTERIORES a la carpeta del año más antiguo
    var archivosAhorroAnteriores = ahorroAnterioresFolder.getFiles();
    while (archivosAhorroAnteriores.hasNext()) {
      var archivo = archivosAhorroAnteriores.next();
      archivo.moveTo(ahorroFoldersCreated[0]);
    }

    // Mover archivos fuera de 01 AHORRO AÑOS ANTERIORES a la carpeta del año anterior más reciente
    var socioFiles = socioFolder.getFiles();
    while (socioFiles.hasNext()) {
      var archivo = socioFiles.next();
      archivo.moveTo(ahorroFoldersCreated[1]);
    }
  }
}