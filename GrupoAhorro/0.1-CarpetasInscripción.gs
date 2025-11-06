function llenarHojaInscripcion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Inscripción');
  
  if (!sheet) {
    Logger.log("No existe la hoja 'Inscripción'. Créala primero.");
    return;
  }

  // Obtener la carpeta contenedora
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  var folder = null;
  while (parents.hasNext()) {
    folder = parents.next();
    break;
  }
  if (!folder) {
    Logger.log("El archivo no está dentro de una carpeta en Drive.");
    return;
  }
  Logger.log("Carpeta principal: " + folder.getName());

  var mainFolder = folder;
  var carpetas = [];
  
  // Recolectar todas las carpetas que parecen ser de socios (empiezan con GA)
  var subIt = mainFolder.getFolders();
  while (subIt.hasNext()) {
    var sf = subIt.next();
    var nombreCarpeta = sf.getName().trim();
    if (nombreCarpeta.toUpperCase().indexOf("GA") === 0) {
      carpetas.push(nombreCarpeta);
    }
  }
  
  Logger.log("Carpetas encontradas: " + carpetas.length);
  
  // Ordenar carpetas por número de socio
  carpetas.sort();
  
  var filaActual = 8; // Empezar en la fila 8
  
  for (var i = 0; i < carpetas.length; i++) {
    var nombreCarpeta = carpetas[i];
    
    // Parsear el nombre: "GA0452041 VES VIDAL DE JESUS ESTRADA SANTIZ"
    var parts = nombreCarpeta.split(/\s+/);
    
    if (parts.length < 4) {
      Logger.log("Formato incorrecto en carpeta: " + nombreCarpeta);
      continue;
    }
    
    var numeroSocio = parts[0];
    var iniciales = parts[1];
    
    var restoParts = parts.slice(2);
    
    if (restoParts.length < 3) {
      Logger.log("No hay suficientes partes para nombre y apellidos en: " + nombreCarpeta);
      continue;
    }
    
    var segundoApellido = restoParts[restoParts.length - 1];
    var primerApellido = restoParts[restoParts.length - 2];
    
    var nombres = restoParts.slice(0, restoParts.length - 2).join(" ");
    
    // Escribir en la hoja
    sheet.getRange("A" + filaActual).setValue(numeroSocio);
    sheet.getRange("B" + filaActual).setValue(nombres);
    sheet.getRange("C" + filaActual).setValue(primerApellido);
    sheet.getRange("D" + filaActual).setValue(segundoApellido);
    sheet.getRange("E" + filaActual).setValue(""); // ZONA vacío
    sheet.getRange("F" + filaActual).setValue(0); // APORTACIÓN FONDO SOCIAL $0.0
    
    Logger.log("Fila " + filaActual + ": " + numeroSocio + " | " + nombres + " " + primerApellido + " " + segundoApellido);
    filaActual++;
  }
  
  Logger.log("¡Proceso completado! Se llenaron " + (filaActual - 8) + " registros.");
}
