function buscarLinksArchivos() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var datos = hoja.getDataRange().getValues();
  // Obtener la carpeta donde está la hoja de cálculo activa
  var fileId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var file = DriveApp.getFileById(fileId);
  var parentFolders = file.getParents();
  if (!parentFolders.hasNext()) {
    Logger.log('ERROR: No se encontró la carpeta contenedora de la hoja de cálculo');
    return;
  }
  var folder = parentFolders.next();
  Logger.log('Buscando en carpeta: ' + folder.getName());
  var archivos = [];

  // Función recursiva para indexar todos los archivos en subcarpetas
  function indexarArchivos(carpeta) {
    var files = carpeta.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      archivos.push({
        nombre: f.getName(),
        url: f.getUrl()
      });
    }
    var subfolders = carpeta.getFolders();
    while (subfolders.hasNext()) {
      indexarArchivos(subfolders.next());
    }
  }
  indexarArchivos(folder);
  Logger.log('Total archivos indexados: ' + archivos.length);

  // Definición de sufijos y columnas destino
  var criterios = [
    { sufijos: [' - CV'], columna: 16 }, // P (índice 16)
    { sufijos: [' - CARTA PRESENTACION', ' - PLAN DE TRABAJO'], columna: 17 }, // Q
    { sufijos: [' - ACUERDO', ' - CARTA ACEPTACION'], columna: 18 },         // R
    { sufijos: [' - RESPONSIVA'], columna: 19 },      // S
    { sufijos: [' - REPORTE 1 DE 3'], columna: 21 },  // U
    { sufijos: [' - REPORTE 2 DE 3'], columna: 23 },  // W
    { sufijos: [' - REPORTE 3 DE 3', ' - REPORTE FINAL'], columna: 25 },   // Y
    { sufijos: [' - CARTA LIBERACION'], columna: 27 },// AA
    { sufijos: [' - SEGUIMIENTO'], columna: 28 },     // AB
    { sufijos: [' - FOTOGRAFIA'], columna: 29 }       // AC
  ];

  // Función para normalizar texto: mayúsculas y sin acentos
  function normalizar(texto) {
    return texto ? texto.normalize('NFD').replace(/\p{Diacritic}/gu, '').toUpperCase() : '';
  }

  for (var i = 1; i < datos.length; i++) {
    var nombre = datos[i][0]; // Columna A (nombre)
    if (!nombre) continue;
    var nombreNorm = normalizar(nombre);
    criterios.forEach(function(criterio) {
      var celda = hoja.getRange(i + 1, criterio.columna);
      var valorActual = celda.getValue();
      if (valorActual && typeof valorActual === 'string' && valorActual.startsWith('http')) return; // Solo deja intactos los enlaces
      // Filtra todos los archivos que cumplen alguno de los sufijos del criterio
      var candidatos = archivos.filter(function(archivo) {
        var base = archivo.nombre.replace(/\.[^/.]+$/, "");
        var baseNorm = normalizar(base);
        return baseNorm.indexOf(nombreNorm) !== -1 &&
          criterio.sufijos.some(function(suf) {
            return baseNorm.endsWith(normalizar(suf));
          });
      });
      if (candidatos.length > 0) {
        // Prioriza PDFs
        var pdfs = candidatos.filter(function(archivo) {
          return archivo.nombre.toLowerCase().endsWith('.pdf');
        });
        var elegidos = pdfs.length > 0 ? pdfs : candidatos;
        // Selecciona el más reciente
        var elegido = elegidos.reduce(function(a, b) {
          var fa = DriveApp.getFileById(a.url.split('/d/')[1]?.split('/')[0] || '');
          var fb = DriveApp.getFileById(b.url.split('/d/')[1]?.split('/')[0] || '');
          try {
            return fa.getLastUpdated() > fb.getLastUpdated() ? a : b;
          } catch (e) {
            return a; // Si hay error, deja el primero
          }
        }, elegidos[0]);
        celda.setValue(elegido.url);
      } else {
        celda.setValue('No');
      }
    });
  }
}