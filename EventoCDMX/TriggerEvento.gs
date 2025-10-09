function enviarEvento() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2025');
  var datos = hoja.getDataRange().getValues();
  var encabezados = datos[0]; // Asume encabezados en la fila 1
  var idxCorreo = encabezados.indexOf('Correo');

  for (var i = 1; i < datos.length; i++) { // Empieza en la fila 2
    var correo = datos[i][idxCorreo];
    if (!correo) continue;

    // Obtener y normalizar el primer nombre desde la columna 'Nombre corto'
    var idxNombreCorto = encabezados.indexOf('Nombre corto');
    var nombreCorto = '';
    if (idxNombreCorto !== -1) {
      nombreCorto = datos[i][idxNombreCorto] || '';
    }
    // Tomar solo el primer nombre
    var primerNombre = nombreCorto.trim().split(' ')[0] || '';
    // Normalizar: primera letra mayúscula, resto minúsculas
    var nombre = primerNombre.charAt(0).toUpperCase() + primerNombre.slice(1).toLowerCase();

    // Datos para la plantilla
    var html = HtmlService.createTemplateFromFile('CorreoEvento');
    html.nombre = nombre;
    var mensaje = html.evaluate().getContent();
    var asunto = "Invitación - Encuentro de promotores 2025 ✨";

    var opciones = {
      htmlBody: mensaje
    };
    GmailApp.sendEmail(correo, asunto, '', opciones);
    Logger.log('Correo enviado a: ' + correo);
  }
}