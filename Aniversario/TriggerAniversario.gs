// Ejecuta esta función una sola vez para crear el trigger diario a las 9:00 am
function crearTriggerDiarioAniversario() {
  // Elimina triggers previos de esta función para evitar duplicados
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'enviarCorreosAniversario') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Crea el trigger diario a las 9:00 am (hora del script)
  ScriptApp.newTrigger('enviarCorreosAniversario')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
}
function enviarCorreosAniversario() {
  var hoy = new Date();
  var anioActual = hoy.getFullYear();
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(anioActual.toString());
  var datos = hoja.getDataRange().getValues();

  // Cabeceras
  var idxNombre = datos[0].indexOf('Nombre');
  var idxCorreo = datos[0].indexOf('Correo');
  var idxFechaIngreso = datos[0].indexOf('FechaIngreso');
  var idxEnlaceArchivo = datos[0].indexOf('EnlaceArchivoVacaciones');
  // Puedes agregar más índices si tienes más columnas


  for (var i = 1; i < datos.length; i++) {
    var nombre = datos[i][idxNombre];
    var correo = datos[i][idxCorreo];
    var fechaIngreso = new Date(datos[i][idxFechaIngreso]);

    // Obtener correo del jefe (si existe la columna)
    var idxCorreoJefe = datos[0].indexOf('CorreoJefe');
    var correoJefe = idxCorreoJefe !== -1 ? datos[i][idxCorreoJefe] : '';

    // Obtener iniciales del nombre (primeras letras de cada palabra)
    function obtenerIniciales(nombre) {
      return nombre.split(/\s+/).map(function(palabra) {
        return palabra.charAt(0).toUpperCase();
      }).join('');
    }
    var iniciales = obtenerIniciales(nombre);

    // Suponiendo que hay una columna con el número de empleado
    var idxNumEmpleado = datos[0].indexOf('NumEmpleado');
    var numEmpleado = idxNumEmpleado !== -1 ? datos[i][idxNumEmpleado] : '';

    // Formato del archivo: año_iniciales_numerodeempleado
    var nombreArchivo = anioActual + '_' + iniciales + '_' + numEmpleado;

    // Buscar el archivo en Google Drive
    var enlaceArchivo = '';
    var archivos = DriveApp.getFilesByName(nombreArchivo);
    if (archivos.hasNext()) {
      var archivo = archivos.next();
      enlaceArchivo = archivo.getUrl();
    } else {
      // Si no se encuentra, dejar vacío o poner un mensaje alternativo
      enlaceArchivo = '#';
    }

    // Verifica si hoy es el aniversario
    if (fechaIngreso.getDate() === hoy.getDate() && fechaIngreso.getMonth() === hoy.getMonth()) {
      var anios = anioActual - fechaIngreso.getFullYear();

      // Cálculo de vacaciones según la tabla proporcionada
      var diasVacaciones = 0;
      if (anios === 1) diasVacaciones = 12;
      else if (anios === 2) diasVacaciones = 14;
      else if (anios === 3) diasVacaciones = 16;
      else if (anios === 4) diasVacaciones = 18;
      else if (anios === 5) diasVacaciones = 20;
      else if (anios >= 6 && anios <= 10) diasVacaciones = 22;
      else if (anios >= 11 && anios <= 15) diasVacaciones = 24;
      else if (anios >= 16 && anios <= 20) diasVacaciones = 26;
      else if (anios >= 21 && anios <= 25) diasVacaciones = 28;
      else if (anios >= 26 && anios <= 30) diasVacaciones = 30;
      var fechaInicio = Utilities.formatDate(new Date(anioActual, fechaIngreso.getMonth(), fechaIngreso.getDate()), Session.getScriptTimeZone(), "yyyy-MM-dd");
      var fechaLimite = Utilities.formatDate(new Date(anioActual + 1, fechaIngreso.getMonth(), fechaIngreso.getDate() - 1), Session.getScriptTimeZone(), "yyyy-MM-dd");

      // Datos para la plantilla
      var html = HtmlService.createTemplateFromFile('CorreoAniversario');
      html.nombre = nombre;
      html.anios = anios;
      html.diasVacaciones = diasVacaciones;
      html.fechaInicio = fechaInicio;
      html.fechaLimite = fechaLimite;
      html.enlaceArchivo = enlaceArchivo;

      var mensaje = html.evaluate().getContent();

      // Asunto personalizado
      var asunto = "Felicidades " + nombre;

      // Envía el correo
      GmailApp.sendEmail({
        to: correo,
        cc: correoJefe,
        subject: asunto,
        htmlBody: mensaje
      });
    }
  }
}