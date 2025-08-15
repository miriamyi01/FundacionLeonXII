function enviarCorreosAniversario() {
  var hoy = new Date();
  var anioActual = hoy.getFullYear();
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Aniversarios_" + anioActual);

  // Leer encabezados de la fila 2
  var encabezados = hoja.getRange(2, 1, 1, hoja.getLastColumn()).getValues()[0];
  // Leer datos desde la fila 3
  var datos = hoja.getRange(3, 1, hoja.getLastRow() - 2, hoja.getLastColumn()).getValues();

  // Cabeceras actualizadas
  var idxCorreo = encabezados.indexOf('Correo');
  var idxFechaIngreso = encabezados.indexOf('Fecha de ingreso');
  var idxEnlaceArchivo = encabezados.indexOf('Link al archivo de vacaciones');
  var idxCorreoJefe = encabezados.indexOf('Correo Jefe Directo');

  for (var i = 0; i < datos.length; i++) {
    var correo = datos[i][idxCorreo];

    // Obtener nombre a partir del correo
    function obtenerNombreDesdeCorreo(correo) {
      var parteUsuario = correo.split('@')[0];
      if (parteUsuario.indexOf('.') !== -1) {
        var partes = parteUsuario.split('.');
        return partes[1].charAt(0).toUpperCase() + partes[1].slice(1).toLowerCase();
      } else {
        return parteUsuario.charAt(0).toUpperCase() + parteUsuario.slice(1).toLowerCase();
      }
    }
    var nombre = obtenerNombreDesdeCorreo(correo);

    var fechaIngreso = new Date(datos[i][idxFechaIngreso]);

    // Obtener correo del jefe (si existe la columna)
    var correoJefe = idxCorreoJefe !== -1 ? datos[i][idxCorreoJefe] : '';

    // Usar siempre el enlace directo de la columna
    var enlaceArchivo = (idxEnlaceArchivo !== -1 && datos[i][idxEnlaceArchivo]) ? datos[i][idxEnlaceArchivo] : '#';

    // Verifica si hoy es el aniversario
    if (fechaIngreso.getDate() === hoy.getDate() && fechaIngreso.getMonth() === hoy.getMonth()) {
      var anios = anioActual - fechaIngreso.getFullYear();
      if (anios >= 1) {
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
        // Función para formatear fecha en español largo
        function fechaLarga(fecha) {
          var dias = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
          var meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
          return dias[fecha.getDay()] + ' ' + fecha.getDate() + ' de ' + meses[fecha.getMonth()] + ', ' + fecha.getFullYear();
        }

        var fechaInicioDate = new Date(anioActual, fechaIngreso.getMonth(), fechaIngreso.getDate());
        var fechaLimiteDate = new Date(anioActual + 1, fechaIngreso.getMonth(), fechaIngreso.getDate());
        fechaLimiteDate.setDate(fechaLimiteDate.getDate() - 1); // Un día antes del aniversario siguiente año

        var fechaInicio = fechaLarga(fechaInicioDate);
        var fechaLimite = fechaLarga(fechaLimiteDate);

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

        // Envía el correo (asegura que los parámetros sean correctos)
        if (correoJefe && correoJefe !== '') {
          var opciones = {
            htmlBody: mensaje,
            cc: correoJefe
          };
          GmailApp.sendEmail(correo, asunto, '', opciones);
          Logger.log('Correo enviado a: ' + correo + ' cc: ' + correoJefe);
        } else {
          var opciones = {
            htmlBody: mensaje
          };
          GmailApp.sendEmail(correo, asunto, '', opciones);
          Logger.log('Correo enviado a: ' + correo);
        }
      }
    }
  }
}