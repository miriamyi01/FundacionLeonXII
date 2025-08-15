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

    // Usar siempre el enlace directo de la columna
    var enlaceArchivo = (idxEnlaceArchivo !== -1 && datos[i][idxEnlaceArchivo]) ? datos[i][idxEnlaceArchivo] : '#';
    Logger.log('Enlace directo de hoja: ' + enlaceArchivo);

    // Verifica si hoy es el aniversario
    if (fechaIngreso.getDate() === hoy.getDate() && fechaIngreso.getMonth() === hoy.getMonth()) {
      var anios = anioActual - fechaIngreso.getFullYear();
      Logger.log('Procesando: ' + nombre + ' (' + correo + ') | Jefe: ' + correoJefe + ' | Años: ' + anios);

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

      Logger.log('Enviando correo a: ' + correo + ' cc: ' + correoJefe + ' | Asunto: ' + asunto);
      // Envía el correo (asegura que los parámetros sean correctos)
      if (correoJefe && correoJefe !== '') {
        var opciones = {
          htmlBody: mensaje,
          cc: correoJefe
        };
        GmailApp.sendEmail(correo, asunto, '', opciones);
      } else {
        var opciones = {
          htmlBody: mensaje
        };
        GmailApp.sendEmail(correo, asunto, '', opciones);
      }
    } else {
      Logger.log('No es aniversario de: ' + nombre + ' (' + correo + ')');
    }
  }
}