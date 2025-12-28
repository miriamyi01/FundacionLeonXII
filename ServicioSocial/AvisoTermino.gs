function avisoTerminoServicioSocial() {
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	var data = sheet.getDataRange().getValues();
	var today = new Date();
	var email = "colaboradores@fundacionleontrece.org"; // Correo al que va dirigido
	var calendarId = "colaboradores@fundacionleontrece.org"; // Correo como ID de calendario principal

	// Función auxiliar para formatear la fecha en español
	function formatDateEs(date, timeZone) {
		var days = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
		var months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
		var d = Utilities.formatDate(date, timeZone, "yyyy-MM-dd");
		var parts = d.split('-');
		var jsDate = new Date(parts[0], parts[1] - 1, parts[2]);
		var dayName = days[jsDate.getDay()];
		var day = jsDate.getDate();
		var monthName = months[jsDate.getMonth()];
		var year = jsDate.getFullYear();
		return dayName.charAt(0).toUpperCase() + dayName.slice(1) + ", " + day + " de " + monthName + " de " + year;
	}

	// Procesar todos los registros (excepto encabezado)
	for (var i = 1; i < data.length; i++) {
		var nombre = data[i][0]; // Columna A
		var universidad = data[i][4]; // Columna E
		var tipoPrograma = data[i][7]; // Columna H
		var licenciatura = data[i][6]; // Columna G
		var fechaTermino = data[i][11]; // Columna L
		var fechaInicio = data[i][10]; // Columna K
		if (!fechaTermino || Object.prototype.toString.call(fechaTermino) !== '[object Date]') continue;

		// Normaliza la fecha (sin horas)
		var fechaTerminoSimple = new Date(fechaTermino.getFullYear(), fechaTermino.getMonth(), fechaTermino.getDate());
		var hoySimple = new Date(today.getFullYear(), today.getMonth(), today.getDate());

		// Calcula diferencias
		var diffDays = Math.floor((fechaTerminoSimple - hoySimple) / (1000 * 60 * 60 * 24));
		if (diffDays === 30 || diffDays === 15 || diffDays === 7) {
			Logger.log(nombre + " - " + universidad + " - Termina en " + diffDays + " días");
			// Formato en español para la fecha
			var fechaFormateada = formatDateEs(fechaTermino, Session.getScriptTimeZone());
			var asunto = "Aviso de término de " + tipoPrograma + ": " + nombre;

			// Usar plantilla HTML
			var template = HtmlService.createTemplateFromFile('AvisoTerminoMensaje');
			template.nombre = nombre;
			template.universidad = universidad;
			template.fechaFormateada = fechaFormateada;
			template.diffDays = diffDays;
			var mensaje = template.evaluate().getContent();

			MailApp.sendEmail({
				to: email,
				subject: asunto,
				htmlBody: mensaje
			});
		}
		// Crear evento en el calendario el día que termina el programa
		var tituloEvento = "Término de " + tipoPrograma + ": " + nombre;
		var eventos = CalendarApp.getCalendarById(calendarId).getEventsForDay(fechaTermino);
		var existeEvento = eventos.some(function(evento) {
			return evento.getTitle() === tituloEvento;
		});
		if (!existeEvento) {
			var fechaInicioFormateada = "";
			if (fechaInicio && Object.prototype.toString.call(fechaInicio) === '[object Date]') {
				fechaInicioFormateada = formatDateEs(fechaInicio, Session.getScriptTimeZone());
			}
			CalendarApp.getCalendarById(calendarId).createAllDayEvent(
				tituloEvento,
				fechaTermino,
				{
					description:
						"Nombre - " + nombre + "\n" +
						"Universidad - " + universidad + "\n" +
						"Licenciatura - " + licenciatura + "\n" +
						"Fecha de inicio - " + (fechaInicioFormateada ? fechaInicioFormateada : "")
				}
			);
		}
	}
}