function avisoTerminoServicioSocial() {
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	var data = sheet.getDataRange().getValues();
	var today = new Date();

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
		var nombre = data[i][0]; // Nombre
		var universidad = data[i][1]; // Universidad
		var tipoPrograma = data[i][7]; // Tipo de programa
		var licenciatura = data[i][6]; // Carrera o licenciatura
		var correoResponsable1 = data[i][10]; // Correo responsable interno 1
		var correoResponsable2 = data[i][11]; // Correo responsable interno 2
		var fechaInicio = data[i][13]; // Fecha de Inicio
		var fechaTermino = data[i][14]; // Fecha de termino
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

			// Enviar correos a colaboradores y los responsables internos
			var destinatarios = ["colaboradores@fundacionleontrece.org"];
			if (correoResponsable1) destinatarios.push(correoResponsable1);
			if (correoResponsable2) destinatarios.push(correoResponsable2);
			
			if (destinatarios.length > 0) {
				MailApp.sendEmail({
					to: destinatarios.join(','),
					subject: asunto,
					htmlBody: mensaje
				});
			}
		}
		
		// Crear evento en los calendarios de colaboradores y los responsables internos
		var tituloEvento = "Término de " + tipoPrograma + ": " + nombre;
		var calendarios = ["colaboradores@fundacionleontrece.org"];
		if (correoResponsable1) calendarios.push(correoResponsable1);
		if (correoResponsable2) calendarios.push(correoResponsable2);
		
		var fechaInicioFormateada = "";
		if (fechaInicio && Object.prototype.toString.call(fechaInicio) === '[object Date]') {
			fechaInicioFormateada = formatDateEs(fechaInicio, Session.getScriptTimeZone());
		}
		
		for (var j = 0; j < calendarios.length; j++) {
			var calendar = CalendarApp.getCalendarById(calendarios[j]);
			if (calendar) {
				var eventos = calendar.getEventsForDay(fechaTermino);
				var existeEvento = eventos.some(function(evento) {
					return evento.getTitle() === tituloEvento;
				});
				if (!existeEvento) {
					calendar.createAllDayEvent(
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
	}
}