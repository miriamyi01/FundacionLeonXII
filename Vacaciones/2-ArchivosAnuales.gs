function crearArchivosVacacionesTodoEnUno() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName('Base Vacaciones');
  const vacacionesSheet = ss.getSheetByName('Vacaciones');
  const reposicionSheet = ss.getSheetByName('Reposición de días');
  const indiceFileName = '001 - Índice: Archivo de vacaciones';
  const drive = DriveApp;
  const yearObjetivo = new Date().getFullYear() + 1;
  const folderName = yearObjetivo.toString();

  // Declarar zonaHoraria antes de cualquier uso
  const zonaHoraria = 'America/Mexico_City';

  // Obtener la carpeta donde está el archivo actual
  const fileActual = drive.getFileById(ss.getId());
  const parentFolders = fileActual.getParents();
  let parentFolder;
  if (parentFolders.hasNext()) {
    parentFolder = parentFolders.next();
  } else {
    throw new Error('No se encontró la carpeta contenedora del archivo actual.');
  }

  // Crear o usar la subcarpeta del año siguiente dentro de la carpeta actual
  let folder;
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    folder = folders.next();
    Logger.log('Carpeta ya existe: ' + folderName);
  } else {
    folder = parentFolder.createFolder(folderName);
    Logger.log('Carpeta creada: ' + folderName);
  }

  // Leer base de socios
  const data = baseSheet.getDataRange().getValues();
  const headers = data[0];
  const socios = data.slice(1).filter(row => row[headers.indexOf('Fecha de ingreso')]);
  Logger.log('Socios encontrados: ' + socios.length);

  // Procesar todos los socios
  let archivosSocios = [];
  for (let s = 0; s < socios.length; s++) {
    const row = socios[s];
    try {
      // Código de socio a 3 dígitos
      const codigo = row[headers.indexOf('Código')].toString().padStart(3, '0');
      // Obtén los dos dígitos y usa Z{dos dígitos} en ambos lugares
      const oficinaRaw = row[headers.indexOf('Oficina')].toString();
      const oficinaDosDigitos = oficinaRaw.substring(0,2).padStart(2, '0');
      const codigoZona = `Z${oficinaDosDigitos}`;
      const iniciales = row[headers.indexOf('Iniciales colaborador')];
      const fechaIngreso = row[headers.indexOf('Fecha de ingreso')];
      // Agrega esta línea para la fecha en español:
      const fechaIngresoStr = traducirFecha(
        Utilities.formatDate(new Date(fechaIngreso), zonaHoraria, "d 'de' MMMM 'de' yyyy")
      );
      const empleado = row[headers.indexOf('Empleado')];
      const antiguedad = row[headers.indexOf(`Años de antigüedad a ${yearObjetivo}`)];
      const vacacionesAniv = row[headers.indexOf(`Vacaciones a partir de aniversario ${yearObjetivo}`)];
      const fechaAniv = row[headers.indexOf(`Fecha aniversario ${yearObjetivo}`)];
      const lider = row[headers.indexOf('Líder')];
      const inicialesLider = row[headers.indexOf('Iniciales Líder')];
      const correoColaborador = row[headers.indexOf('Correo')];
      const correoLider = row[headers.indexOf('Correo Líder')];

      // Utilidad para mapear días y meses a español
      function traducirFecha(fechaStr) {
        const dias = {
          'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles', 'Thursday': 'Jueves',
          'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
        };
        const meses = {
          'January': 'enero', 'February': 'febrero', 'March': 'marzo', 'April': 'abril',
          'May': 'mayo', 'June': 'junio', 'July': 'julio', 'August': 'agosto',
          'September': 'septiembre', 'October': 'octubre', 'November': 'noviembre', 'December': 'diciembre'
        };
        let res = fechaStr;
        Object.keys(dias).forEach(d => {
          res = res.replace(d, dias[d]);
        });
        Object.keys(meses).forEach(m => {
          res = res.replace(m, meses[m]);
        });
        return res;
      }

      // Formatear fecha aniversario y vencimiento en español (sin parámetro 'es')
      const fechaAnivDate = new Date(fechaAniv);
      const fechaAnivStr = traducirFecha(Utilities.formatDate(fechaAnivDate, zonaHoraria, "EEEE, d 'de' MMMM 'de' yyyy"));
      const fechaVencimientoDate = new Date(fechaAnivDate);
      fechaVencimientoDate.setFullYear(fechaVencimientoDate.getFullYear() + 1);
      fechaVencimientoDate.setDate(fechaVencimientoDate.getDate() - 1);
      const fechaVencimientoStr = traducirFecha(Utilities.formatDate(fechaVencimientoDate, zonaHoraria, "EEEE, d 'de' MMMM 'de' yyyy"));
      const vencimientoNombre = traducirFecha(Utilities.formatDate(fechaVencimientoDate, zonaHoraria, "d MMM yyyy")).toUpperCase();

      // Nombre archivo con Z{dos dígitos}
      const nombreArchivo = `${codigo} - ${codigoZona} VACACIONES ${iniciales} VTO. ${vencimientoNombre}`;

      // Verifica si el archivo ya existe en la carpeta del año objetivo
      let archivoYaExiste = false;
      let archivoIdExistente = null;
      const archivosEnCarpeta = folder.getFilesByName(nombreArchivo);
      if (archivosEnCarpeta.hasNext()) {
        archivoYaExiste = true;
        archivoIdExistente = archivosEnCarpeta.next().getId();
      }

      let nuevoArchivo;
      if (!archivoYaExiste) {
        nuevoArchivo = SpreadsheetApp.create(nombreArchivo);
        folder.addFile(DriveApp.getFileById(nuevoArchivo.getId()));
        Logger.log(`Archivo creado: ${nombreArchivo}`);

        let copiaVacaciones = vacacionesSheet.copyTo(nuevoArchivo);
        copiaVacaciones.setName('Vacaciones');
        let copiaReposicion = reposicionSheet.copyTo(nuevoArchivo);
        copiaReposicion.setName('Reposición de días');

        // Eliminar hoja por defecto (solo si hay más de una hoja)
        const hojas = nuevoArchivo.getSheets();
        if (hojas.length > 1) {
          hojas.forEach(function(hoja) {
            const nombre = hoja.getName();
            if (nombre !== 'Vacaciones' && nombre !== 'Reposición de días') {
              nuevoArchivo.deleteSheet(hoja);
            }
          });
        }

        // Llenar datos en hoja Vacaciones
        const vacSheet = nuevoArchivo.getSheetByName('Vacaciones');
        vacSheet.getRange('C1').setDataValidation(null);
        vacSheet.getRange('C1').setValue(empleado); // Nombre en C1
        vacSheet.getRange('I1').setValue(codigo);
        vacSheet.getRange('K1').setValue(iniciales);
        vacSheet.getRange('C3').setValue(fechaIngresoStr); // <-- formato solo "9 de enero de 2025"
        vacSheet.getRange('C4').setValue(antiguedad);
        vacSheet.getRange('J3').setValue(vacacionesAniv);
        vacSheet.getRange('J4').setValue(fechaAnivStr);
        vacSheet.getRange('J5').setValue(fechaVencimientoStr);

        // Llenar datos en hoja Reposición
        const repSheet = nuevoArchivo.getSheetByName('Reposición de días');
        repSheet.getRange('C1').setDataValidation(null);
        repSheet.getRange('C1').setValue(empleado); // Nombre en C1

        // Copiar rangos protegidos y permisos VACACIONES
        {
          const protecciones = vacacionesSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
          protecciones.forEach(p => {
            const rango = p.getRange();
            const nuevaProteccion = vacSheet.getRange(rango.getA1Notation()).protect();
            nuevaProteccion.setDescription(p.getDescription());
            // Permiso al colaborador para su propio archivo
            if (p.getDescription().indexOf('AGREGAR CORREO: CADA COLABORADOR VACACIONES') !== -1) {
              nuevaProteccion.addEditor(correoColaborador);
            }
            if (p.getDescription().indexOf('AGREGAR CORREO: LIDER VACACIONES') !== -1) {
              nuevaProteccion.addEditor(correoLider);
            }
          });
        }
        // Copiar rangos protegidos y permisos REPOSICION
        {
          const protecciones = reposicionSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
          protecciones.forEach(p => {
            const rango = p.getRange();
            const nuevaProteccion = repSheet.getRange(rango.getA1Notation()).protect();
            nuevaProteccion.setDescription(p.getDescription());
            // Permiso al colaborador para su propio archivo
            if (p.getDescription().indexOf('AGREGAR CORREO: CADA COLABORADOR REPOSICION') !== -1) {
              nuevaProteccion.addEditor(correoColaborador);
            }
            if (p.getDescription().indexOf('AGREGAR CORREO: LIDER REPOSICION') !== -1) {
              nuevaProteccion.addEditor(correoLider);
            }
          });
        }

        archivosSocios.push({
          codigoZona: codigoZona,
          fechaIngreso: fechaIngresoStr,
          // Solo el nombre de la oficina, sin los dos dígitos
          nombreOficina: row[headers.indexOf('Oficina')].toString().replace(/^\d{2}\s*/, ''),
          idColaborador: codigo, // ya es 3 dígitos
          siglasColaborador: iniciales,
          colaborador: empleado,
          correo: correoColaborador,
          jefeDirecto: lider,
          correoJefeDirecto: correoLider,
          url: nuevoArchivo.getUrl()
        });
      } else {
        Logger.log(`Archivo ya existe: ${nombreArchivo}`);
        nuevoArchivo = SpreadsheetApp.openById(archivoIdExistente);

        // Opcional: solo llena datos si están vacíos
        const vacSheet = nuevoArchivo.getSheetByName('Vacaciones');
        if (vacSheet.getRange('C1').getValue() == "") {
          vacSheet.getRange('C1').setDataValidation(null);
          vacSheet.getRange('C1').setValue(empleado); // Nombre en C1
          vacSheet.getRange('I1').setValue(codigo);
          vacSheet.getRange('K1').setValue(iniciales);
          vacSheet.getRange('C3').setValue(fechaIngresoStr); // <-- formato solo "9 de enero de 2025"
          vacSheet.getRange('C4').setValue(antiguedad);
          vacSheet.getRange('J3').setValue(vacacionesAniv);
          vacSheet.getRange('J4').setValue(fechaAnivStr);
          vacSheet.getRange('J5').setValue(fechaVencimientoStr);
        }
        const repSheet = nuevoArchivo.getSheetByName('Reposición de días');
        if (repSheet.getRange('C1').getValue() == "") {
          repSheet.getRange('C1').setDataValidation(null);
          repSheet.getRange('C1').setValue(empleado); // Nombre en C1
        }

        archivosSocios.push({
          codigoZona: codigoZona,
          fechaIngreso: fechaIngresoStr,
          // Solo el nombre de la oficina, sin los dos dígitos
          nombreOficina: row[headers.indexOf('Oficina')].toString().replace(/^\d{2}\s*/, ''),
          idColaborador: codigo, // ya es 3 dígitos
          siglasColaborador: iniciales,
          colaborador: empleado,
          correo: correoColaborador,
          jefeDirecto: lider,
          correoJefeDirecto: correoLider,
          url: nuevoArchivo.getUrl()
        });
      }
    } catch (e) {
      Logger.log('Error creando archivo socio idx ' + s + ': ' + e);
    }
  }

  // Duplicar hoja índice dentro del mismo archivo, borrar datos desde fila 5, renombrar y llenar datos
  try {
    // Buscar archivo índice en carpeta padre
    const filesPadre = parentFolder.getFilesByName(indiceFileName);
    if (filesPadre.hasNext()) {
      const indiceFile = filesPadre.next();
      const ssIndice = SpreadsheetApp.openById(indiceFile.getId());
      // Buscar hoja actual (la primera)
      const hojaActual = ssIndice.getSheets()[0];
      // Duplicar hoja
      const hojaNueva = hojaActual.copyTo(ssIndice);
      hojaNueva.setName(`Aniversarios_${yearObjetivo}`);
      // Borrar datos desde fila 5 en adelante
      const lastRow = hojaNueva.getLastRow();
      if (lastRow >= 5) {
        hojaNueva.deleteRows(5, lastRow - 4);
      }

      // Ordenar archivosSocios por zona y por ID
      archivosSocios.sort(function(a, b) {
        // Extraer número de zona (Z02 -> 2, Z13 -> 13)
        const zonaA = parseInt(a.codigoZona.replace('Z', ''), 10);
        const zonaB = parseInt(b.codigoZona.replace('Z', ''), 10);
        if (zonaA !== zonaB) {
          return zonaA - zonaB;
        }
        // Si zona igual, ordenar por idColaborador (número)
        return parseInt(a.idColaborador, 10) - parseInt(b.idColaborador, 10);
      });

      // Insertar datos de colaboradores faltantes
      archivosSocios.forEach((socio) => {
        // Formatea la fecha de ingreso en dd/MM/yyyy para el condensado
        let fechaIngresoCondensado = "";
        if (socio.fechaIngreso) {
          const idx = data.findIndex(r => r[headers.indexOf('Código')].toString().padStart(3, '0') === socio.idColaborador);
          if (idx > 0) {
            const fechaRaw = data[idx][headers.indexOf('Fecha de ingreso')];
            if (fechaRaw) {
              const fechaObj = new Date(fechaRaw);
              fechaIngresoCondensado = Utilities.formatDate(fechaObj, zonaHoraria, "dd/MM/yyyy");
            }
          }
        }
        // Calcula la siguiente fila vacía en el índice
        let filaDestino = hojaNueva.getLastRow() + 1;
        if (filaDestino < 5) filaDestino = 5;
        // Asegura que hay suficientes filas
        if (hojaNueva.getMaxRows() < filaDestino) {
          hojaNueva.insertRowsAfter(hojaNueva.getMaxRows(), filaDestino - hojaNueva.getMaxRows());
        }
        hojaNueva.getRange(filaDestino, 1, 1, 9).setValues([
          [
            socio.codigoZona,
            fechaIngresoCondensado,
            socio.nombreOficina,
            socio.idColaborador,
            socio.siglasColaborador,
            socio.colaborador,
            socio.correo,
            socio.jefeDirecto,
            socio.correoJefeDirecto,
          ]
        ]);
        hojaNueva.getRange(filaDestino, 10).setValue(socio.url);
        // Formato: Lato 12, fondo blanco
        hojaNueva.getRange(filaDestino, 1, 1, 10).setFontFamily('Lato');
        hojaNueva.getRange(filaDestino, 1, 1, 10).setFontSize(12);
        hojaNueva.getRange(filaDestino, 1, 1, 10).setBackground('#FFFFFF');
      });
      // Formato especial para columnas
      // Encabezados (primera fila): centrados
      hojaNueva.getRange(1, 1, 1, 10).setHorizontalAlignment('center');
      // Primeras 4 columnas en negritas (todas las filas menos encabezado)
      hojaNueva.getRange(2, 1, hojaNueva.getMaxRows() - 1, 4).setFontWeight('bold');
      // Todas las filas centradas en las primeras 5 columnas (menos encabezado)
      hojaNueva.getRange(2, 1, hojaNueva.getMaxRows() - 1, 5).setHorizontalAlignment('center');
      // Resto de columnas (6 a 10) alineadas a la izquierda (menos encabezado)
      hojaNueva.getRange(2, 6, hojaNueva.getMaxRows() - 1, 5).setHorizontalAlignment('left');
      // Bordes azul oscuro en todas las celdas menos encabezado
      hojaNueva.getRange(2, 1, hojaNueva.getMaxRows() - 1, 10).setBorder(true, true, true, true, true, true, '#002147', SpreadsheetApp.BorderStyle.SOLID);

      Logger.log('Hoja índice duplicada, renombrada y llenada con colaboradores faltantes');
    } else {
      Logger.log('No se encontró archivo índice para duplicar');
    }
  } catch (e) {
    Logger.log('Error duplicando hoja índice: ' + e);
  }
}