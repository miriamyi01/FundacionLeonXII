function archivosVacaciones() {
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

  // Procesar solo socios con Fecha de ingreso no vacía
  let archivosSocios = [];
  for (let s = 0; s < socios.length; s++) {
    const row = socios[s];
    if (!row[headers.indexOf('Fecha de ingreso')]) continue; // <-- Solo si tiene fecha de ingreso
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
      // Si la antigüedad es 0, no crear archivo de vacaciones
      if (Number(antiguedad) === 0) {
        Logger.log('Socio con antigüedad 0, no se crea archivo.');
        continue;
      }
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
      Logger.log(`Procesando socio idx ${s}: ${row[headers.indexOf('Empleado')]} (${codigo})`);
      if (!archivoYaExiste) {
        nuevoArchivo = SpreadsheetApp.create(nombreArchivo);

        folder.addFile(DriveApp.getFileById(nuevoArchivo.getId()));
        Logger.log(`Archivo creado: ${nombreArchivo}`);

        let copiaVacaciones = vacacionesSheet.copyTo(nuevoArchivo);
        copiaVacaciones.setName('Vacaciones');
        let copiaReposicion = reposicionSheet.copyTo(nuevoArchivo);
        copiaReposicion.setName('Reposición de días');

        // Copiar formatos condicionales
        try {
          // Copiar SOLO las reglas que usan rangos de la hoja copiada (no de la hoja original)
          // Esto es necesario porque copyTo crea una hoja con un nuevo ID y las reglas deben apuntar a la hoja destino
          function copiarReglasCondicionales(origen, destino) {
            // Mapear los rangos de la hoja original a la hoja destino
            const reglas = origen.getConditionalFormatRules();
            const reglasFiltradas = [];
            reglas.forEach(r => {
              // Solo reglas cuyos rangos están todos en la hoja original
              const origenRanges = r.getRanges();
              if (origenRanges.every(range => range.getSheet().getName() === origen.getName())) {
                // Crear nuevos rangos en la hoja destino con el mismo A1Notation
                const nuevosRanges = origenRanges.map(range => destino.getRange(range.getA1Notation()));
                // Clonar la regla y asignar los nuevos rangos
                const builder = r.copy();
                builder.setRanges(nuevosRanges);
                reglasFiltradas.push(builder.build());
              }
            });
            destino.setConditionalFormatRules(reglasFiltradas);
          }

          copiarReglasCondicionales(vacacionesSheet, copiaVacaciones);
          copiarReglasCondicionales(reposicionSheet, copiaReposicion);
        } catch (e) {
          Logger.log('Error copiando formatos condicionales: ' + e);
        }

        // Configura el idioma y zona horaria en español/México DESPUÉS de copiar hojas y reglas
        nuevoArchivo.setSpreadsheetLocale('es');
        nuevoArchivo.setSpreadsheetTimeZone('America/Mexico_City');

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
        vacSheet.getRange('C3').setValue(fechaIngresoStr);
        vacSheet.getRange('C4').setValue(antiguedad);
        vacSheet.getRange('J3').setValue(vacacionesAniv);
        vacSheet.getRange('J4').setValue(fechaAnivStr);
        vacSheet.getRange('J5').setValue(fechaVencimientoStr);
        // Poner nombre del líder en G11
        vacSheet.getRange('G11').setValue(lider);

        // Llenar datos en hoja Reposición
        const repSheet = nuevoArchivo.getSheetByName('Reposición de días');
        repSheet.getRange('C1').setDataValidation(null);
        repSheet.getRange('C1').setValue(empleado); // Nombre en C1

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
          vacSheet.getRange('C3').setValue(fechaIngresoStr);
          vacSheet.getRange('C4').setValue(antiguedad);
          vacSheet.getRange('J3').setValue(vacacionesAniv);
          vacSheet.getRange('J4').setValue(fechaAnivStr);
          vacSheet.getRange('J5').setValue(fechaVencimientoStr);
          vacSheet.getRange('G11').setValue(lider);
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

      // SIEMPRE ejecutar permisos y logs, tanto si es nuevo como si ya existe
      const vacSheet = nuevoArchivo.getSheetByName('Vacaciones');
      const repSheet = nuevoArchivo.getSheetByName('Reposición de días');

      // SOLO el día de aniversario comparte y asigna rangos
      const hoy = new Date();
      hoy.setHours(0,0,0,0);
      const aniversario = new Date(fechaAnivDate);
      aniversario.setHours(0,0,0,0);

      if (hoy.getTime() === aniversario.getTime()) {
        // Compartir archivo como editor para que tengan acceso
        if (correoColaborador) {
          nuevoArchivo.addEditor(correoColaborador);
          Logger.log(`Se agregó acceso de editor al archivo para colaborador: ${correoColaborador}`);
        }
        if (correoLider) {
          nuevoArchivo.addEditor(correoLider);
          Logger.log(`Se agregó acceso de editor al archivo para líder: ${correoLider}`);
        }
        
        // Copiar rangos protegidos y editores de la hoja original de Vacaciones
        {
          const proteccionesOriginales = vacacionesSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
          Logger.log(`VACACIONES: Protecciones originales encontradas: ${proteccionesOriginales.length}`);
          vacSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
          proteccionesOriginales.forEach(function(p, idx) {
            const rango = p.getRange();
            const descripcion = p.getDescription();
            Logger.log(`VACACIONES: Copiando protección #${idx+1} - Rango: ${rango.getA1Notation()}, Descripción: "${descripcion}"`);
            const nuevaProteccion = vacSheet.getRange(rango.getA1Notation()).protect();
            nuevaProteccion.setDescription(descripcion);

            if (descripcion === 'AGREGAR CORREO: CADA COLABORADOR VACACIONES') {
              // Originales + colaborador
              nuevaProteccion.removeEditors(nuevaProteccion.getEditors());
              p.getEditors().forEach(function(editor) {
                nuevaProteccion.addEditor(editor);
              });
              if (correoColaborador) {
                nuevaProteccion.addEditor(correoColaborador);
                Logger.log(`VACACIONES: Permiso otorgado: colaborador (${correoColaborador}) en rango ${rango.getA1Notation()}`);
              }
            } else if (descripcion === 'AGREGAR CORREO: LIDER VACACIONES') {
              // Originales + líder
              nuevaProteccion.removeEditors(nuevaProteccion.getEditors());
              p.getEditors().forEach(function(editor) {
                nuevaProteccion.addEditor(editor);
              });
              if (correoLider) {
                nuevaProteccion.addEditor(correoLider);
                Logger.log(`VACACIONES: Permiso otorgado: líder (${correoLider}) en rango ${rango.getA1Notation()}`);
              }
            } else {
              // Solo copia los editores originales para los demás rangos
              nuevaProteccion.removeEditors(nuevaProteccion.getEditors());
              p.getEditors().forEach(function(editor) {
                nuevaProteccion.addEditor(editor);
              });
            }
          });
        }

        // Copiar rangos protegidos y editores de la hoja original de Reposición
        {
          const proteccionesOriginales = reposicionSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
          Logger.log(`REPOSICION: Protecciones originales encontradas: ${proteccionesOriginales.length}`);
          repSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
          proteccionesOriginales.forEach(function(p, idx) {
            const rango = p.getRange();
            const descripcion = p.getDescription();
            Logger.log(`REPOSICION: Copiando protección #${idx+1} - Rango: ${rango.getA1Notation()}, Descripción: "${descripcion}"`);
            const nuevaProteccion = repSheet.getRange(rango.getA1Notation()).protect();
            nuevaProteccion.setDescription(descripcion);

            if (descripcion === 'AGREGAR CORREO: CADA COLABORADOR REPOSICION') {
              // Originales + colaborador
              nuevaProteccion.removeEditors(nuevaProteccion.getEditors());
              p.getEditors().forEach(function(editor) {
                nuevaProteccion.addEditor(editor);
              });
              if (correoColaborador) {
                nuevaProteccion.addEditor(correoColaborador);
                Logger.log(`REPOSICION: Permiso otorgado: colaborador (${correoColaborador}) en rango ${rango.getA1Notation()}`);
              }
            } else if (descripcion === 'AGREGAR CORREO: LIDER REPOSICION') {
              // Originales + líder
              nuevaProteccion.removeEditors(nuevaProteccion.getEditors());
              p.getEditors().forEach(function(editor) {
                nuevaProteccion.addEditor(editor);
              });
              if (correoLider) {
                nuevaProteccion.addEditor(correoLider);
                Logger.log(`REPOSICION: Permiso otorgado: líder (${correoLider}) en rango ${rango.getA1Notation()}`);
              }
            } else {
              // Solo copia los editores originales para los demás rangos
              nuevaProteccion.removeEditors(nuevaProteccion.getEditors());
              p.getEditors().forEach(function(editor) {
                nuevaProteccion.addEditor(editor);
              });
            }
          });
        }
      } else {
        Logger.log('No se otorgaron permisos ni rangos porque hoy no es el día de aniversario');
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

      // Eliminar hojas Aniversarios_{año} donde año < yearObjetivo - 10
      const limite = yearObjetivo - 10;
      ssIndice.getSheets().forEach(sheet => {
        const match = sheet.getName().match(/^Aniversarios_(\d{4})$/);
        if (match) {
          const anio = parseInt(match[1], 10);
          if (anio < limite) {
            ssIndice.deleteSheet(sheet);
            Logger.log(`Hoja eliminada: ${sheet.getName()}`);
          }
        }
      });

      let hojaNueva = ssIndice.getSheetByName(`Aniversarios_${yearObjetivo}`);
      if (!hojaNueva) {
        // Buscar hoja actual (la primera)
        const hojaActual = ssIndice.getSheets()[0];
        hojaNueva = hojaActual.copyTo(ssIndice);
        hojaNueva.setName(`Aniversarios_${yearObjetivo}`);
        // Borrar datos desde fila 5 en adelante
        const lastRow = hojaNueva.getLastRow();
        if (lastRow >= 5) {
          hojaNueva.deleteRows(5, lastRow - 4);
        }
        // Formato especial para columnas
        hojaNueva.getRange(1, 1, 1, 10).setHorizontalAlignment('center');
        hojaNueva.getRange(2, 1, hojaNueva.getMaxRows() - 1, 4).setFontWeight('bold');
        hojaNueva.getRange(2, 1, hojaNueva.getMaxRows() - 1, 5).setHorizontalAlignment('center');
        hojaNueva.getRange(2, 6, hojaNueva.getMaxRows() - 1, 5).setHorizontalAlignment('left');
        hojaNueva.getRange(2, 1, hojaNueva.getMaxRows() - 1, 10).setBorder(true, true, true, true, true, true, '#002147', SpreadsheetApp.BorderStyle.SOLID);
      }

      // Ordenar archivosSocios por zona y por ID
      archivosSocios.sort(function(a, b) {
        const zonaA = parseInt(a.codigoZona.replace('Z', ''), 10);
        const zonaB = parseInt(b.codigoZona.replace('Z', ''), 10);
        if (zonaA !== zonaB) {
          return zonaA - zonaB;
        }
        return parseInt(a.idColaborador, 10) - parseInt(b.idColaborador, 10);
      });



      // Asegura que la hoja tenga al menos 5 filas antes de buscar o escribir
      if (hojaNueva.getMaxRows() < 5) {
        hojaNueva.insertRowsAfter(hojaNueva.getMaxRows(), 5 - hojaNueva.getMaxRows());
      }

      // Obtener todos los IDs de colaboradores ya presentes en la hoja (solo si hay filas suficientes)
      let idsExistentes = [];
      const lastRowAniv = hojaNueva.getLastRow();
      if (lastRowAniv >= 5) {
        idsExistentes = hojaNueva.getRange(5, 4, lastRowAniv - 4, 1).getValues()
          .map(val => val[0] ? val[0].toString() : "");
      }

      let nuevosAgregados = 0;
      archivosSocios.forEach((socio) => {
        // Si el colaborador ya está, no lo agregues
        if (idsExistentes.includes(socio.idColaborador)) return;

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

        // Busca la primera fila vacía a partir de la fila 5
        let filaDestino = 5;
        const maxRows = hojaNueva.getMaxRows();
        // Usa lastRowAniv en vez de lastRow
        let foundEmpty = false;
        for (let i = 5; i <= hojaNueva.getLastRow(); i++) {
          const celda = hojaNueva.getRange(i, 4).getValue();
          if (!celda) {
            filaDestino = i;
            foundEmpty = true;
            break;
          }
        }
        if (!foundEmpty) {
          filaDestino = hojaNueva.getLastRow() + 1;
        }

        // Asegura que hay suficientes filas (si la hoja está vacía o se requiere agregar)
        if (maxRows < filaDestino) {
          hojaNueva.insertRowsAfter(maxRows, filaDestino - maxRows);
        }

        // Solo escribe si la fila destino es válida (mayor o igual a 5)
        if (filaDestino >= 5) {
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
          hojaNueva.getRange(filaDestino, 1, 1, 10).setFontFamily('Lato');
          hojaNueva.getRange(filaDestino, 1, 1, 10).setFontSize(12);
          hojaNueva.getRange(filaDestino, 1, 1, 10).setBackground('#FFFFFF');
          // ACTUALIZA idsExistentes para evitar sobrescribir en la siguiente iteración
          idsExistentes.push(socio.idColaborador);
        }
        nuevosAgregados++;
      });

      if (nuevosAgregados > 0) {
        Logger.log(`Hoja índice actualizada. Colaboradores agregados: ${nuevosAgregados}`);
      }
    } else {
      Logger.log('No se encontró archivo índice para duplicar');
    }
  } catch (e) {
    Logger.log('Error actualizando hoja índice: ' + e);
  }
}