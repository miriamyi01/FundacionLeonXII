function permisosDiarios() {
    // Obtener la hoja de socios (igual que en archivo 2)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const baseSheet = ss.getSheetByName('Base Vacaciones');
    const vacacionesSheet = ss.getSheetByName('Vacaciones');
    const reposicionSheet = ss.getSheetByName('Reposición de días');
    const permisosSheet = ss.getSheetByName('Permisos');
    const yearObjetivo = new Date().getFullYear();
    const zonaHoraria = 'America/Mexico_City';

    // Carpeta del año actual
    const drive = DriveApp;
    const fileActual = drive.getFileById(ss.getId());
    const parentFolders = fileActual.getParents();
    let parentFolder;
    if (parentFolders.hasNext()) {
        parentFolder = parentFolders.next();
    } else {
        Logger.log('No se encontró la carpeta contenedora del archivo actual.');
        return;
    }
    const folderName = yearObjetivo.toString();
    let folder;
    const folders = parentFolder.getFoldersByName(folderName);
    if (folders.hasNext()) {
        folder = folders.next();
    } else {
        Logger.log('No se encontró la carpeta del año actual: ' + folderName);
        return;
    }

    // Leer base de socios igual que en archivo 2
    const data = baseSheet.getDataRange().getValues();
    const headers = data[0];
    const socios = data.slice(1).filter(row => row[headers.indexOf('Fecha de ingreso')]);
    Logger.log('SOCIOS ENCONTRADOS: ' + socios.length);

    // Utilidad para traducir meses a español
    function traducirFecha(fechaStr) {
        const meses = {
            'January': 'enero', 'February': 'febrero', 'March': 'marzo', 'April': 'abril',
            'May': 'mayo', 'June': 'junio', 'July': 'julio', 'August': 'agosto',
            'September': 'septiembre', 'October': 'octubre', 'November': 'noviembre', 'December': 'diciembre'
        };
        let res = fechaStr;
        Object.keys(meses).forEach(m => {
            res = res.replace(m, meses[m]);
        });
        return res;
    }

    // Convierte "dd/MM/yy" a Date
    function parseFechaDMY(fechaStr) {
        if (fechaStr instanceof Date) return fechaStr;
        if (typeof fechaStr !== 'string') return new Date(fechaStr);
        const partes = fechaStr.split('/');
        if (partes.length !== 3) return new Date(fechaStr);
        let d = parseInt(partes[0], 10);
        let m = parseInt(partes[1], 10) - 1;
        let y = parseInt(partes[2], 10);
        if (y < 100) y += 2000;
        return new Date(y, m, d);
    }

    // Verifica si el colaborador/líder tiene permisos de editor en el archivo
    function faltaEditor(archivo, correo) {
        if (!correo) return false;
        return !archivo.getEditors().map(e => e.getEmail()).includes(correo);
    }

    // Verifica si el colaborador/líder está en los rangos protegidos de la hoja
    function faltaProteccion(sheet, correo, descripcion) {
        if (!correo) return false;
        const protecciones = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        // Si no existe el rango con esa descripción, falta la protección
        const proteccion = protecciones.find(p => p.getDescription() === descripcion);
        if (!proteccion) return true;
        // Si existe pero el correo no está, también falta
        return !proteccion.getEditors().map(e => e.getEmail()).includes(correo);
    }

    // Copiar protección desde origen si falta en destino
    function copiarProteccionSiFalta(origenSheet, destinoSheet, correo, descripcion) {
        if (!correo) return;
        const proteccionesDestino = destinoSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        let proteccionDestino = proteccionesDestino.find(p => p.getDescription() === descripcion);

        if (!proteccionDestino) {
            // Buscar en origen y copiar
            const proteccionOrigen = origenSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
                .find(p => p.getDescription() === descripcion);
            if (proteccionOrigen) {
                const rango = proteccionOrigen.getRange();
                proteccionDestino = destinoSheet.getRange(rango.getA1Notation()).protect();
                proteccionDestino.setDescription(descripcion);
                proteccionDestino.removeEditors(proteccionDestino.getEditors());
                proteccionOrigen.getEditors().forEach(editor => proteccionDestino.addEditor(editor));
            }
        }
        // Si ya existe o se acaba de crear, agregar el correo si falta
        if (proteccionDestino && !proteccionDestino.getEditors().map(e => e.getEmail()).includes(correo)) {
            proteccionDestino.addEditor(correo);
            Logger.log(`Permiso otorgado: ${correo} en rango ${proteccionDestino.getRange().getA1Notation()} (${descripcion})`);
        }
    }

    // Copiar todas las protecciones desde origen si falta alguna en destino
    function copiarTodasProtecciones(origenSheet, destinoSheet) {
        const proteccionesOrigen = origenSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        const proteccionesDestino = destinoSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        const descripcionesDestino = proteccionesDestino.map(p => p.getDescription());

        proteccionesOrigen.forEach(function(p) {
            // Si no existe una protección con la misma descripción en destino, la copiamos
            if (!descripcionesDestino.includes(p.getDescription())) {
                const rango = p.getRange();
                const nuevaProteccion = destinoSheet.getRange(rango.getA1Notation()).protect();
                nuevaProteccion.setDescription(p.getDescription());
                nuevaProteccion.removeEditors(nuevaProteccion.getEditors());
                p.getEditors().forEach(editor => nuevaProteccion.addEditor(editor));
                // Copiar también propietarios si aplica
                if (p.canDomainEdit() !== undefined) {
                    nuevaProteccion.setDomainEdit(p.canDomainEdit());
                }
                Logger.log(`Permiso otorgado: ${p.getEditors().map(e => e.getEmail()).join(', ')} en rango ${nuevaProteccion.getRange().getA1Notation()} (${p.getDescription()})`);
            }
        });
    }

    // Cambia la lógica: si falta cualquier rango protegido en la hoja destino, también debe otorgar permisos
    function faltanProtecciones(origenSheet, destinoSheet) {
        const proteccionesOrigen = origenSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        const descripcionesDestino = destinoSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).map(p => p.getDescription());
        // Si alguna protección de origen no existe en destino, retorna true
        return proteccionesOrigen.some(p => !descripcionesDestino.includes(p.getDescription()));
    }

    for (let i = 0; i < socios.length; i++) {
        const row = socios[i];
        // Solo procesar si hay fecha de ingreso y antigüedad > 0
        const fechaIngreso = row[headers.indexOf('Fecha de ingreso')];
        const antiguedad = row[headers.indexOf(`Años de antigüedad a ${yearObjetivo}`)];
        if (!fechaIngreso || Number(antiguedad) === 0) {
            Logger.log(`Socio omitido por falta de fecha de ingreso o antigüedad 0`);
            continue;
        }

        // Código de socio a 3 dígitos
        const codigo = row[headers.indexOf('Código')].toString().padStart(3, '0');
        const oficinaRaw = row[headers.indexOf('Oficina')].toString();
        const oficinaDosDigitos = oficinaRaw.substring(0,2).padStart(2, '0');
        const codigoZona = `Z${oficinaDosDigitos}`;
        const iniciales = row[headers.indexOf('Iniciales colaborador')];
        const empleado = row[headers.indexOf('Empleado')];
        const fechaAniv = row[headers.indexOf(`Fecha aniversario ${yearObjetivo}`)];
        const correoColaborador = row[headers.indexOf('Correo')];
        const correoLider = row[headers.indexOf('Correo Líder')];

        // Formatear fecha aniversario en español y formato corto
        const fechaAnivDate = parseFechaDMY(fechaAniv);
        const fechaAnivStr = traducirFecha(Utilities.formatDate(fechaAnivDate, zonaHoraria, "dd/MM/yy"));

        // Formatear vencimientoNombre igual que en archivo 2
        const fechaVencimientoDate = new Date(fechaAnivDate);
        fechaVencimientoDate.setFullYear(fechaVencimientoDate.getFullYear() + 1);
        fechaVencimientoDate.setDate(fechaVencimientoDate.getDate() - 1);
        const vencimientoNombre = traducirFecha(
            Utilities.formatDate(fechaVencimientoDate, zonaHoraria, "d MMM yyyy")
        ).toUpperCase();

        // Nombre archivo igual que en archivo 2
        const nombreArchivo = `${codigo} - ${codigoZona} VACACIONES ${iniciales} VTO. ${vencimientoNombre}`;

        // Buscar archivo solo en la carpeta del año actual
        let files = folder.getFilesByName(nombreArchivo);
        if (!files.hasNext()) continue;
        const archivoEncontrado = files.next();

        const nuevoArchivo = SpreadsheetApp.openById(archivoEncontrado.getId());
        const vacSheet = nuevoArchivo.getSheetByName('Vacaciones');
        const repSheet = nuevoArchivo.getSheetByName('Reposición de días');
        const permSheet = nuevoArchivo.getSheetByName('Permisos');

        // Otorgar permisos si falta algún permiso de editor, protección de colaborador/líder, o falta algún rango protegido en general
        const faltaProteccionGeneral =
            faltanProtecciones(vacacionesSheet, vacSheet) ||
            faltanProtecciones(reposicionSheet, repSheet) ||
            faltanProtecciones(permisosSheet, permSheet);

        const faltaPermisoColaborador =
            faltaEditor(nuevoArchivo, correoColaborador) ||
            faltaProteccion(vacSheet, correoColaborador, 'AGREGAR CORREO: CADA COLABORADOR VACACIONES') ||
            faltaProteccion(repSheet, correoColaborador, 'AGREGAR CORREO: CADA COLABORADOR REPOSICION') ||
            faltaProteccion(permSheet, correoColaborador, 'AGREGAR CORREO: CADA COLABORADOR PERMISOS');

        const faltaPermisoLider =
            faltaEditor(nuevoArchivo, correoLider) ||
            faltaProteccion(vacSheet, correoLider, 'AGREGAR CORREO: LIDER VACACIONES') ||
            faltaProteccion(repSheet, correoLider, 'AGREGAR CORREO: LIDER REPOSICION') ||
            faltaProteccion(permSheet, correoLider, 'AGREGAR CORREO: LIDER PERMISOS');

        let otorgarPermisos = false;
        if (fechaAniv) {
            const hoy = new Date();
            hoy.setHours(0,0,0,0);
            const fechaAnivDateCmp = new Date(fechaAnivDate.getTime());
            fechaAnivDateCmp.setHours(0,0,0,0);
            if (hoy >= fechaAnivDateCmp && (faltaPermisoColaborador || faltaPermisoLider || faltaProteccionGeneral)) {
                otorgarPermisos = true;
            }
        }

        if (otorgarPermisos) {
            Logger.log(`[PERMISOS] Socio: ${empleado} | Aniversario: ${fechaAnivStr} | Archivo: ${nombreArchivo}`);

            // Compartir archivo como editor solo si no lo tiene
            [correoColaborador, correoLider].forEach(correo => {
                if (faltaEditor(nuevoArchivo, correo)) {
                    nuevoArchivo.addEditor(correo);
                    Logger.log(`Se agregó acceso de editor al archivo para: ${correo}`);
                }
            });

            // Copiar todas las protecciones de cada hoja si falta alguna
            copiarTodasProtecciones(vacacionesSheet, vacSheet);
            copiarTodasProtecciones(reposicionSheet, repSheet);
            copiarTodasProtecciones(permisosSheet, permSheet);

            // Asegurar que colaborador/líder estén en los rangos que les corresponden
            function asegurarEditorEnProteccion(sheet, correo, descripcion) {
                if (!correo) return;
                const proteccion = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
                    .find(p => p.getDescription() === descripcion);
                if (proteccion && !proteccion.getEditors().map(e => e.getEmail()).includes(correo)) {
                    proteccion.addEditor(correo);
                    Logger.log(`Permiso otorgado: ${correo} en rango ${proteccion.getRange().getA1Notation()} (${descripcion})`);
                }
            }

            // Colaborador
            asegurarEditorEnProteccion(vacSheet, correoColaborador, 'AGREGAR CORREO: CADA COLABORADOR VACACIONES');
            asegurarEditorEnProteccion(repSheet, correoColaborador, 'AGREGAR CORREO: CADA COLABORADOR REPOSICION');
            asegurarEditorEnProteccion(permSheet, correoColaborador, 'AGREGAR CORREO: CADA COLABORADOR PERMISOS');
            // Líder
            asegurarEditorEnProteccion(vacSheet, correoLider, 'AGREGAR CORREO: LIDER VACACIONES');
            asegurarEditorEnProteccion(repSheet, correoLider, 'AGREGAR CORREO: LIDER REPOSICION');
            asegurarEditorEnProteccion(permSheet, correoLider, 'AGREGAR CORREO: LIDER PERMISOS');
        } else {
            if (fechaAniv) {
                const hoy = new Date();
                hoy.setHours(0,0,0,0);
                const fechaAnivDateCmp = new Date(fechaAnivDate.getTime());
                fechaAnivDateCmp.setHours(0,0,0,0);
                if (hoy < fechaAnivDateCmp) {
                    Logger.log(`NO SE OTORGAN PERMISOS: Aún no es su aniversario - ${empleado} - ${fechaAnivStr}`);
                } else {
                    Logger.log(`NO SE OTORGAN PERMISOS: Ya tiene todos los permisos - ${empleado} - ${fechaAnivStr}`);
                }
            } else {
                Logger.log(`NO SE OTORGAN PERMISOS: Sin fecha de aniversario válida - ${empleado}`);
            }
        }
    }
}