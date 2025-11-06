function ActualizaciónBase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Base Vacaciones");
  if (!sheet) return;
  if (ss.getActiveSheet().getName() !== "Base Vacaciones") return;

  var lastCol = sheet.getLastColumn();
  var anioSiguiente = new Date().getFullYear() + 1;
  var encabezados = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var encabezadoBuscado = 'Años de antigüedad a ' + anioSiguiente;
  if (encabezados.includes(encabezadoBuscado)) {
    Logger.log("Ya existen datos y fórmulas para el " + anioSiguiente + ".");
    return;
  }

  sheet.insertColumnsAfter(lastCol, 3);

  // Color aleatorio único para los nuevos encabezados
  function colorAleatorioUnico(existentes) {
    var letras = '0123456789ABCDEF', color;
    do {
      color = '#' + Array.from({length:6},()=>letras[Math.floor(Math.random()*16)]).join('');
    } while (existentes.includes(color));
    return color;
  }
  var coloresExistentes = sheet.getRange(1, 1, 1, lastCol).getBackgrounds()[0];
  var colorUnico = colorAleatorioUnico(coloresExistentes);

  // Encabezados y formato
  var nuevosEncabezados = [
    'Años de antigüedad a ' + anioSiguiente,
    'Vacaciones a partir de aniversario ' + anioSiguiente,
    'Fecha aniversario ' + anioSiguiente
  ];
  sheet.getRange(1, lastCol + 1, 1, 3)
    .setValues([nuevosEncabezados])
    .setBackground(colorUnico);

  var lastRow = sheet.getLastRow();
  var rangoNuevasColumnas = sheet.getRange(1, lastCol + 1, lastRow, 3);
  rangoNuevasColumnas.setBorder(true, true, true, true, false, false, "#888888", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setHorizontalAlignment("center");

  // Fórmulas para filas de datos
  if (lastRow > 1) {
    var formulas = [];
    for (var i = 2; i <= lastRow; i++) {
      var celdaAntiguedad = sheet.getRange(i, lastCol + 1).getA1Notation();
      formulas.push([
        `=IF(AND(ISNUMBER($J${i}), ${anioSiguiente}-YEAR($J${i})>0), ${anioSiguiente}-YEAR($J${i}), "")`,
        `=IFERROR(LOOKUP(${celdaAntiguedad}, {1,2,3,4,10,15,20,25}, {8,10,12,14,16,18,20,22}), "")`,
        `=IF(ISNUMBER($J${i}), DATE(${anioSiguiente}, MONTH($J${i}), DAY($J${i})), "")`
      ]);
    }
    sheet.getRange(2, lastCol + 1, lastRow - 1, 3).setFormulas(formulas);
    sheet.getRange(2, lastCol + 1, lastRow - 1, 2).setNumberFormat('0');
  }
  Logger.log("Fórmulas y formato aplicados correctamente en las nuevas columnas para el " + anioSiguiente + ".");
}