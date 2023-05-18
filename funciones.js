const HOJA = SpreadsheetApp.openById('14COXrGrCi0htxdSg2zDa8shcky8F3pltw8tGbNjvySs').getSheetByName('AGENDA');

function doGet() {
    return HtmlService.createTemplateFromFile('web')
        .evaluate()
        .setTitle('Agenda Google Apps Script');
}

function obtenerDatosHTML(nombre) {
   return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

function obtenerContactos() {
    return HOJA.getDataRange().getValues();
}

function insertarContacto(nombre, apellidos, correo, telf) {
    HOJA.appendRow([nombre, apellidos, correo, telf]);
}

function borrarContacto(numFila) {
    HOJA.deleteRow(numFila);
}

function modificarContacto(numFila, datos) {
    let celdas = HOJA.getRange('A' + numFila + ':D' + numFila);
    celdas.setValues([[datos.nombre, datos.apellidos, datos.correo, datos.telf]]);
}