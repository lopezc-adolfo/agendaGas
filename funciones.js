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

function importarContactos() {
    let url = 'https://randomuser.me/api/?results=5&inc=name,email,phone, picture';
    let respuesta = UrlFetchApp.fetch(url).getContentText();
    let datos = JSON.parse(respuesta);

    datos.results.forEach(insertarContactoJSON);
}

function insertarContactoJSON(contacto) {
    HOJA.appendRow([contacto.name.first, contacto.name.last, contacto.email, contacto.phone, contacto.picture.medium]);
}