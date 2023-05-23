const HOJA = SpreadsheetApp.openById('14COXrGrCi0htxdSg2zDa8shcky8F3pltw8tGbNjvySs').getSheetByName('AGENDA');
const CARPETA = DriveApp.getFolderById('1-uKHP8AJNQaMOf-tMS6X09MgPee2k9q7');
const CABECERA_URL_IMAGEN_DRIVE = 'https://drive.google.com/uc?export=view&id=';

function doGet() {
    return HtmlService.createTemplateFromFile('web')
        .evaluate()
        .setTitle('Agenda Google Apps Script');
}

function obtenerDatosHTML(nombre) {
   return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

function obtenerDatos() {
    return HOJA.getDataRange().getValues();
}

function insertarContacto(contacto, imagen) {
    if(imagen) {
        let blob = Utilities.newBlob(imagen.datos, imagen.tipo, imagen.nombre);
        let archivo = CARPETA.createFile(blob);
        contacto.imagen = CABECERA_URL_IMAGEN_DRIVE + archivo.getId();
    }
    HOJA.appendRow([contacto.nombre, contacto.apellidos, contacto.correo, contacto.telf, contacto.imagen]);
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
    HOJA.appendRow([contacto.name.first, contacto.name.last, contacto.email, contacto.phone, contacto.picture.large]);
}