function doGet() {
    return HtmlService.createTemplateFromFile('web')
        .evaluate()
        .setTitle('Agenda Google Apps Script');
}

function obtenerDatosHTML(nombre) {
   return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

function obtenerContactos() {
 let hoja = SpreadsheetApp.openById('14COXrGrCi0htxdSg2zDa8shcky8F3pltw8tGbNjvySs').getSheetByName('AGENDA');
 let datos = hoja.getDataRange().getValues();
 return datos;
}