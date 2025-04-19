function onOpen() {
  DocumentApp.getUi()
    .createMenu('Carátula')
    .addItem('Generar', 'showModal')
    .addToUi();
}

function showModal(){
  const html = HtmlService
    .createHtmlOutputFromFile('dialogFinal')
    .setWidth(450)
    .setHeight(530);
  DocumentApp.getUi().showModalDialog(html, 'Personalización de carátula');
}

function getData() {
  const ss = SpreadsheetApp.openById('1uCE8KJoVQtJzpVaK22ewlwkRwX9WUVxVDS1Cfjh5ql0');
  const sheet = ss.getSheetByName('Data');
  return {
    teacherNames: sheet.getRange('A12:A20').getValues().flat().filter(v => v !== ''),
    courseNames:  sheet.getRange('C2:C9').getValues().flat().filter(v => v !== ''),
    studentNames: sheet.getRange('D2:D93').getValues().flat().filter(v => v !== '')
  };
}

function createCover(course, teacher, student, font){
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  //Header
  body.appendParagraph('UNIVERSIDAD NACIONAL DE SAN AGUSTÍN')
    .setFontFamily(font).setFontSize(21).setBold(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setForegroundColor('#000000');
  body.appendParagraph('FACULTAD DE INGENIERÍA DE PROCESOS Y SERVICIOS')
    .setFontFamily(font).setFontSize(16).setBold(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setForegroundColor('#000000');
  body.appendParagraph('').setFontSize(16);

  body.appendParagraph('ESCUELA PROFESIONAL DE INGENIERÍA DE SISTEMAS')
    .setFontFamily(font).setFontSize(16).setBold(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setForegroundColor('#000000');

  body.appendParagraph('').setFontSize(16);
  //Image
  const urlImage = 'https://upload.wikimedia.org/wikipedia/commons/f/f9/Escudo_UNSA.png';
  const blob = UrlFetchApp.fetch(urlImage).getBlob();
  body.appendParagraph('')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .appendInlineImage(blob)
    .setWidth(4.24 * 72)
    .setHeight(5.30 * 72);
  //Middle Text
    body.appendParagraph('').setFontSize(16);
    body.appendParagraph('<Nombre de Actividad>')
    .setFontFamily(font).setFontSize(21).setBold(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setForegroundColor('#000000');
  body.appendParagraph('').setFontSize(16);

  body.appendParagraph(course)
    .setFontFamily(font).setFontSize(16)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setForegroundColor('#000000');

  body.appendParagraph('').setFontSize(16);

  body.appendParagraph(`Docente: ${teacher}`)
    .setFontFamily(font).setFontSize(16).setBold(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setForegroundColor('#000000');

  body.appendParagraph('').setFontSize(16);

  body.appendParagraph(`Alumno: ${student}`)
    .setFontFamily(font).setFontSize(16).setBold(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setForegroundColor('#000000');

  body.appendParagraph('').setFontSize(14);
  body.appendParagraph('').setFontSize(14);
  body.appendParagraph('').setFontSize(14);
  body.appendParagraph('').setFontSize(14);
  //Footer
  body.appendParagraph('Arequipa - 2025')
    .setFontFamily(font).setFontSize(16).setBold(true)
    .setForegroundColor('#d9d9d9')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  body.removeChild(body.getChild(0));

}
