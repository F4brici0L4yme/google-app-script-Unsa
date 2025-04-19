function onOpen() {
  DocumentApp.getUi()
    .createMenu('Carátula')
    .addItem('Generar', 'showModal')
    .addToUi();
}

function showModal(){
  const html = HtmlService
    .createHtmlOutputFromFile('dialogv4')
    .setWidth(400)
    .setHeight(530);
  DocumentApp.getUi().showModalDialog(html, 'Personalización de carátula');
}

function getData() {
  const ss = SpreadsheetApp.openById('1rggfyIeU4zJaUnIck-agsO-7M_TP_Hkdh_9Y2tWppMA');
  const sheet = ss.getSheetByName('Data');
  return {
    teacherNames: sheet.getRange('A9:A14').getValues().flat().filter(v => v !== ''),
    courseNames:  sheet.getRange('C2:C9').getValues().flat().filter(v => v !== ''),
    studentNames: sheet.getRange('D2:D45').getValues().flat().filter(v => v !== '')
  };
}

function createCover(course, teacher, student, font){
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  body.appendParagraph('UNIVERSIDAD NACIONAL DE SAN AGUSTÍN')
    .setFontFamily(font).setFontSize(21).setBold(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('FACULTAD DE INGENIERÍA DE PROCESOS Y SERVICIOS')
    .setFontFamily(font).setFontSize(16).setBold(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('').setFontSize(16);

  body.appendParagraph('ESCUELA PROFESIONAL DE INGENIERÍA DE SISTEMAS')
    .setFontFamily(font).setFontSize(16).setBold(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('').setFontSize(16);

  const urlImage = 'https://upload.wikimedia.org/wikipedia/commons/f/f9/Escudo_UNSA.png';
  const blob = UrlFetchApp.fetch(urlImage).getBlob();
  body.appendParagraph('')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .appendInlineImage(blob)
    .setWidth(4.24 * 72)
    .setHeight(5.30 * 72);

    body.appendParagraph('<Nombre de Actividad>')
    .setFontFamily(font).setFontSize(21).setBold(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('').setFontSize(16);

  body.appendParagraph(course)
    .setFontFamily(font).setFontSize(16)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('').setFontSize(16);

  body.appendParagraph(`Docente: ${teacher}`)
    .setFontFamily(font).setFontSize(16).setBold(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('').setFontSize(16);

  body.appendParagraph(`Alumno: ${student}`)
    .setFontFamily(font).setFontSize(16).setBold(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('').setFontSize(14);
  body.appendParagraph('').setFontSize(14);
  body.appendParagraph('').setFontSize(14);
  body.appendParagraph('').setFontSize(14);

  body.appendParagraph('Arequipa - 2025')
    .setFontFamily(font).setFontSize(16).setBold(true)
    .setForegroundColor('#d9d9d9')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}
