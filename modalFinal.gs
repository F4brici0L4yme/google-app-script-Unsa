const SHEET_ID   = '1uCE8KJoVQtJzpVaK22ewlwkRwX9WUVxVDS1Cfjh5ql0';
const SHEET_NAME = 'Data';

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Carátula')
    .addItem('Generar', 'showModal')
    .addToUi();
}

function showModal() {
  const html = HtmlService
    .createHtmlOutputFromFile('dialogFinal')
    .setWidth(500)
    .setHeight(500);
  DocumentApp.getUi().showModalDialog(html, 'Personalización de carátula');
}

function getData() {
  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName(SHEET_NAME);

  const teacherNames = sheet.getRange('A12:A').getValues().flat().filter(v => v !== '');
  const courseNames  = sheet.getRange('C2:C').getValues().flat().filter(v => v !== '');
  const studentNames = sheet.getRange('D2:D').getValues().flat().filter(v => v !== '');

  return {
    teacherNames,
    courseNames,
    studentNames
  };
}

function createCover(course, teacher, student, font) {
  const body  = DocumentApp.getActiveDocument().getBody();
  body.clear();
  const align = DocumentApp.HorizontalAlignment.CENTER;
  const black = '#000000';
  const gray  = '#d9d9d9';

  const append = (text, size, bold, color = black) =>
    body.appendParagraph(text)
        .setFontFamily(font)
        .setFontSize(size)
        .setBold(bold)
        .setForegroundColor(color)
        .setAlignment(align);

  const blank = size =>
    body.appendParagraph('').setFontSize(size);

  append('UNIVERSIDAD NACIONAL DE SAN AGUSTÍN',               21, true);
  append('FACULTAD DE INGENIERÍA DE PROCESOS Y SERVICIOS',    16, true);
  blank(16);
  append('ESCUELA PROFESIONAL DE INGENIERÍA DE SISTEMAS',    16, true);
  blank(16);

  const logoBlob = UrlFetchApp
    .fetch('https://upload.wikimedia.org/wikipedia/commons/f/f9/Escudo_UNSA.png')
    .getBlob();
  body.appendParagraph('')
      .setAlignment(align)
      .appendInlineImage(logoBlob)
      .setWidth(4.24 * 72)
      .setHeight(5.30 * 72);

  blank(16);
  append('<Nombre de Actividad>', 21, true);
  blank(16);
  append(course,               16, true);
  blank(16);
  append(`Docente: ${teacher}`, 16, true);
  blank(16);
  append(`Alumno: ${student}`,  16, true);

  for (let i = 0; i < 5; i++) blank(14);

  append('Arequipa - 2025', 16, true, gray);

  body.removeChild(body.getChild(0));
  body.appendParagraph('').setFontSize(11).setForegroundColor('#000000').setBold(false);
}
