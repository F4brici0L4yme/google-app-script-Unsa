const SHEET_ID   = '1uCE8KJoVQtJzpVaK22ewlwkRwX9WUVxVDS1Cfjh5ql0';
const SHEET_NAME = 'Data';

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Carátula')
    .addItem('Generar', 'showModal')
    .addToUi();
}

function showModal() {
  const html = HtmlService.createHtmlOutputFromFile('dialogFinal')
    .setWidth(500)
    .setHeight(600);
  DocumentApp.getUi().showModalDialog(html, 'Personalización de Carátula');
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

function createCover(course, teacher, students, font) {
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
  append('ESCUELA PROFESIONAL DE INGENIERÍA DE SISTEMAS',     16, true);
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
  append(course, 16, true);
  blank(16);
  append(`Docente: ${teacher}`, 16, true);
  blank(16);

  if (students.length === 1) {
    append('Alumno:', 16, true);
    append(students[0], 16, false);
  } else if (students.length >= 2 && students.length <= 3) {
    append('Integrantes:', 16, true);
    students.forEach(s => append(s, 16, false));
  } else {
    body.removeChild(body.getChild(0));
    append('Integrantes:', 16, true);
    const tableData = [];
    for (let i = 0; i < students.length; i += 2) {
      const row = [
        students[i],
        students[i + 1] || '' // Si hay número impar
      ];
      tableData.push(row);
    }
    const table = body.appendTable(tableData);
    table.setBorderWidth(0);
    table.getRows().forEach(row => {
      row.getCells().forEach(cell => {
        const paragraph = cell.getChild(0).asParagraph();
        paragraph.setFontSize(14);
        paragraph.setFontFamily(font);
        paragraph.setForegroundColor(black);
        paragraph.setAlignment(align);
        paragraph.setBold(false);
        cell.setBorderWidth(0);
      });
    });
  }

  for (let i = 0; i < 5; i++) blank(14);
  append('Arequipa - 2025', 16, true, gray);

  body.removeChild(body.getChild(0));
  body.appendParagraph('').setFontSize(11).setForegroundColor('#000000').setBold(false);
}
