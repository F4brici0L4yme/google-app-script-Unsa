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
    .setHeight(500);
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
  const body = DocumentApp.getActiveDocument().getBody();
  body.clear();

  const align = DocumentApp.HorizontalAlignment.CENTER;
  const black = '#000000';
  const gray  = '#d9d9d9';

  // Función para agregar un párrafo centrado
  function append(text, size, bold = false, color = black) {
    body.appendParagraph(text)
        .setFontFamily(font)
        .setFontSize(size)
        .setBold(bold)
        .setForegroundColor(color)
        .setAlignment(align);
  }

  // Función para línea en blanco
  function blank(size) {
    body.appendParagraph('')
        .setFontFamily(font)
        .setFontSize(size)
        .setAlignment(align);
  }

  // Formatea cada celda de una tabla
  function formatTable(table, size) {
    const rows = table.getNumRows();
    for (let r = 0; r < rows; r++) {
      const row = table.getRow(r);
      const cells = row.getNumCells();
      for (let c = 0; c < cells; c++) {
        const cell = row.getCell(c);
        // Recorre todos los párrafos dentro de la celda
        for (let i = 0; i < cell.getNumChildren(); i++) {
          const child = cell.getChild(i);
          if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const p = child.asParagraph();
            p.setFontFamily(font)
             .setFontSize(size)
             .setBold(false)
             .setForegroundColor(black)
             .setAlignment(align);
          }
        }
        cell.setBorderWidth(0);
      }
    }
  }

  // ----- Inicia creación de carátula -----
  append('UNIVERSIDAD NACIONAL DE SAN AGUSTÍN',               21, true);
  append('FACULTAD DE INGENIERÍA DE PROCESOS Y SERVICIOS',    16, true);
  blank(16);
  append('ESCUELA PROFESIONAL DE INGENIERÍA DE SISTEMAS',    16, true);
  blank(16);

  // Logo
  const logo = UrlFetchApp
    .fetch('https://upload.wikimedia.org/wikipedia/commons/f/f9/Escudo_UNSA.png')
    .getBlob();
  body.appendParagraph('')
      .setAlignment(align)
      .appendInlineImage(logo)
      .setWidth(4.24 * 72)
      .setHeight(5.30 * 72);

  blank(16);
  append('<Nombre de Actividad>', 21, true);
  blank(16);
  append(course, 16, true);
  blank(16);
  const teacherParagraph = body.appendParagraph('');
  const teacherText1 = teacherParagraph.appendText('Docente: ');
  const teacherText2 = teacherParagraph.appendText(teacher);
  teacherParagraph.setFontFamily(font).setFontSize(16).setAlignment(align);
  teacherText1.setBold(true);
  teacherText2.setBold(false);

  blank(16);

  // Dependiendo del número de estudiantes
  if (students.length === 1) {
    const studentParagraph = body.appendParagraph('');
    const studentText1 = studentParagraph.appendText('Alumno: ');
    const studentText2 = studentParagraph.appendText(students[0]);
    studentParagraph.setFontFamily(font).setFontSize(16).setAlignment(align);
    studentText1.setBold(true);
    studentText2.setBold(false);
  }
  else if (students.length <= 3) {
    append('Integrantes:', 16, true);
    students.forEach(name => append(name, 16, false));
  }
  else {
    append('Integrantes:', 16, true);

    // Prepara filas de dos columnas
    const tableData = [];
    for (let i = 0; i < students.length; i += 2) {
      tableData.push([ students[i], students[i+1] || '' ]);
    }

    const table = body.appendTable(tableData);
    table.setBorderWidth(0);

    // Formatea toda la tabla a tamaño 16 y alineado
    formatTable(table, 16);
  }

  // Pie de página
  for (let i = 0; i < 3; i++) blank(14);
  append('Arequipa - 2025', 16, true, gray);

  // Limpieza de primer párrafo vacío
  body.removeChild(body.getChild(0));
  body.appendParagraph('')
      .setFontFamily(font)
      .setFontSize(11)
      .setBold(false)
      .setForegroundColor(black)
      .setAlignment(align);
}
