function modalDialog() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Carátula')
  .addItem('Generar', 'showModal')
  .addToUi();
}

function showModal(){
  const ui = DocumentApp.getUi();
  const html = HtmlService.createHtmlOutputFromFile('dialogv3')
    .setWidth(300)
    .setHeight(400);
  ui.showModalDialog(html, 'Personalization');
}

function getData() {
  const ss = SpreadsheetApp.openById('1rggfyIeU4zJaUnIck-agsO-7M_TP_Hkdh_9Y2tWppMA');
  const sheet = ss.getSheetByName('Data');
  const teacherNames = sheet.getRange('A9:A14').getValues().flat().filter(v => v !== '');
  const courseNames = sheet.getRange('C2:C9').getValues().flat().filter(v => v !== '');
  const studentNames = sheet.getRange('D2:D45').getValues().flat().filter(v => v !== '');

  return {
    teacherNames,
    courseNames,
    studentNames
  };
}

function createCover(course, teacher, student, font){
  const doc = DocumentApp.getActiveDocument();
  const urlImage = "'https://upload.wikimedia.org/wikipedia/commons/f/f9/Escudo_UNSA.png'";
  const body = doc.getBody();

  const headText1 = 'UNIVERSIDAD NACIONAL DE SAN AGUSTÍN';
  const headText2 = 'FACULTAD DE INGENIERÍA DE PROCESOS Y SERVICIOS';
  const headText3 = 'ESCUELA PROFESIONAL DE INGENIERÍA DE SISTEMAS';

  const answer = UrlFetchApp.fetch(url);
  const imageBlob = answer.getBlob();

  //const worksName
  //const courseText
  //const teacherName
  //const studentName
  const footer = 'Arequipa - 2025';
  const cursor = doc.getCursor();
  cursor.insertText(headText1).setFontFamily(font).setFontSize(21).setBold(true).setAlignment(DocumentAppHorizontalAlignment.CENTER);
  cursor.insertText(headText2).setFontFamily(font).setFontSize(16).setBold(true).setAlignment(DocumentAppHorizontalAlignment.CENTER);
  cursor.insertText(headText3).setFontFamily(font).setFontSize(16).setBold(true).setAlignment(DocumentAppHorizontalAlignment.CENTER);

  const imgParagraph = body.appendParagraph("");
  const finalImage = imgParagraph.appendInlineImage(imageBlob);
  imgParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  finalImage.setWidth(300);
  finalImage.setHeight(300);

  body.appendParagraph(""); //To fill works Name
  cursor.insertText(course);
  textInserted.setFontFamily(font).setFontSize(16);
  cursor.insertText("Docente: ${teacher}").setFontFamily(font).setFontSize(16).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  cursor.insertText("Docente: ${student}").setFontFamily(font).setFontSize(16).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph("");
  body.appendParagraph("");
  body.appendParagraph("");
  body.appendParagraph("");
  cursor.insertText(footer).setFontFamily(font).setFontSize(16).setForegroundColor("#d9d9d9").setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

}