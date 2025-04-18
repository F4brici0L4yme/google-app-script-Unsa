function modalDialog() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Insert Docentes')
  .addItem('Choose Text', 'showModal')
  .addToUi();
}

function showModal(){
  const ui = DocumentApp.getUi();
  const html = HtmlService.createHtmlOutputFromFile('dialogv2')
    .setWidth(200)
    .setHeight(300);
  ui.showModalDialog(html, 'Select Phrase');
}

function getData() {
  const ss = SpreadsheetApp.openById('1rggfyIeU4zJaUnIck-agsO-7M_TP_Hkdh_9Y2tWppMA');
  const sheet = ss.getSheetByName('Data');
  const data = sheet.getRange('A2:A7').getValues();

  return data.flat().filter(f => f !== '');
}

function insertText(text, font, size){
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  const textInserted = cursor.insertText(text)
  textInserted.setFontFamily(font);
  textInserted.setFontSize(size);
}


