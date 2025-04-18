function modalDialog() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Insert Text')
  .addItem('Choose Text', 'showModal')
  .addToUi();
}

function showModal(){
  const html = HtmlService.createHtmlOutputFromFile('dialog');
  html.setWidth(200);
  html.setHeight(300);
  DocumentApp.getUi().showModalDialog(html, 'Personalization')
}

function insertText(text, font, size){
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  const textInserted = cursor.insertText(text)
  textInserted.setFontFamily(font);
  textInserted.setFontSize(size);
}