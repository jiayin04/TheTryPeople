function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('TryMenu')
    .addItem('Show Sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('My Trybar')
    .setWidth(500);
  
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}
