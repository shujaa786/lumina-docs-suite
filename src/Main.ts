function onOpen(): void {
  DocumentApp.getUi()
    .createMenu('Lumina')
    .addItem('Open Toolbar', 'showSidebar')
    .addToUi();
}

function showSidebar(): void {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Lumina Styling Toolbar')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}


