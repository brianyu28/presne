function onOpen(e) {
    DocumentApp.getUi().createAddonMenu()
                       .addItem('Start', 'showSidebar')
                       .addToUi();
}

function onInstall(e) {
    onOpen(e);
}

function showSidebar() {
    var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
                        .setTitle('Pre-SNE');
    DocumentApp.getUi().showSidebar(ui);
}

function runPreSNE() {
    var result = {};
    var body = DocumentApp.getActiveDocument().getBody().getText();
    result['contents'] = body;
    return result;
}
