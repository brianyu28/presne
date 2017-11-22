/*
 * Pre-SNE v0.0.0
 */

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
    result['contents'] = getCorrections(body);
    return result;
}

function getCorrections(body) {
    errors = [];

    if (body.indexOf('UC') != -1 && body.indexOf('Undergraduate Council') == -1) {
        errors.push({
            'title': 'Style: First Reference',
            'contents': 'Refer to UC as Undergraduate Council on first reference.'
        });
    }
    
    return generateErrors(errors);
}

function generateErrors(errors) {
    var contents = '';
    for (i in errors) {
        contents += '<div style="color:red;font-weight:bold;">';
        contents += errors[i].title;
        contents += '</div>';
        contents += errors[i].contents;
        contents += '<br/><br/>';
    }
    if (contents.length == 0) {
        contents = '<div style="color:green;font-weight:bold;">Looks good!</div>';
    }
    return contents;
}
