/*
 * Pre-SNE v0.0.0
 * Code.js / Code.gs
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

    // First reference errors.
    check(errors, firstReference(body, "Undergraduate Council", "UC"));

    // Simple substitution errors.
    check(errors, shouldBe(body, "Hist. and Lit.", "Hist and Lit", /\bHist\. and Lit\./));

    return generateErrors(errors);
}

/*
 * Handles first reference errors, where some phrase should be
 * some other phrase on first reference.
 */
function firstReference(text, first, subsequent) {
    var firstRegex = new RegExp("\\b" + first + "\\b");
    var subsequentRegex = new RegExp("\\b" + subsequent  +  "\\b");
    var subsequentMatch = subsequentRegex.exec(text);
    var firstMatch = firstRegex.exec(text);
    if (subsequentMatch !== null && (firstMatch == null ||
                                     firstMatch.index  > subsequentMatch.index))
        return {
            'title': 'First Reference',
            'text': textAround(text, subsequentMatch.index, subsequentMatch[0].length),
            'contents': 'Refer to ' + subsequent + ' as ' + first + ' on first reference.'
        }
    else
        return null;
}

/*
 * Handles when one phrase should be one other phrase.
 */
function shouldBe(text, wrong, right, wrongRegex) {
    var wrongMatch = wrongRegex.exec(text);
    if (wrongMatch !== null)
        return {
            'title': 'Incorrect Use',
            'text': textAround(text, wrongMatch.index, wrongMatch[0].length),
            'contents': wrong + ' should be ' + right + '.'
        }
    else
        return null;
}

/*
 * Adds error to existing errors if there is an error.
 */
function check(errors, error) {
    if (error !== null)
        errors.push(error);
}

function generateErrors(errors) {
    var contents = '';
    for (i in errors) {
        contents += '<div style="color:red;font-weight:bold;">';
        contents += errors[i].title;
        contents += '</div>';
        if (errors[i].text !== undefined) {
            contents += '<div style="font-style:oblique">';
            contents += errors[i].text;
            contents += '</div>';
        }
        contents += errors[i].contents;
        contents += '<br/><br/>';
    }
    if (contents.length == 0) {
        contents = '<div style="color:green;font-weight:bold;">Looks good!</div>';
    }
    return contents;
}

/*
 * Gets text around a particular index.
 * Gets interval characters before and after.
 */
function textAround(text, index, length) {
    var interval = 10;
    contents = '';
    if (index > interval)
        contents += '...';
    contents += text.substring(index - interval, index + length + interval);
    if (index + length + interval < text.length)
        contents += '...';
    return contents;
}
