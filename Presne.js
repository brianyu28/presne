/*
 * Pre-SNE v0.0.0
 * Presne.js / Code.gs
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

    // Get selected text, otherwise the whole document.
    var text = getSelectedText();
    if (text === null)
        text = DocumentApp.getActiveDocument().getBody().getText();
    else
        text = text.join("\n");

    result['contents'] = getCorrections(text);
    return result;
}

function getCorrections(text) {
    var errors = [];
    var houses = ['Adams', 'Eliot', 'Mather', 'Cabot', 'Kirkland', 'Pforzheimer', 'Currier',
              'Leverett', 'Quincy', 'Dudley', 'Lowell', 'Winthrop', 'Dunster'];
    var states = [['Alabama', 'Ala.'], ['Arizona', 'Ariz.'], ['Arkansas', 'Ark.'],
        ['California', 'Calif.'], ['Colorado', 'Colo.'], ['Connecticut', 'Conn.'],
        ['Delaware', 'Del.'], ['Florida', 'Fla.'], ['Georgia', 'Ga.'], ['Illinois', 'Ill.'],
        ['Indiana', 'Ind.'], ['Kansas', 'Kan.'], ['Kentucky', 'Ky.'], ['Louisiana', 'La.'],
        ['Maryland', 'Md.'], ['Michigan', 'Mich.'],
        ['Minnesota', 'Minn.'], ['Mississippi', 'Miss.'], ['Missouri', 'Mo.'],
        ['Montana', 'Mont.'], ['Nebraska', 'Neb.'], ['Nevada', 'Nev.'],
        ['New Hampshire', 'N.H.'], ['New Jersey', 'N.J.'], ['New Mexico', 'N.M.'],
        ['New York', 'N.Y.'], ['North Carolina', 'N.C.'], ['North Dakota', 'N.D.'],
        ['Oklahoma', 'Okla.'], ['Oregon', 'Ore.'], ['Pennsylvania', 'Pa.'],
        ['Rhode Island', 'R.I.'], ['South Carolina', 'S.C.'], ['South Dakota', 'S.D.'],
        ['Tennessee', 'Tenn.'], ['Vermont', 'Vt.'], ['Virginia', 'Va.'],
        ['Washington', 'Wa.'], ['West Virginia', 'W.Va.'], ['Wisconsin', 'Wis.'],
        ['Wyoming', 'Wyo.']];
    var months = [['January', 'Jan.'], ['February', 'Feb.'], ['March', 'Mar.'],
        ['April', 'Apr.'], ['August', 'Aug.'], ['September', 'Sept.'],
        ['October', 'Oct.'], ['November', 'Nov.'], ['December', 'Dec.']];

    // Punctuation errors.
    check(errors, replacementError(text, '‘', /‘\d\d/, 'Quotation Error',
        'Quotation mark should be facing the other direction!'));

    // First reference errors.
    check(errors, firstReference(text, 'Administrative Board', 'Ad Board'));
    check(errors, firstReference(text, 'Harvard University Dining Services', 'HUDS'));
    check(errors, firstReference(text, 'Harvard Lampoon, a semi-secret Sorrento Square social organization that used to occasionally publish a so-called humor magazine', 'Lampoon'));
    check(errors, firstReference(text, 'Undergraduate Council', 'UC'));
    check(errors, firstReference(text, 'University Health Services', 'UHS'));
    check(errors, firstReference(text, 'Malkin Athletic Center', 'the Mac'));
    check(errors, firstReference(text, 'Board of Overseers', 'the Overseers'));
    check(errors, firstReference(text, 'Board of Overseers', 'The Overseers'));
    houses.forEach(function(house) {
        check(errors, firstReference(text, house + ' House', house));
    });

    // Simple substitution errors.
    check(errors, shouldBe(text, 'Hist. and Lit.', 'Hist and Lit', /\bHist\. and Lit\./));
    check(errors, shouldBe(text, 'Arco', 'ARCO', null));
    check(errors, shouldBe(text, 'vice-president', 'vice president', null));
    check(errors, shouldBe(text, 'Raza', 'RAZA', null));
    check(errors, shouldBe(text, '%', 'percent', null));
    check(errors, shouldBe(text, 'Native American', 'American Indian', null));
    check(errors, shouldBe(text, 'utility entry port', 'manhole', null));
    check(errors, shouldBe(text, 'first year student', 'first-year student', null));
    check(errors, shouldBe(text, 'liveable', 'livable', null));
    check(errors, shouldBe(text, 'member of the Board of Overseers', 'Overseer', null));
    check(errors, shouldBe(text, 'Brattle Theater', 'Brattle Theatre', null));
    check(errors, shouldBe(text, 'Sanders Theater', 'Sanders Theatre', null));
    check(errors, shouldBe(text, 'Cantabridgians', 'Cantabrigians', null));
    check(errors, shouldBe(text, 'Massachusetts Hall', 'Mass. Hall', null));
    check(errors, shouldBe(text, 'Asian-American-Association', 'Asian American Association', null));
    check(errors, shouldBe(text, 'Massachusetts Avenue', 'Mass Ave.', null));
    check(errors, shouldBe(text, 'Massachusetts Bay Transportation Authority', 'MBTA', null));
    check(errors, shouldBe(text, 'Massachusetts Institute of Technology', 'MIT', null));
    check(errors, shouldBe(text, 'American Federation of Labor and Congress of Industrial Organizations', 'AFL-CIO', null));
    check(errors, shouldBe(text, 'Federal Bureau of Investigation', 'FBI', null));
    check(errors, shouldBe(text, 'Central Intelligence Agency', 'CIA', null));
    check(errors, shouldBe(text, 'Reserve Officers\' Training Corps', 'ROTC', null));
    check(errors, shouldBe(text, 'actress', 'actor', null));
    check(errors, shouldBe(text, 'advisor', 'adviser', null));
    check(errors, shouldBe(text, 'alumna', 'alum', null));
    check(errors, shouldBe(text, 'alumnae', 'alums', null));
    check(errors, shouldBe(text, 'alumnus', 'alum', null));
    check(errors, shouldBe(text, 'alumni', 'alums', null));
    check(errors, shouldBe(text, 'administration', 'Administration', null));
    check(errors, shouldBe(text, 'Street', 'St.', /\d+\s\w+\sStreet/));
    check(errors, shouldBe(text, 'Boulevard', 'Blvd.', /\d+\s\w+\sBoulevard/));
    check(errors, shouldBe(text, 'Road', 'Rd.', /\d+\s\w+\sRoad/));
    check(errors, shouldBe(text, 'St.', 'Street', /[^\d]\s\w+\sSt\./));
    check(errors, shouldBe(text, 'Blvd.', 'Boulevard', /[^\d]\s\w+\sBlvd\./));
    check(errors, shouldBe(text, 'Rd.', 'Road', /[^\d]\s\w+\sRd\./));
    states.forEach(function(state) {
        check(errors, shouldBe(text, state[0], state[1], null));
    });
    months.forEach(function(month)  {
        check(errors, shouldBe(text, month[0], month[1], new RegExp(month[0] + " \\d")));
        var numMatch = new RegExp(month[1] + " (\\d+)(?:st|nd|rd|th)");
        var match = numMatch.exec(text);
        if (match !== null) {
            check(errors, {
                'title': 'Date Formatting',
                'text': textAround(text, match.index, match[0].length),
                'contents': match[0] + ' should be ' + month[1] + ' ' + match[1]
            });
        }
    });

    // Time
    var time = /(\d\d?)?:(\d\d)?\s?(am|pm|AM|PM)\b/;
    var match = time.exec(text);
    if (match !== null) {

        if (!match[3].localeCompare("am") || !match[3].localCompare("AM"))
            var correct = match[1] + ':' + match[2] + ' a.m.';
        else
            var correct = match[1] + ':' + match[2] + ' p.m.';

        // If we have a :00 time.
        if (!match[2].localeCompare("00")) {
            if (!match[1].localeCompare("12")) {
                if (!match[3].localeCompare("am") || !match[3].localCompare("AM"))
                    correct = "midnight";
                else
                    correct = "noon";
            } else {
                if (!match[3].localeCompare("am") || !match[3].localCompare("AM"))
                    correct = match[1] + ' a.m.';
                else
                    correct = match[1] + ' p.m.';
            }
        }

        check(errors, {
            'title': 'Time Formatting',
            'text': textAround(text, match.index, match[0].length),
            'contents': match[0] + ' should be ' + correct
        });
    }

    // Avoid usage errors.
    check(errors, avoidUse(text, 'pro-life', /[Pp]ro-life/));
    check(errors, avoidUse(text, 'pro-choice', /[Pp]ro-choice/));
    check(errors, avoidUse(text, 'for her part', /[Ff]or her part/));
    check(errors, avoidUse(text, 'for his part', /[Ff]or his part/));
    check(errors, avoidUse(text, 'Asked whether', null));
    check(errors, avoidUse(text, 'it was learned', null));
    check(errors, avoidUse(text, 'The Crimson has learned', null));
    check(errors, avoidUse(text, 'CSA', null));

    // Preferable use errors.
    check(errors, preferableUse(text, 'approximately', 'about', /[Aa]pproximately/));

    // Cautious use errors.
    check(errors, cautiousUse(text, 'grumbled', null));
    check(errors, cautiousUse(text, 'sniffed', null));
    check(errors, cautiousUse(text, 'huffed', null));
    check(errors, cautiousUse(text, 'businessman', null));

    return generateErrors(errors);
}

/*
 * Handles first reference errors, where some phrase should be
 * some other phrase on first reference.
 */
function firstReference(text, first, subsequent) {
    var firstRegex = new RegExp('\\b' + first + '\\b');
    var subsequentRegex = new RegExp('\\b' + subsequent  +  '\\b');
    var subsequentMatch = subsequentRegex.exec(text);
    var firstMatch = firstRegex.exec(text);
    if (subsequentMatch !== null && (firstMatch == null ||
                                     firstMatch.index  > subsequentMatch.index))
        return {
            'title': 'First Reference',
            'text': textAround(text, subsequentMatch.index, subsequentMatch[0].length),
            'contents': 'Refer to \'' + subsequent + '\' as \'' + first + '\' on first reference'
        }
    else
        return null;
}

function replacementError(text, wrong, wrongRegex, title, message) {
    Logger.log(wrong);
    if (wrongRegex === null)
        wrongRegex = new RegExp('\\b' + wrong + '\\b');
    var wrongMatch = wrongRegex.exec(text);
    if (wrongMatch !== null)
        return {
            'title': title,
            'text': textAround(text, wrongMatch.index, wrongMatch[0].length),
            'contents': message
        }
    else
        return null;
}

/*
 * Handles when one phrase should be one other phrase.
 */
function shouldBe(text, wrong, right, wrongRegex) {
    return replacementError(text, wrong, wrongRegex,
        'Usage Error',
        '\'' + wrong + '\' should be \'' + right + '\'');
}

/*
 * Handles a phrase to be avoided.
 */
function avoidUse(text, wrong, wrongRegex) {
    return replacementError(text, wrong, wrongRegex,
        'Avoid Use',
        'Avoid use of \'' + wrong + '\'');
}

function preferableUse(text, wrong, right, wrongRegex) {
    return replacementError(text, wrong, wrongRegex,
        'Preferable Use',
        '\'' + right + '\' is preferable to \'' + wrong + '\'');
}

function cautiousUse(text, wrong, wrongRegex) {
    return replacementError(text, wrong, wrongRegex,
        'Cautious Use',
        'Be cautious with your use of \'' + wrong + '\', generally avoided');
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
            contents += '<div style="font-style:oblique;font-size:80%;">';
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
    var interval = 16;
    contents = '';
    if (index > interval)
        contents += '...';
    contents += text.substring(index - interval, index + length + interval);
    if (index + length + interval < text.length)
        contents += '...';
    return contents;
}

/*
 * Function to get selected text.
 * Adapted from Google.
 * https://developers.google.com/apps-script/quickstart/docs
 */
function getSelectedText() {
    var selection = DocumentApp.getActiveDocument().getSelection();
    if (selection) {
    var text = [];
    var elements = selection.getSelectedElements();
    for (var i = 0; i < elements.length; i++) {
        if (elements[i].isPartial()) {
            var element = elements[i].getElement().asText();
            var startIndex = elements[i].getStartOffset();
            var endIndex = elements[i].getEndOffsetInclusive();
            text.push(element.getText().substring(startIndex, endIndex + 1));
        } else {
            var element = elements[i].getElement();

            // Only translate elements that can be edited as text; skip images and
            // other non-text elements.
            if (element.editAsText) {
                var elementText = element.asText().getText();
                // This check is necessary to exclude images, which return a blank
                // text element.
                if (elementText != '') {
                    text.push(elementText);
                }
            }
        }
    }
    if (text.length == 0) {
        return null;
    }
    return text;
  } else {
      return null;
  }
}
