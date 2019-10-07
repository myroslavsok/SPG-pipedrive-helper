// Addon setup settings

var ADDON_TITLE = 'SPG pipedrive helper';

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
    onOpen(e);
}

/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
    SpreadsheetApp.getUi().createAddonMenu()
        .addItem('Start', 'showSidebar')
        .addToUi();
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
    var ui = HtmlService.createHtmlOutputFromFile('sidebar').setTitle(ADDON_TITLE);
    SpreadsheetApp.getUi().showSidebar(ui);
}


// Functionality part
var ORGANIZATIONS_LIMIT = 5000000; // TODO make this part more clear
var URL_ORGANIZATIONS = 'https://api.pipedrive.com/v1/organizations?start=0&limit=' + ORGANIZATIONS_LIMIT + '&api_token=';
var MARK_COLOR = '#99CC99';

// var pipedriveDataFieldNames = {
//     organizationName: 'name',
//     organizationStand: '48389fc5d6a135fb61fa640d7bc0535ad5823b68',
//     organizationDescription: '5788bcb5084f1a5793599bd082df00619a9ddadf',
//     organizationWebsite: 'cb65d2bb8ea467c23f826a488bb0d5488ed72408',
//     organizationProfilePage: '06ebfe7f9f3beead99525fae25a8baede1648164'
// };
// var spreadSheetFieldNameEquivalents = {};


function findResemblances(columnName, pipedriveApiKeyValue) {
    if (!columnName) {
        return;
    }
    var PDOrganizationsResponse = getAllPDOrganizations(pipedriveApiKeyValue);
    var PDOrganizations = [].concat.apply([], PDOrganizationsResponse.data); // array of arrays to array
    var SSOrganizations = getAllSSOrganizationsByColumnName(columnName); // columnName must not be empty
    // Find organizations resemblances
    var resemblingOrganizations = [];
    SSOrganizations.forEach(function(SSOrganization) {
        PDOrganizations.forEach(function(PDOrganization) {
            for (var key in PDOrganization) {
                if (PDOrganization.hasOwnProperty(key) && PDOrganization[key] === SSOrganization.value) {
                    resemblingOrganizations.push(SSOrganization);
                }
            }
        });
    });
    markOrganizationResemblances(resemblingOrganizations);
    return {
        SSOrganizations: SSOrganizations,
        PDOrganizations: PDOrganizations,
        resemblingOrganizations: resemblingOrganizations,
    };
}


function getAllPDOrganizations(pipedriveApiKeyValue) {
    var url = URL_ORGANIZATIONS + pipedriveApiKeyValue;
    var options = {
        "method": "GET",
        "followRedirects": true,
        "muteHttpExceptions": true
    };
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
        return JSON.parse(response.getContentText());
    }
    return response.getResponseCode + ' failed code';
}


function getAllSSOrganizationsByColumnName(columnName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var colNumber = data[0].indexOf(columnName); // Column position number
    if (colNumber !== -1) {
        var SSOrganizationsColumnArr = [];
        // Skips first row with column titles
        for (var i = 1; i <= sheet.getLastRow() - 1; i++) {
            SSOrganizationsColumnArr.push({
                rowNumber: i,
                colNumber: colNumber,
                colName: columnName,
                value: data[i][colNumber]
            });
        }
        return SSOrganizationsColumnArr;
    }
}

function markOrganizationResemblances(resemblingOrganizations) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getDataRange();
    resemblingOrganizations.forEach(function(resemblingOrganization) {
       range.offset(resemblingOrganization.rowNumber, 0, 1).setBackground(MARK_COLOR);
    });
}


// Apply white color to the background of each row
function clearMark() {
    var range = SpreadsheetApp.getActiveSheet().getDataRange();
    for (var i = range.getRow(); i < range.getLastRow(); i++) {
        range.offset(i, 0, 1).setBackgroundColor('white');
    }
}


// function test() {
//     var sheet = SpreadsheetApp.getActiveSheet();
//     var data = sheet.getDataRange().getValues();
//     var colName = 'Name';
//     var col = data[0].indexOf(colName);
//     Logger.log(sheet.getLastRow() + " Is the last Column.");
//     if (col != -1) {
//         Logger.log(data[1][col]);
//         Logger.log('col: ' + col);
//     }
// }
