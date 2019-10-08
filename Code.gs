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
var ORGANIZATIONS_LIMIT = 500; // https://pipedrive.readme.io/docs/core-api-concepts-pagination maximum limit value is 500
var MARK_COLOR = '#99CC99';


function findResemblances(columnName, pipedriveApiKeyValue) {
    if (!columnName) { // ColumnName must not be empty
        return;
    }
    var PDOrganizations = getAllPDOrganizations(pipedriveApiKeyValue);
    var SSOrganizations = getAllSSOrganizationsByColumnName(columnName);

    var resemblingOrganizations = [];
    SSOrganizations.forEach(function (SSOrganization) {    // Find organizations resemblances
        PDOrganizations.forEach(function (PDOrganization) {
            for (var key in PDOrganization) {
                if (PDOrganization.hasOwnProperty(key) && // check for existence on object key
                    PDOrganization[key] && SSOrganization.value && // Check values for undefined and null necessary for .toString()
                    (PDOrganization[key]).toString().toLowerCase() === (SSOrganization.value).toString().toLowerCase()) { // Target check
                    resemblingOrganizations.push(SSOrganization);
                }
            }
        });
    });
    markOrganizationResemblances(resemblingOrganizations);
    return {
        SSOrganizations: SSOrganizations,
        PDOrganizations: PDOrganizations,
        resemblingOrganizations: resemblingOrganizations
    };
}


function getAllPDOrganizations(pipedriveApiKeyValue) {
    var options = {
        "method": "GET",
        "followRedirects": true,
        "muteHttpExceptions": true
    };

    var responsesDataArr = [];
    var paginationOffset = 0;
    do {
        var targetURL = generateSearchUrl(paginationOffset, ORGANIZATIONS_LIMIT, pipedriveApiKeyValue); // Generate new link with page offset
        var response = UrlFetchApp.fetch(targetURL, options);
        if (response.getResponseCode() === 200) {
            var responseObj = JSON.parse(response.getContentText()); // Parse response to JS object
            paginationOffset += ORGANIZATIONS_LIMIT; // Increase page offset
            if (responseObj.data) { // Avoid last extra query with null value
                responsesDataArr.push(responseObj.data);
            }
        } else {
            return response.getResponseCode + ' failed response code';
        }
    } while (responseObj.data); // Run cycle until we get data from server
    return [].concat.apply([], responsesDataArr) // Array of arrays (array of pages) to one array;
}

// Generate fetch link for pipedrive organizations
function generateSearchUrl(paginationOffset, dataLimit, apiToken) {
    return 'https://api.pipedrive.com/v1/organizations?start=' + paginationOffset + '&limit=' + dataLimit + '&api_token=' + apiToken;
}

function getAllSSOrganizationsByColumnName(columnName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var colNumber = data[0].indexOf(columnName); // Column position number
    if (colNumber !== -1) {
        var SSOrganizationsColumnArr = [];
        for (var i = 1; i <= sheet.getLastRow() - 1; i++) { // Skips first row with column titles
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
    resemblingOrganizations.forEach(function (resemblingOrganization) {
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
