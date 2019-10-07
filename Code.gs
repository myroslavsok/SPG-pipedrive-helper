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
var urlOrganizations = 'https://api.pipedrive.com/v1/organizations?start=0&api_token=';

var pipedriveDataFieldNames = {
    organizationName: 'name',
    organizationStand: '48389fc5d6a135fb61fa640d7bc0535ad5823b68',
    organizationDescription: '5788bcb5084f1a5793599bd082df00619a9ddadf',
    organizationWebsite: 'cb65d2bb8ea467c23f826a488bb0d5488ed72408',
    organizationProfilePage: '06ebfe7f9f3beead99525fae25a8baede1648164'
};

var spreadSheetFieldNameEquivalents = {

};

function findResemblances(columnName) {
    var PDOrganizations = getAllPDOrganizations();
    var SSOrganizations = getAllSSOrganizationsByColumnName(columnName);
    Logger.log('PDOrganizations ' + PDOrganizations);
    Logger.log('SSOrganizations ' + SSOrganizations);

}


function getAllPDOrganizations() {
    var token = "455f56a33c14b568424d956a37638cb19453b2a8";
    var url = urlOrganizations + token;
    var options = {
        "method": "GET",
        "followRedirects": true,
        "muteHttpExceptions": true
    };
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
        var responseObject = JSON.parse(response.getContentText());
        return responseObject;
    }
    return response.getResponseCode;
}

function getAllSSOrganizationsByColumnName(columnName) {
    return columnName;
}

