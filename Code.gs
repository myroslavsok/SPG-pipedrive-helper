// Addon setup settings
var ADDON_TITLE = 'SPG Pipedrive Look Up';

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
        .addItem('Settings', 'showSidebarSettings')
        .addItem('Look Up', 'showSidebarLookUp')
        .addToUi();
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebarLookUp() {
    var ui = HtmlService.createHtmlOutputFromFile('sidebar-look-up').setTitle(ADDON_TITLE);
    SpreadsheetApp.getUi().showSidebar(ui);
}

function showSidebarSettings() {
    var ui = HtmlService.createHtmlOutputFromFile('sidebar-settings').setTitle(ADDON_TITLE);
    SpreadsheetApp.getUi().showSidebar(ui);
}


// Functionality sidebar-look-up part
var SEARCH_QUERY_LIMIT = 500; // https://pipedrive.readme.io/docs/core-api-concepts-pagination maximum limit value is 500
var MARK_COLOR = '#99CC99';

var pipedriveSmartEmail = '';
var pipedriveApiKey = '';


function getColumnNamesAndIndexes() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var columnNames = sheet.getDataRange().offset(0, 0, 1).getValues()[0]; // Getting column names

    var columnNameAndIndexArr = [];
    for (var i = 0; i < columnNames.length; i++) {
        var columnNameObject = {
            columnName: columnNames[i],
            columnIndex: sheet.getRange(1, (i + 1), 1, 1).getA1Notation().match(/([A-Z]+)/)[0] // Get column index
        };
        columnNameAndIndexArr.push(columnNameObject);
    }
    return columnNameAndIndexArr;
}


function searchInPDByColumnValues(columnName, creds) {
    pipedriveApiKey = creds.pdak; // Set 'credentials'
    pipedriveSmartEmail = creds.pdse;

    const columnValueCells = getValueCellsByColumnName(columnName); // Getting column cells with values by column name

    // Start search by values from one column
    var options = {
        "method": "GET",
        "followRedirects": true,
        "muteHttpExceptions": true
    };
    var responseArr = [];
    columnValueCells.forEach(function (columnValueCell) {
        var targetUrl = generateSearchUrl(columnValueCell.value, 0, SEARCH_QUERY_LIMIT, pipedriveApiKey);
        try {
            var response = UrlFetchApp.fetch(targetUrl, options);
            if (response.getResponseCode() === 200) {
                Logger.log('response ' + response);
                var responseObj = JSON.parse(response.getContentText()); // Parse response to JS object
                if (responseObj.data) {   // Return only existed values
                    var foundResultItems = responseObj.data.map(function (foundItem) {
                        // responseObj.data is array of found items (deal, person, organization etc)
                        return {
                            id: foundItem.id,
                            title: foundItem.title,
                            type: foundItem.type,
                            cellValue: columnValueCell.value,
                            cellRowNumber: columnValueCell.rowNumber,
                            cellColumnNumber: columnValueCell.colNumber,
                            columnName: columnValueCell.colName,
                            url: generateFoundItemUrl(pipedriveSmartEmail.split('@')[0], foundItem.type, foundItem.id)
                        }
                    });
                    markAndCommentCellWithResemblances(foundResultItems); // Mark found cell with color and add comments (notes) to it
                    responseArr.push(foundResultItems);
                }
            }
        } catch (e) {
            Logger.log('error: ' + e); // Ignore possible error so as not to stop search
        }
    });

    return responseArr

// TODO handle pagination

//     var options = {
//         "method": "GET",
//         "followRedirects": true,
//         "muteHttpExceptions": true
//     };
//
//     var responsesDataArr = [];
//     var paginationOffset = 0;
//     do {
//         var targetURL = generateSearchUrl(paginationOffset, SEARCH_QUERY_LIMIT, pipedriveApiKeyValue); // Generate new link with page offset
//         var response = UrlFetchApp.fetch(targetURL, options);
//         if (response.getResponseCode() === 200) {
//             var responseObj = JSON.parse(response.getContentText()); // Parse response to JS object
//             paginationOffset += SEARCH_QUERY_LIMIT; // Increase page offset
//             if (responseObj.data) { // Avoid last extra query with null value
//                 responsesDataArr.push(responseObj.data);
//             }
//         } else {
//             return response.getResponseCode + ' failed response code';
//         }
//     } while (responseObj.data); // Run cycle until we get data from server
//     return [].concat.apply([], responsesDataArr) // Array of arrays (array of pages) to one array;
}


function getValueCellsByColumnName(columnName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var colNumber = data[0].indexOf(columnName); // Column position number

    var columnValues = [];
    for (var i = 1; i <= sheet.getLastRow() - 1; i++) { // Skips first row with column titles
        columnValues.push({
            rowNumber: i,
            colNumber: colNumber,
            colName: columnName,
            value: data[i][colNumber]
        });
    }

    return columnValues.filter(function (columnCell) {
        if (!!columnCell.value === true && columnCell.value !== 'null' && columnCell.value !== 'undefined') {
            return columnCell;
        }
    }); // delete empty values (do not implement search using empty values, null, undefined or it will mark as found everything)
}


// Generate fetch url for search (SearchResults API)
function generateSearchUrl(term, paginationOffset, dataLimit, apiToken) {
    var formatedTerm = term
        .slice(0, SEARCH_QUERY_LIMIT) // Temp length limit is 500 characters
        .replace(new RegExp(/[`~#&*]/gi), ''); // Delete special characters so as to encode url with encodeURI
    // https://stackoverflow.com/questions/332872/encode-url-in-javascript/332897#332897?newreg=d7be5054e80b4f948cdcf3129cbc3aaa

    var targetUrl = 'https://api.pipedrive.com/v1/searchResults?term=' + formatedTerm + '&start=' + paginationOffset + '&limit=' + dataLimit + '&api_token=' + apiToken;
    return encodeURI(targetUrl);
}


function generateFoundItemUrl(PDsmartEmail, itemType, itemId) {
    return 'https://' + PDsmartEmail + '.pipedrive.com/' + itemType + '/' + itemId;
}


function markAndCommentCellWithResemblances(foundResultItems) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getDataRange();
    var cellColumnNumber = ++foundResultItems[0].cellColumnNumber;
    var cellRowNumber = ++foundResultItems[0].cellRowNumber;
    var targetCell = range.getCell(cellRowNumber, cellColumnNumber);
    targetCell.setBackground(MARK_COLOR); // Mark/paint cell

    // Add notes to cell
    var comment = '';
    foundResultItems.forEach(function (foundItem) {
        comment += foundItem.url + '\n';
    });
    targetCell.setComment(comment);
}


// Apply white color to the background of each row
function clearMark() {
    var range = SpreadsheetApp.getActiveSheet().getDataRange();
    for (var i = range.getRow(); i < range.getLastRow(); i++) {
        range.offset(i, 0, 1).setBackgroundColor('white');
        range.offset(i, 0, 1).setComment('');
    }
}


