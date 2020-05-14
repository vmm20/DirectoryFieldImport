
// Creates a menu entry in the Google Sheets UI when the document is opened.
function onOpen(e) {
  SpreadsheetApp.getUi().createMenu('Directory Import')
  .addItem('Import directory field', 'showSidebar')
  .addToUi();
}
function onInstall(e) {
  onOpen(e);
}


// Opens a sidebar in the spreadsheet containing the add-on's user interface.
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('sidebar');
  ui.menuDefault = "---Select---";
  ui.user = getUserInfo();
  ui.lastColumn = SpreadsheetApp.getActiveSheet().getLastColumn();
  SpreadsheetApp.getUi().showSidebar(ui.evaluate().setTitle('Import Directory Field'));
}



/**
* Intended for use when the Import button is clicked in the sidebar.
*/
function import(mode, emailColumn, fieldAbbr, destColumn) {
  var numRows = SpreadsheetApp.getActiveSheet().getLastRow();
  var emailRange = SpreadsheetApp.getActiveSheet().getRange(1, emailColumn, numRows, 1);
  var destRange = SpreadsheetApp.getActiveSheet().getRange(1, destColumn, numRows, 1);
  var emailValues = emailRange.getValues();
  var oldDestValues = destRange.getValues();
  var destValues = [];
  var failed = 0;
  var skipped = 0;
  
  var destValue = "";
  if(mode == 'directory') {
    emailValues.forEach(function(emailRow, emailRowIndex) {
      var emailValue = emailRow[0];
      var oldDestValue = oldDestValues[emailRowIndex][0];
      if(oldDestValue) {
        skipped++;
        destValues.push([oldDestValue]);
        return;
      }
      var userDirObj = null;
      try {
        userDirObj = AdminDirectory.Users.get(emailValue, 
                                              {viewType: "domain_public",
                                               fields: "name,organizations(title,department)",
                                              });
      } catch(e) {
        failed++;
      }
      destValues.push([userDirObj != null && fields[fieldAbbr].directory ? fields[fieldAbbr].directory(userDirObj) : ""]);
    });
  } else if(mode == 'contacts') {
    var contObjs = getContObjs();
    emailValues.forEach(function(emailRow) {
      var emailValue = emailRow[0];
      var oldDestValue = oldDestValues[emailRowIndex][0];
      if(oldDestValue) {
        skipped++;
        destValues.push([oldDestValue]);
        return;
      }
      var userContObj = null;
      try {
        userContObj = contObjs.filter(function(contObj) {
          return contObj.emailAddresses && contObj.emailAddresses.length > 0 && contObj.emailAddresses[0].value == emailValue;
        })[0];
      } catch(e) {
        failed++;
      }
      destValues.push([userContObj != null && fields[fieldAbbr].contacts ? fields[fieldAbbr].contacts(userContObj) : ""]);
    });
  }
  var valuesImported = numRows-skipped-failed;
  if(destValues[0][0] == "") {
    destValues[0][0] = fields[fieldAbbr].dispName + " (" + mode + ")";  // add column header if first cell is not filled in
  }
  destRange.setValues(destValues);
  var alertText = valuesImported + (valuesImported==1 ? " value has" : " values have") + " been imported.";
  if(skipped) {
    if(skipped == 1) {
      alertText += "\n  " + skipped + " value has been skipped because its cell already contained a value.";
    } else {
      alertText += "\n  " + skipped + " values have been skipped because their cells already contained values.";
    }
  }
  if(failed) {
    if(failed == 1) {
      alertText += "\n  " + failed + " value has failed to import.  The email address is likely invalid.";
    } else {
      alertText += "\n  " + failed + " values have failed to import.  The email addresses are likely invalid.";
    }
  }
  SpreadsheetApp.getUi().alert("Import Complete", alertText, 
    SpreadsheetApp.getUi().ButtonSet.OK);
}


function getContObjs() {
  var contacts = [];
  var pageToken = "";
  do {
    var connectionList = People.People.Connections.list('people/me', {
      pageSize: 200,
      personFields: 'names,emailAddresses,organizations',
      pageToken: pageToken,
    });
    contacts = contacts.concat(connectionList.connections);
    pageToken = connectionList.nextPageToken;
  } while(pageToken){}
  Logger.log(contacts);
  return contacts;
}


function getUserInfo() {
  var userInfo = JSON.parse(UrlFetchApp.fetch("https://www.googleapis.com/oauth2/v2/userinfo", {
    "headers": {
      "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
    }, "muteHttpExceptions": true
  }));  // gets HTTP response containing user data
  var dirResponse = JSON.parse(UrlFetchApp.fetch("https://www.googleapis.com/admin/directory/v1/users?customer=my_customer&viewType=domain_public", {
    "headers": {
      "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
    }, "muteHttpExceptions": true
  }));  // gets HTTP response verifying the user's directory access
  userInfo.hasDirAccess = userInfo.hd != null && dirResponse.users != null && dirResponse.users.length > 1;
  return userInfo;
}

