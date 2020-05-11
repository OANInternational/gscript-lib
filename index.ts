//#region KOBOTOOLS

/**
 * Get all submitions of a Kobo form as a matrix
 *
 * @param {number} id_form the id of the KoboForm
 * @return {any[][]} the data as a matrix of values
 */
function kobo_getFormEntries(id_form: string): any[][] {
  const datacsv = UrlFetchApp.fetch(
    `https://kc.humanitarianresponse.info/api/v1/data/${id_form}.csv`,
    getheaders_()
  ).getContentText();
  return CSVToArray_(datacsv);
}
function kobo_getFormData(id_form: string) {
  const datas = kobo_getFormsData_();
  return datas.find((data) => data.formid === id_form);
}
function kobo_writeData(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  Entriesdata
) {
  for (var i = 0; i < Entriesdata.length; i++) {
    for (var j = 0; j < Entriesdata[i].length; j++) {
      sheet.getRange(i + 1, j + 1).setValue(Entriesdata[i][j]);
    }
  }
}
function kobo_get_credentials_() {
  const user = PropertiesService.getScriptProperties().getProperty('kobo_user');
  const pass = PropertiesService.getScriptProperties().getProperty('kobo_pass');
  return [user, pass];
}
function kobo_getFormsData_() {
  const data = UrlFetchApp.fetch(
    `https://kc.humanitarianresponse.info/api/v1/forms`,
    getheaders_()
  ).getContentText();
  return JSON.parse(data);
}
function getheaders_() {
  const cred = kobo_get_credentials_();
  return {
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(cred[0] + ':' + cred[1]),
    },
    muteHttpExceptions: true,
  };
}
function CSVToArray_(strData, strDelimiter: string = ',') {
  // Create a regular expression to parse the CSV values.
  var objPattern = new RegExp(
    // Delimiters.
    '(\\' +
      strDelimiter +
      '|\\r?\\n|\\r|^)' +
      // Quoted fields.
      '(?:"([^"]*(?:""[^"]*)*)"|' +
      // Standard fields.
      '([^"\\' +
      strDelimiter +
      '\\r\\n]*))',
    'gi'
  );

  // Create an array to hold our data. Give the array
  // a default empty first row.
  var arrData = [[]];

  // Create an array to hold our individual pattern
  // matching groups.
  var arrMatches = null;

  // Keep looping over the regular expression matches
  // until we can no longer find a match.
  while ((arrMatches = objPattern.exec(strData))) {
    // Get the delimiter that was found.
    var strMatchedDelimiter = arrMatches[1];

    // Check to see if the given delimiter has a length
    // (is not the start of string) and if it matches
    // field delimiter. If id does not, then we know
    // that this delimiter is a row delimiter.
    if (strMatchedDelimiter.length && strMatchedDelimiter != strDelimiter) {
      // Since we have reached a new row of data,
      // add an empty row to our data array.
      arrData.push([]);
    }

    // Now that we have our delimiter out of the way,
    // let's check to see which kind of value we
    // captured (quoted or unquoted).
    if (arrMatches[2]) {
      // We found a quoted value. When we capture
      // this value, unescape any double quotes.
      var strMatchedValue = arrMatches[2].replace(new RegExp('""', 'g'), '"');
    } else {
      // We found a non-quoted value.
      var strMatchedValue = arrMatches[3];
    }

    // Now that we have our value string, let's add
    // it to the data array.
    arrData[arrData.length - 1].push(strMatchedValue);
  }

  // Return the parsed data.
  return arrData;
}

//#endregion

//#region UI

function ui_showLoading() {
  // Display a modal dialog box with custom HtmlService content.
  var htmlOutput = HtmlService.createHtmlOutput();
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Cargando...');
}
function ui_hideLoading() {
  var output = HtmlService.createHtmlOutput(
    '<script>google.script.host.close();</script>'
  );
  SpreadsheetApp.getUi().showModalDialog(output, 'Cargando...');
}

//#endregion

//#region FIREBASE

//#endregion
