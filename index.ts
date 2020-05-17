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
function ui_createMenu() {}

//#endregion

//#region FIREBASE

// https://github.com/grahamearley/FirestoreGoogleAppsScript
// lib id: 1VUSl4b1r1eoNcRWotZM3e87ygkxvXltOgyDZhixqncz9lQ3MjfT1iKFwç

function fire_get_credentials_(): any {
  const cred = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('fire_contoan')
  );
  return {
    private_key: cred.private_key,
    client_email: cred.client_email,
    project_id: cred.project_id,
  };
}
function fire_get_firestore_(): any {
  const cred = fire_get_credentials_();

  return FirestoreApp.getFirestore(
    cred.client_email,
    cred.private_key,
    cred.project_id
  );
}
function fire_getUsers() {
  const firestore = fire_get_firestore_();
  const users = firestore.getDocuments('Users');
  const new_users = [];

  // FORMAT RECIVING DATA TO OBJECT ARRAY WITH MIN INFO
  for (let i = 0; i < users.length; i++) {
    const user = users[i]['fields'];
    const new_user = {
      id: user['id'],
      name: user['name'] + ' ' + user['last_name'],
      account_id: null,
      account_value: null,
    };

    // GET ACCOUNT INFO FROM THE USER IF IT HAS ONE ASSIGNED
    if (user['own_account']) {
      const account = firestore
        .query('Accounts')
        .where('user_id', '==', new_user['id'])
        .execute();

      // CHECK IF AN ACCOUNT WAS FOUND
      if (account.length > 0) {
        new_user['account_id'] = account[0]['fields']['id'];
        new_user['account_value'] = account[0]['fields']['value'];
      }
    }

    new_users.push(new_user);
  }
  return new_users;
}
function fire_getAccounts() {
  const firestore = fire_get_firestore_();
  const accounts = firestore.getDocuments('Accounts');
  const new_accounts = [];

  // FORMAT RECIVING DATA TO OBJECT ARRAY WITH MIN INFO
  for (let i = 0; i < accounts.length; i++) {
    const account = accounts[i]['fields'];
    const new_account = {
      id: account['id'],
      name: account['name'],
      number: account['number'],
      type: account['type'],
      transactions: account['transactions_id'].length,
      value: account['value'],
    };
    new_accounts.push(new_account);
  }
  return new_accounts;
}
function fire_getProjects() {
  const firestore = fire_get_firestore_();
  const projects = firestore.getDocuments('Projects');
  const new_projects = [];

  // FORMAT RECIVING DATA TO OBJECT ARRAY WITH MIN INFO
  for (let i = 0; i < projects.length; i++) {
    const project = projects[i]['fields'];
    const new_project = {
      id: project['id'],
      title: project['title'],
      description: project['description'],
      department: project['department'],
      type: project['type'],
      project_id: project['project_id'],
      intervention_id: project['intervention_id'],
    };
    new_projects.push(new_project);
  }
  return new_projects;
}
function fire_getMovements() {
  const firestore = fire_get_firestore_();
  const movs = firestore.getDocuments('Accounting');
  const new_movs = [];

  // FORMAT RECIVING DATA TO OBJECT ARRAY WITH MIN INFO
  for (let i = 0; i < movs.length; i++) {
    const mov = movs[i]['fields'];
    const new_mov = {
      id: mov['id'],
      concept: mov['concept'],
      creation_date: new Date(mov['creation_date']),
      creator_user: mov['creator_user'],
      execution_date: new Date(mov['execution_date']),
      user_in_charge: mov['user_in_charge'],
      type: mov['type'],
      amount: mov['amount'],
      vat: mov['vat'] / 100,
      code: mov['code'],
      origin: mov['origin'],
      place: mov['place'],
      project: mov['project'],
      intervention: mov['intervention'] || null,
      phase: mov['phase'] || null,
      account_id: mov['account_id'],
      target_id: mov['target_id'],
      image: null,
    };
    if (mov['images'].length > 0) {
      new_mov.image = mov['images']['0']['download_url'];
    }
    new_movs.push(new_mov);
  }
  return new_movs;
}
function fire_writeDataFromJson(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  datas,
  rows_above: number = 1,
  colums_before: number = 0
) {
  const row_offset = 1;

  // CLEAR THE HOLE SHEET BEFORE WRITING
  //sheet.clear();

  for (let i = 0; i < datas.length; i++) {
    let j = colums_before + 1;
    for (const key in datas[i]) {
      if (datas[i].hasOwnProperty(key)) {
        sheet
          .getRange(i + rows_above + row_offset, j + colums_before)
          .setValue(datas[i][key]);
        j++;
      }
    }
  }
}

//#endregion
