// Use this code for Google Docs, Forms, or new Sheets.
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .createMenu('Simple Outreach')
    .addItem('Parse Emails', 'showSidebar')
    .addItem('Add Key', 'openDialog')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Simple Outreach')
    .setWidth(400);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .showSidebar(html);
}

function openDialog() {
  var html = HtmlService.createTemplateFromFile('Dialog')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .showModalDialog(html, 'Please enter your API key');
}

function openDatetimeDialog() {
  var html = HtmlService.createTemplateFromFile('Datetime')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .showModalDialog(html, 'Please select schedule date');
}

function setEmailData(subject, message) {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('F2:F' + sheet.getLastRow()).setValue(subject);
  sheet.getRange('G2:G' + sheet.getLastRow()).setValue(message);
}

function setScheduledDate(date) {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('D2:D' + sheet.getLastRow()).setValue(date);
}

function saveApiKey(apiKey) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('apiKey', apiKey);
  var apiKey = scriptProperties.getProperty('apiKey');
  Browser.msgBox('Api key "' + apiKey + '" was seted');
  showSidebar();
}

function getApiKey() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var apiKey = scriptProperties.getProperty('apiKey');
  //  Browser.msgBox('returned "' + apiKey );
  return apiKey;
}

function sheetSetTitle() {
  var currentdate = new Date();
  var name = "SO: " + currentdate.getDate() + "/" + (currentdate.getMonth() + 1) + "/" + currentdate.getFullYear() + " @ " + currentdate.getHours() + ":" + currentdate.getMinutes() + ":" + currentdate.getSeconds();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.insertSheet(name, 0);
  ss.setActiveSheet(sheet);

  //var sheet = SpreadsheetApp.getActiveSheet();
  var firstTitle = sheet.getRange(1, 1).getValue();
  if (firstTitle != 'Domain') {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1).setValue('Domain').setBackground('#1155CC').setFontColor('#ffffff').setHorizontalAlignment("center").setBorder(null, null, true, null, false, false).setFontWeight("bold");
    sheet.getRange(1, 2).setValue('Url').setBackground('#1155CC').setFontColor('#ffffff').setHorizontalAlignment("center").setBorder(true, null, true, null, false, false).setFontWeight("bold");
    sheet.getRange(1, 3).setValue('Email Address').setBackground('#1155CC').setFontColor('#ffffff').setHorizontalAlignment("center").setBorder(true, null, true, null, false, false).setFontWeight("bold");
    sheet.getRange(1, 4).setValue('Scheduled Date').setBackground('#1155CC').setFontColor('#ffffff').setHorizontalAlignment("center").setBorder(true, null, true, null, false, false).setFontWeight("bold");
    sheet.getRange(1, 5).setValue('Mail Merge Status').setBackground('#1155CC').setFontColor('#ffffff').setHorizontalAlignment("center").setBorder(true, null, true, null, false, false).setFontWeight("bold");
    sheet.getRange(1, 6).setValue('Mail Subject').setBackground('#1155CC').setFontColor('#ffffff').setHorizontalAlignment("center").setBorder(true, null, true, null, false, false).setFontWeight("bold");
    sheet.getRange(1, 7).setValue('Mail Body').setBackground('#1155CC').setFontColor('#ffffff').setHorizontalAlignment("center").setBorder(true, null, true, null, false, false).setFontWeight("bold");
  }
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
  sheet.autoResizeColumn(3);
  sheet.autoResizeColumn(4);
  sheet.autoResizeColumn(5);
  sheet.autoResizeColumn(6);
  sheet.autoResizeColumn(7);
}

function schedule() {
  sendScheduledEmails();
}

function sendScheduledEmails() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  if (sheets.length < 1) {
    return;
  }

  for (key in sheets) {
    var sheet = sheets[key];
    Browser.msgBox(sheet.getName());
    if (!sheet.getLastRow()) {
      continue;
    }
    var startRow = 2; // First row of data to process
    var numRows = sheet.getLastRow() - 1; // Number of rows to process
    if (!numRows) {
      continue;
    }
    var numCols = sheet.getLastColumn();
    // Fetch the range of cells A2:B3
    var dataRange = sheet.getRange(startRow, 1, numRows, numCols);
    // Fetch values for each row in the Range.
    var variablesIndex = getCustomVariablesIndex(sheet);
    var data = dataRange.getValues();
    //Browser.msgBox(data);

    for (var i = 0; i < data.length; ++i) {
      var row = data[i];
      if (typeof row[4] == "undefined") {
        continue;
      }
      if (row[4] != 'SCHEDULED') {
        continue;
      }
      var emailAddress = row[2];
      if (!Date.parse(row[3])) {
        sheet.getRange(startRow + i, 5).setValue('WRONG_DATE');
        continue;
      }
      if (!validateEmail(emailAddress)) {
        sheet.getRange(startRow + i, 5).setValue('WRONG_EMAIL');
        continue;
      }
      var currentDate = new Date();
      var scheduledDate = new Date(row[3]);
      if (
        currentDate.getFullYear() == scheduledDate.getFullYear() &&
        currentDate.getMonth() == scheduledDate.getMonth() &&
        currentDate.getDate() == scheduledDate.getDate() &&
        currentDate.getHours() == scheduledDate.getHours()
      ) {
        sendCustomEmail(emailAddress, row[5], row[6], row, variablesIndex);
        sheet.getRange(startRow + i, 5).setValue('SENT');
        SpreadsheetApp.flush();
      }

    }
  }
}



function findEmailAction(apiKey, listDomains) {
  Browser.msgBox('Begining Of Search');
  sheetSetTitle();

  var sheet = SpreadsheetApp.getActiveSheet();
  for (domain in listDomains) {
    var reqUrl = 'http://simpleoutreach.com/api/email/' + encodeURIComponent(listDomains[domain]) + '?token=' + apiKey;
    // Browser.msgBox(reqUrl);
    try {
      var response = UrlFetchApp.fetch(reqUrl);
    } catch (e) {
      Browser.msgBox('Wrong API key or inactive user.');
      Browser.msgBox(e);
      Logger.log(e);
      return false;
    }

    var json = response.getContentText();
    //Browser.msgBox(json);
    var data = JSON.parse(json);
    //Logger.log(data);
    if (!data.data.emails.length) {
      data.data.emails.push('no data');
    }
    for (email in data.data.emails) {
      //Browser.msgBox(data.data.url +' '+ data.data.emails[email]);
      // (new Date).format("dd-mm-yyyy")
      sheet.appendRow([data.data.domain, data.data.url, data.data.emails[email], '', 'PENDING', '']);
    }
  }
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
  sheet.autoResizeColumn(3);
  sheet.autoResizeColumn(4);
  sheet.autoResizeColumn(5);
  sheet.autoResizeColumn(6);
  sheet.autoResizeColumn(7);

  var column = sheet.getRange("D2:D");
  column.setNumberFormat("yyyy-mm-dd");

  var cell = SpreadsheetApp.getActiveSheet().getRange("F2");
  cell.setValue(new Date());
  cell.setNumberFormat('yyyy-mm-dd');
  cell.setValue('');

  Browser.msgBox('End Of Search');
}

function sendCustomEmail(emailAddress, subject, message, data, variablesIndex) {
  for (var key in variablesIndex) {
    message = message.replace('{' + key + '}', data[variablesIndex[key]]);
  }

  for (var key in variablesIndex) {
    subject = subject.replace('{' + key + '}', data[variablesIndex[key]]);
  }

  MailApp.sendEmail(emailAddress, subject, message);
}

function sendTestEmailAction(subject, message) {
  var emailAddress = Session.getEffectiveUser().getEmail();

  var sheet = SpreadsheetApp.getActiveSheet();
  var variablesIndex = getCustomVariablesIndex(sheet);

  var numCols = sheet.getLastColumn();
  var dataRange = sheet.getRange(2, 1, 1, numCols);
  var data = dataRange.getValues();
  var row = data[0];
  sendCustomEmail(emailAddress, subject, message, row, variablesIndex);

  Browser.msgBox('Email was sent to: ' + emailAddress);
}



function validateEmail(email) {
  var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return re.test(email);
}

function getCustomVariablesIndex(sheet) {
  var numCols = sheet.getLastColumn();
  var values = sheet.getRange(1, 1, 1, numCols).getValues()[0];
  var result = {};
  for (var index in values) {
    var key = values[index];
    result[values[index]] = index;
  }
  return result;
}

function sendEmailsAction(subject, message) {
  Browser.msgBox('Begining Of Sending');
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = sheet.getLastRow() - 1; // Number of rows to process
  var numCols = sheet.getLastColumn();
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, numCols);
  // Fetch values for each row in the Range.
  var variablesIndex = getCustomVariablesIndex(sheet);
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[2];
    if (Date.parse(row[3])) {
      sheet.getRange(startRow + i, 5).setValue('SCHEDULED');
      sheet.getRange(startRow + i, 6).setValue(subject);
      sheet.getRange(startRow + i, 7).setValue(message);
      continue;
    }
    if (!validateEmail(emailAddress)) {
      sheet.getRange(startRow + i, 5).setValue('WRONG_EMAIL');
      continue;
    }
    //Browser.msgBox(emailAddress +' | '+ subject +' | '+ message +' | '+ Date.parse(row[3]));
    sendCustomEmail(emailAddress, subject, message, row, variablesIndex);
    sheet.getRange(startRow + i, 5).setValue('SENT');
    SpreadsheetApp.flush();
  }

  Browser.msgBox('End Of Sending');
}

function alertAction(message) {
  Browser.msgBox(message);
}