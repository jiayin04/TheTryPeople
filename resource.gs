function generateResourceForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Resources');
  // var data = sheet.getDataRange().getValues();
  
  var form;
  var existingForms = FormApp.openByUrl(ss.getFormUrl());
  
  if (existingForms) {
    Logger.log('Form already exists. Form URL: ' + existingForms.getPublishedUrl());
    var htmlDisplay = HtmlService.createHtmlOutput('<a href="' + existingForms.getPublishedUrl() + '" target="_blank">Click here to view the form.</a>');
    SpreadsheetApp.getUi().showModalDialog(htmlDisplay, 'Form already exists');
    return;
  } else {
    form = FormApp.create('Resource Request Form');
  }

  form.addTextItem().setTitle('Employee Name');
  form.addTextItem().setTitle('Resource ID');
  form.addTextItem().setTitle('Quantity Requested').setValidation(FormApp.createTextValidation().requireNumber().build());
  
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  SpreadsheetApp.getUi().alert('Resource Request Form generated and linked to this spreadsheet.');
}


function processRequests() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resourceSheet = ss.getSheetByName('Resources');
  var requestSheet = ss.getSheetByName('Requests');
  var formResponsesSheet = ss.getSheetByName('Form Responses 1');
  
  var resourceData = resourceSheet.getDataRange().getValues();
  var requestData = requestSheet.getDataRange().getValues();
  var formResponsesData = formResponsesSheet.getDataRange().getValues();

  for (var i = 1; i < formResponsesData.length; i++) {
    var timestamp = formResponsesData[i][0];
    var employeeName = formResponsesData[i][1];
    var resourceID = formResponsesData[i][2];
    var quantityRequested = formResponsesData[i][3];
    var status = formResponsesData[i][4];
    
    if (status !== 'Processed') {
      for (var j = 1; j < resourceData.length; j++) {
        if (resourceData[j][0] == resourceID) {
          var availableQuantity = resourceData[j][3];
          if (availableQuantity >= quantityRequested) {
            resourceSheet.getRange(j + 1, 4).setValue(availableQuantity - quantityRequested);
            formResponsesSheet.getRange(i + 1, 5).setValue('Processed');
            requestSheet.appendRow([timestamp, employeeName, resourceID, quantityRequested, 'Approved']);
          } else {
            requestSheet.appendRow([timestamp, employeeName, resourceID, quantityRequested, 'Denied']);
          }
          break;
        }
      }
    }
  }
}

function requestData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formResponsesSheet = ss.getSheetByName('Form Responses 3');
  var requestData = formResponsesSheet.getDataRange().getValues();
  
  return requestData;
}
