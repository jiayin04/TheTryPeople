function createEmployeeFeedbackForm() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var formId = scriptProperties.getProperty('EMPLOYEE_FEEDBACK_FORM_ID');

  // Check if the form already exists
  if (formId) {
    try {
      var form = FormApp.openById(formId);
      Logger.log('Form already exists. Form URL: ' + form.getPublishedUrl());
      var htmlDisplay = HtmlService.createHtmlOutput('<a href="' + form.getPublishedUrl() + '" target="_blank">Click here to view the form.</a>');
      SpreadsheetApp.getUi().showModalDialog(htmlDisplay, 'Form already exists');
      return;
    } catch (e) {
      Logger.log('Error opening form by ID: ' + e.message);
      scriptProperties.deleteProperty('EMPLOYEE_FEEDBACK_FORM_ID'); // Remove the invalid ID
    }
  }

  // Create a new form
  var form = FormApp.create('Employee Feedback Form')
    .setDescription('Please provide your feedback and set personal goals.');

  form.addTextItem().setTitle('Employee Name').setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('Position')
    .setRequired(true)
    .setChoiceValues(['Developer', 'Designer', 'Manager', 'QA Engineer', 'HR']);

  form.addScaleItem()
    .setTitle('Performance Rating')
    .setBounds(1, 5)
    .setLabels('Poor', 'Excellent')
    .setRequired(true);

  form.addParagraphTextItem().setTitle('Feedback').setRequired(true);
  form.addParagraphTextItem().setTitle('Goals').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);

  // Create a new Google Sheet for responses
  var timestamp = new Date().getTime();
  var sheetName = 'Employee Feedback Form';
  var spreadsheet = SpreadsheetApp.create('Employee Feedback Responses ' + timestamp);
  var sheet = spreadsheet.getSheets()[0];
  sheet.setName(sheetName);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

  scriptProperties.setProperty('EMPLOYEE_FEEDBACK_FORM_ID', form.getId());
  Logger.log('Form created. Form URL: ' + form.getPublishedUrl());
  var htmlDisplay = HtmlService.createHtmlOutput('<a href="' + form.getPublishedUrl() + '" target="_blank">Click here to view the form.</a>');
  SpreadsheetApp.getUi().showModalDialog(htmlDisplay, 'Form created');
}

function generatePerformanceReport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  
  // Log the names of all sheets for debugging purposes
  Logger.log('Sheets in the spreadsheet:');
  sheets.forEach(function(sheet) {
    Logger.log(sheet.getName());
  });
  
  var sheet = null;
  var identifierPrefix = 'Employee Feedback Form';

  // Find the sheet with the unique identifier
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    if (sheetName.startsWith(identifierPrefix)) {
      sheet = sheets[i];
      break;
    }
  }

  // Check if the sheet was found
  if (!sheet) {
    Logger.log('No suitable sheet found in the spreadsheet');
    SpreadsheetApp.getUi().alert('No suitable sheet found in the spreadsheet');
    return;
  }

  var data = sheet.getDataRange().getValues();
  Logger.log('Data retrieved from sheet: ' + data.length + ' rows');

  // Create and format the document
  var doc = DocumentApp.create('Performance Report');
  var body = doc.getBody();

  var title = body.appendParagraph("Performance Report\n")
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setUnderline(true)
    .setBold(true);

  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // Collect rating data for chart creation
  var ratings = [];
  for (var i = 1; i < data.length; i++) {
    ratings.push(data[i][3]);
  }

  // Create and insert a chart into the document
  var chart = createRatingChart(ratings);
  if (chart) {
    var chartBlob = chart.getAs('image/png');
    var paragraph = body.appendParagraph('');
    paragraph.appendInlineImage(chartBlob);
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('\n');
  }

  // Add employee details
  for (var i = 1; i < data.length; i++) {
    var employee = data[i][1];
    var position = data[i][2];
    var rating = data[i][3];
    var feedback = data[i][4];
    var goals = data[i][5];
    var email = data[i][6];

    body.appendParagraph(employee + " - " + position)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    var detailsTable = body.appendTable();
    detailsTable.setBorderWidth(1);
    detailsTable.setBorderColor('#000000');

    var row = detailsTable.appendTableRow();
    row.appendTableCell('Performance Rating').setBackgroundColor('#f0f0f0');
    row.appendTableCell(rating.toString()).setBackgroundColor('#ffffff');

    row = detailsTable.appendTableRow();
    row.appendTableCell('Feedback').setBackgroundColor('#f0f0f0');
    row.appendTableCell(feedback).setBackgroundColor('#ffffff');

    row = detailsTable.appendTableRow();
    row.appendTableCell('Goals').setBackgroundColor('#f0f0f0');
    row.appendTableCell(goals).setBackgroundColor('#ffffff');

    row = detailsTable.appendTableRow();
    row.appendTableCell('Email').setBackgroundColor('#f0f0f0');
    row.appendTableCell(email).setBackgroundColor('#ffffff');

    body.appendParagraph('\n');
  }

  var docUrl = doc.getUrl();
  Logger.log('Performance report generated: ' + docUrl);
  var htmlDisplay = HtmlService.createHtmlOutput('<a href="' + docUrl + '" target="_blank">Click here to view the Performance Report</a>');
  SpreadsheetApp.getUi().showModalDialog(htmlDisplay, 'Performance Report Generated');
}


function createRatingChart(ratings) {
  var dataTable = Charts.newDataTable();
  dataTable.addColumn(Charts.ColumnType.STRING, "Rating");
  dataTable.addColumn(Charts.ColumnType.NUMBER, "Rating Point Count");

  var ratingCounts = {};
  ratings.forEach(function (rating) {
    if (ratingCounts[rating]) {
      ratingCounts[rating]++;
    } else {
      ratingCounts[rating] = 1;
    }
  });

  for (var rating in ratingCounts) {
    dataTable.addRow([rating, ratingCounts[rating]]);
  }

  var chart = Charts.newBarChart()
    .setTitle('Performance Ratings Distribution')
    .setXAxisTitle('Rating Point')
    .setYAxisTitle('Count')
    .setDimensions(600, 400)
    .setDataTable(dataTable)
    .build();

  return chart;
}
