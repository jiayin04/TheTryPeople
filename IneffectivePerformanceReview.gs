function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu')
    .addItem('Generate Performance Report', 'generatePerformanceReport')
    .addItem('Create Employee Feedback Form', 'createEmployeeFeedbackForm')
    .addToUi();
}

function createEmployeeFeedbackForm() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var formId = scriptProperties.getProperty('EMPLOYEE_FEEDBACK_FORM_ID');

  // Check if the form already exists
  if (formId) {
    var form = FormApp.openById(formId);
    Logger.log('Form already exists. Form URL: ' + form.getPublishedUrl());
    var htmlDisplay = HtmlService.createHtmlOutput('<a href="' + form.getPublishedUrl() + '" target="_blank">Click here to view the form.</a>');
    SpreadsheetApp.getUi().showModalDialog(htmlDisplay, 'Form already exists');
    return;
  }

  // Create a new form
  form = FormApp.create('Employee Feedback Form')
    .setDescription('Please provide your feedback and set personal goals.');

  form.addTextItem().setTitle('Employee Name').setRequired(true);

  // Add multiple-choice item for Position
  form.addMultipleChoiceItem()
    .setTitle('Position')
    .setRequired(true)
    .setChoiceValues(['Developer', 'Designer', 'Manager', 'QA Engineer', 'HR']);

  // Add Likert scale for Performance Rating
  form.addScaleItem()
    .setTitle('Performance Rating')
    .setBounds(1, 5) // Likert scale from 1 to 5
    .setLabels('Poor', 'Excellent') // Corrected
    .setRequired(true);

  form.addParagraphTextItem().setTitle('Feedback').setRequired(true);
  form.addParagraphTextItem().setTitle('Goals').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);

  scriptProperties.setProperty('EMPLOYEE_FEEDBACK_FORM_ID', form.getId());
  Logger.log('Form created. Form URL: ' + form.getPublishedUrl());
  var htmlDisplay = HtmlService.createHtmlOutput('<a href="' + form.getPublishedUrl() + '" target="_blank">Click here to view the form.</a>');
  SpreadsheetApp.getUi().showModalDialog(htmlDisplay, 'Form created');
}

function generatePerformanceReport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Form Responses 9');

  // Check if the sheet exists
  if (!sheet) {
    Logger.log('Sheet "Form Responses 1" not found');
    SpreadsheetApp.getUi().alert('Sheet "Form Responses 1" not found');
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
    .setYAxisTitle('Rating')
    .setDimensions(600, 400)
    .setDataTable(dataTable)
    .build();

  return chart;
}
