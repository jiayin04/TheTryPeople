// TASK REMINDER IN CALENDAR
// Function to set task reminders based on the start date
function initializeTaskManagement() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Task Allocation');
  var data = sheet.getDataRange().getValues();
  var calendar = CalendarApp.getDefaultCalendar();

  for (var i = 1; i < data.length; i++) {
    var task = data[i][0];
    var assignedTo = data[i][1];
    var startDate = data[i][3];
    var endDate = data[i][4];

    // Check if it is a valid task
    if (task && startDate && endDate) {
      var eventDate = new Date(startDate);
      var event = calendar.createEvent(task, eventDate, eventDate)
        .setDescription('Task assigned to ' + assignedTo + '\nExpected End Date: ' + endDate);

      // Set reminder 1 day before the task starts
      event.addPopupReminder(24 * 60);
    }
  }
  SpreadsheetApp.getUi().alert('Added successful');
}

// UPDATE PROGRESS LOG BY EMPLOYEES
// Function to log progress updates (Employess only)
function handleProgressLogging(formData) {
  var task = formData.task;
  var assignedTo = formData.assignedTo;
  var progress = formData.progress;
  var desc = formData.desc;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Task Allocation');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == task && data[i][1] == assignedTo) {
      sheet.getRange(i + 1, 3).setValue(progress / 100); // Divide by 100 to convert to percentage
      sheet.getRange(i + 1, 8).setValue(desc);
      break;
    }
  }

  return 'Progress logged successfully!';
}

// CREATE A WEB APP
function doGet(e) {
  // return HtmlService.createHtmlOutputFromFile('ProgressDialog');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Task Allocation');
  var data = sheet.getDataRange().getValues();
  var names = [];
  var tasksMap = {};

  for (var i = 3; i < data.length; i++) { //starting from the row 3
    var name = data[i][1].trim(); // remove the one with empty string
    if (name && names.indexOf(name) === -1) {
      names.push(name);
    }

    var task = data[i][0].trim();
    if (task) {
      if (!tasksMap[name]) {
        tasksMap[name] = [];
      }
      tasksMap[name].push(task);
    }
  }

  var template = HtmlService.createTemplateFromFile('ProgressDialog');
  template.names = names;
  template.tasksMap = JSON.stringify(tasksMap); //stringtify it when passing to the script
  return template.evaluate();
}

// Function to open the web app URL
function openWebApp() {
  var html = HtmlService.createHtmlOutput('<html><script>'
    + 'window.open("https://script.google.com/macros/s/AKfycbxEhLrBqy7wr-LHTW7xhVCLnlfXY9MbHook2BqKbUdVSJEYaKJ10BujazALV6IwUCL1Ng/exec");'
    + 'google.script.host.close();'
    + '</script></html>')
    .setWidth(100)
    .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Web App...');
}

// GENERATION OF REPORT
// Function to generate a real-time progress report in Google Docs
function generateProgressReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Task Allocation');
  var data = sheet.getDataRange().getValues();
  var doc = DocumentApp.create('Task Allocation Report');
  var body = doc.getBody();

  // Adding the title with Heading 1 format
  body.appendParagraph("Task Allocation Report\n")
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setUnderline(true)
    .setBold(true);

  // Summary variables
  var totalTasks = data.length; // Adjusted for zero-based index
  var completedTasks = 0;
  var ongoingTasks = 0;
  var pendingTasks = 0;

  // Calculate the summary metrics
  for (var i = 0; i < data.length; i++) {
    var progress = data[i][2].toString();
    if (progress == "100%") {
      completedTasks++;
    } else if (progress > "0%") {
      ongoingTasks++;
    } else {
      pendingTasks++;
    }
  }

  // Adding a summary section
  body.appendParagraph("Summary\n")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
  // .setUnderline(true)
  // .setBold(true);

  var summaryTable = body.appendTable();

  row = summaryTable.appendTableRow();
  row.appendTableCell('Completed Tasks');
  row.appendTableCell(completedTasks.toString());

  row = summaryTable.appendTableRow();
  row.appendTableCell('Ongoing Tasks');
  row.appendTableCell(ongoingTasks.toString());

  row = summaryTable.appendTableRow();
  row.appendTableCell('Pending Tasks');
  row.appendTableCell(pendingTasks.toString());

  var row = summaryTable.appendTableRow();
  row.appendTableCell('Total Tasks');
  row.appendTableCell(totalTasks.toString());

  summaryTable.getRow(0).editAsText();

  // Adding the task details section
  body.appendParagraph("\nTask Details\n")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
  // .setUnderline(true)
  // .setBold(true);

  var table = body.appendTable();
  var headerRow = table.appendTableRow();
  headerRow.appendTableCell('Task');
  headerRow.appendTableCell('Assigned To');
  headerRow.appendTableCell('Progress');
  headerRow.appendTableCell('Start Date');
  headerRow.appendTableCell('End Date');
  headerRow.appendTableCell('To Be Started In');
  headerRow.appendTableCell('Duration');
  // headerRow.editAsText().setBold(true);

  // Iterate over the data and add it to the table
  for (var i = 3; i < data.length; i++) {
    var progress = data[i][2].toString();
    // var statusIcon = createStatusIcon(progress);

    var row = table.appendTableRow();
    row.appendTableCell(data[i][0].toString()); // Task
    row.appendTableCell(data[i][1].toString()); // Assigned To
    row.appendTableCell((data[i][2] * 100).toFixed(2) + '%'); // Progress
    row.appendTableCell(formatDate(data[i][3])); // Start Date
    row.appendTableCell(formatDate(data[i][4])); // End Date
    row.appendTableCell(data[i][5].toString() + ' days'); // To Be Started In
    row.appendTableCell(data[i][6].toString() + ' days'); // Duration

    // row.appendTableCell().appendParagraph(statusIcon); // Status Icon

    // var statusIconCell = row.appendTableCell();
    // statusIconCell.appendParagraph(statusIcon);
  }

  // Create a new sheet for the chart
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Check if the 'Task Chart' sheet exists
  var chartSheet = spreadsheet.getSheetByName('Task Chart');
  if (!chartSheet) {
    // Create a new sheet for the chart if it doesn't exist
    chartSheet = spreadsheet.insertSheet('Task Chart');
  } else {
    // Clear the existing sheet
    chartSheet.clear();
  }

  var chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange('A4:C8')) // Assuming your data range
    .setPosition(5, 5, 0, 0)
    .build();
  chartSheet.insertChart(chart);

  // Convert the chart to an image and add it to the document
  var charts = chartSheet.getCharts();
  var image = charts[0].getAs('image/png');
  body.appendParagraph("\nTask Progression").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendImage(image);

  // Gantt Chart for timeline display
  var imgURL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTiZImxE3AZCfUPmDE3gH8TBVxvzRak-XDqMhFsQ9-sdaBKNjeTsvr7tjKYdK6TqdpqhU5htEAKvoFu/pubchart?oid=135366856&format=image"; //published image
  var imageBlob = UrlFetchApp.fetch(imgURL).getBlob();
  body.appendParagraph("\nGantt Chart").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendImage(imageBlob);

  // Log the document URL and alert the user
  var docURL = doc.getUrl();
  Logger.log('Report generated: ' + docURL);
  var htmlDisplay = HtmlService.createHtmlOutput('<a href="' + docURL + '"target="_blank"> Click here to view the Progress Report </a>');
  SpreadsheetApp.getUi().showModalDialog(htmlDisplay, 'Progress Report Generated');
}

// Format the date
function formatDate(date) {
  var jsDate = new Date(date);
  return jsDate.toLocaleDateString();
}
