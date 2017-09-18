function sendEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var countsSheet = ss.getSheets()[0];
  var dataRange = countsSheet.getRange(2, 1, 8, 2);

  var ListsSheet = ss.getSheets()[1];
  var EAClientUpdates = ListsSheet.getRange((2, 1, ListsSheet.getMaxRows() - 1, 1)).getValue();
  var EAMigrations = ListsSheet.getRange((2, 3, ListsSheet.getMaxRows() - 1, 1)).getValue();
  var DemosOffered = ListsSheet.getRange((3, 4, ListsSheet.getMaxRows() - 1, 1)).getValue();
  var DemosScheduled = ListsSheet.getRange((3, 7, ListsSheet.getMaxRows() - 1, 1)).getValue();
  var MigrationsScheduled = ListsSheet.getRange((3, 8, ListsSheet.getMaxRows() - 1, 1)).getValue();



  // Create one JavaScript object per row of data.
  objects = getRowsData(dataSheet, dataRange);

  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var i = 0; i < objects.length; ++i) {
    // Get a row object
    var rowData = objects[i];

    // Generate a personalized email.
    // Given a template string, replace markers (for instance ${"First Name"}) with
    // the corresponding value in a row object (for instance rowData.firstName).
    var emailText = fillInTemplateFromObject(emailTemplate, rowData);
    var emailSubject = "[DRAFT] Migrations and Conversions Queue";

    MailApp.sendEmail(rowData.emailAddress, emailSubject, emailText);
  } 
}

// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
  var email = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);

  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  for (var i = 0; i < templateVars.length; ++i) {
    // normalizeHeader ignores ${"} so we can call it directly here.
    var variableData = data[normalizeHeader(templateVars[i])];
    email = email.replace(templateVars[i], variableData || "");
  }

  return email;
}
