function createPrintableSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getActiveSheet();
  var sourceSheetName = sourceSheet.getName();
  var dataRange = sourceSheet.getDataRange();
  var values = dataRange.getValues();
  var hiddenRows = [];

  // Find hidden rows
  for (var i = 0; i < values.length; i++) {
    if (sourceSheet.isRowHiddenByUser(i + 1)) {
      hiddenRows.push(i + 1);
    }
  }

  // Create a new sheet with the same format
  var newSheetName = 'Printable ' + sourceSheetName;
  var newSheet = spreadsheet.duplicateSheet(sourceSheet.getId());
  newSheet.setName(newSheetName);

  // Clear the data in the new sheet
  newSheet.getDataRange().clearContent();

  // Copy visible rows to the new sheet
  var rowOffset = 0;
  for (var i = 0; i < values.length; i++) {
    if (!hiddenRows.includes(i + 1)) {
      sourceSheet.getRange(i + 1, 1, 1, values[i].length).copyTo(newSheet.getRange(i + 1 - rowOffset, 1), { formatOnly: false });
    } else {
      rowOffset++;
    }
  }

  Logger.log('Hidden Rows: ' + hiddenRows.join(', '));
}
