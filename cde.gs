function createPrintableSheetWithFormulas() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getActiveSheet();
  var sourceSheetName = sourceSheet.getName();
  var dataRange = sourceSheet.getDataRange();
  var values = dataRange.getValues();
  var formulas = dataRange.getFormulas();
  var hiddenRows = [];
  var visibleRowsMap = {};

  // Find hidden rows and create a map for visible rows
  for (var i = 0; i < values.length; i++) {
    if (sourceSheet.isRowHiddenByUser(i + 1)) {
      hiddenRows.push(i + 1);
    } else {
      visibleRowsMap[i + 1] = Object.keys(visibleRowsMap).length + 1;
    }
  }

  // Create a new sheet with the same format
  var newSheetName = 'Printable ' + sourceSheetName;
  var newSheet = spreadsheet.duplicateSheet(sourceSheet.getId());
  newSheet.setName(newSheetName);

  // Clear the data in the new sheet
  newSheet.getDataRange().clearContent();

  // Copy visible rows to the new sheet and adjust formulas
  var rowOffset = 0;
  for (var i = 0; i < values.length; i++) {
    if (!hiddenRows.includes(i + 1)) {
      // Copy row data
      sourceSheet.getRange(i + 1, 1, 1, values[i].length).copyTo(newSheet.getRange(visibleRowsMap[i + 1], 1), { formatOnly: false });

      // Adjust formulas
      for (var col = 1; col <= values[i].length; col++) {
        var formula = formulas[i][col - 1];
        if (formula) {
          var adjustedFormula = formula.replace(/R(\d+)C(\d+)/g, (match, row, col) => {
            return `R${visibleRowsMap[row] || row}C${col}`;
          });
          newSheet.getRange(visibleRowsMap[i + 1], col).setFormula(adjustedFormula);
        }
      }
    }
  }

  Logger.log('Hidden Rows: ' + hiddenRows.join(', '));
}


function createPrintableSheetWithValues() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getActiveSheet();
  var sourceSheetName = sourceSheet.getName();
  var dataRange = sourceSheet.getDataRange();
  var values = dataRange.getValues(); // Get all values, including calculated results
  var hiddenRows = [];
  var visibleRowsMap = {};

  // Find hidden rows and create a map for visible rows
  for (var i = 0; i < values.length; i++) {
    if (sourceSheet.isRowHiddenByUser(i + 1)) {
      hiddenRows.push(i + 1);
    } else {
      visibleRowsMap[i + 1] = Object.keys(visibleRowsMap).length + 1;
    }
  }

  // Create a new sheet with the same format
  var newSheetName = 'Printable ' + sourceSheetName;
  var newSheet = spreadsheet.duplicateSheet(sourceSheet.getId());
  newSheet.setName(newSheetName);

  // Clear the data in the new sheet
  newSheet.getDataRange().clearContent();

  // Copy visible rows to the new sheet, replacing formulas with their values
  var rowOffset = 0;
  for (var i = 0; i < values.length; i++) {
    if (!hiddenRows.includes(i + 1)) {
      // Copy row data (values only)
      newSheet.getRange(visibleRowsMap[i + 1], 1, 1, values[i].length).setValues([values[i]]);
    }
  }

  Logger.log('Hidden Rows: ' + hiddenRows.join(', '));
}
