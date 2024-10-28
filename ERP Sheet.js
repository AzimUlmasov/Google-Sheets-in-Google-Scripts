function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Convert Formula to Value')
    .addItem('Convert', 'convertFormulasToValues')
    .addToUi();
}

function onEdit(e) {
  // Check if event object and range property are defined
  if (e && e.range) {
    addTimestamp(e);

    var editedColumn = e.range.getColumn();
    var editedRow = e.range.getRow();
    var sheet = e.source.getSheetByName("Sells");

    // Check if the edited cell is in column A (1) and has data
    if (editedColumn === 1 && sheet.getRange(editedRow, 1).getValue() !== "") {
      var cell = sheet.getRange(editedRow, 2); // Corresponding cell in column B
      convertCellToValue(cell);
    }
  }
}

function addTimestamp(e){
  try {
    var startRow = 2;
    var targetColumn = 1;
    var ws = "Sells";
    var row = e.range.getRow();
    var col = e.range.getColumn();

    if (col === targetColumn && row >= startRow && e.source.getActiveSheet().getName() === ws) {
      var currentDate = new Date();
      e.source.getActiveSheet().getRange(row, 6).setValue(currentDate);
      if (e.source.getActiveSheet().getRange(row, 5).getValue() == "") {
        e.source.getActiveSheet().getRange(row, 5).setValue(currentDate);
      }
    }
  } catch (error) {
    console.error("Error in addTimestamp function: ", error);
  }
}


function convertCellToValue(cell) {
  Utilities.sleep(100); // Wait for 1 seconds
  var isCurrency = (cell.getDisplayValue().charAt(0) === "$");
  var value = cell.getValue(); // This evaluates the formula
  cell.setValue(value); // Set the value back into the cell (effectively removing the formula)

  // If the cell was displaying a currency symbol, reapply the currency format
  if (isCurrency) {
    cell.setNumberFormat("$#,##0.00");
  }
}

function convertFormulasToValues() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sells");
  var rangeA = sheet.getRange("A2:A");
  var rangeB = sheet.getRange("B2:B");
  var valuesA = rangeA.getValues();
  var valuesB = rangeB.getValues();

  for (var i = 0; i < valuesA.length; i++) {
    var cell = rangeB.getCell(i + 1, 1);
    var value = cell.getFormula(); // Get the formula of the cell
    
    if (valuesA[i][0] != "" && value.charAt(0) === "=") {
      var isCurrency = (cell.getDisplayValue().charAt(0) === "$");
      var evaluatedValue = cell.getValue(); // This evaluates the formula
      
      cell.setValue(evaluatedValue); // Set the value back into the cell (effectively removing the formula)

      // If the cell was displaying a currency symbol, reapply the currency format
      if (isCurrency) {
        cell.setNumberFormat("$#,##0.00");
      }
    }
  }
}

// Custom Configuration
function applyFormula() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sells");
  var range = sheet.getRange("D2:D400");

  // Get the values in the D2:D range Стоимость 
  var values = range.getValues();

  var currency = sheet.getRange("H1").getValues()[0][0];

  // Loop through each row in the range
  for (var i = 0; i < range.getNumRows(); i++) {
    var cell = range.getCell(i + 1, 1);
    var value = cell.getValue();
    var displayValue = cell.getDisplayValue();
    var row = i + 2; // Adjust row number to match spreadsheet row number
    if (sheet.getRange("G" + row).getValue() !== '') {
      continue
    } 

    if (displayValue.charAt(0) === "$") {
      // Remove "$" and "," from the string, convert to number, and multiply by hValue
      sheet.getRange("G" + row).setValue(value * currency);
    } else {
      // Set the value in the cell to be the same as D2
      sheet.getRange("G" + row).setValue(value);
    }
  }
}
