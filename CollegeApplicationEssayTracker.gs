/**
* @OnlyCurrentDoc
*/

let sheet = SpreadsheetApp.getActiveSpreadsheet(); //don't edit this line
let checklistSheet = sheet.getSheetByName('Checklist'); //if change name, change stuff in the '' --> sheet.getSheetByName('here');
let essayTypeSheet = sheet.getSheetByName('Essay Types');
let count = 0;

function updateButton() {
  var sheetName = sheet.getActiveSheet().getName();

  if (checkAlreadyThereChecklist(sheetName) == false) {
    var row = 6; //row on checklist sheet
    var cell = checklistSheet.getRange(row, 2);
    var currentValue;
    var dateValue = sheet.getActiveSheet().getRange(3,5).getValue();

    while ((currentValue = cell.getValue()) !== "") {
      row++;
      cell = checklistSheet.getRange(row, 2);
    }

    checklistSheet.getRange(row, 2).setValue(sheetName);
    checklistSheet.getRange(row, 3).setValue(dateValue);
  }
  if (checkAlreadyThereChecklist(sheetName) == true) {
    var dateValue = sheet.getActiveSheet().getRange(3,5).getValue();
    var data = checklistSheet.getRange("B6:B").getValues(); // Assuming the data is in the second column (B)

    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == sheetName) { // Assuming the adjacent column is column C (3)
        checklistSheet.getRange(i + 6, 3).setValue(dateValue); // Set the value 2 in the adjacent column
      }
    }
  }

  var sheetRow = 6;

  var nextCell = sheet.getActiveSheet().getRange(sheetRow + 1, 4).getValue();
  while (nextCell !== "") {
    sheetRow++;
    nextCell = sheet.getActiveSheet().getRange(sheetRow + 1, 4).getValue();
    checkEssayType(sheetName);
    count++;
  }
  if (nextCell == ""){
    checkEssayType(sheetName);
  }
}

function checkAlreadyThereChecklist(sheetName) { //for checklist sheet
  var lastRow = checklistSheet.getLastRow();
  var numRows = lastRow - 6 + 1;
  var range = checklistSheet.getRange(6,2,numRows);
  var values = range.getValues();
  var contains = false;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == sheetName) {
      contains = true;
      return contains;
    }
  }

  return contains;
}

function checkAlreadyThere(column, sheetName) { //for essatTypeSheet
  var lastRow = essayTypeSheet.getLastRow();
  var numRows;
  if (lastRow == 5) {
    numRows = 1;
  }
  else {numRows = lastRow - 6 + 1;}
  var range = essayTypeSheet.getRange(6,column,numRows);
  var values = range.getValues();
  var contains = false;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == sheetName) {
      contains = true;
      return contains;
    }
  }

  return contains;
}

function checkEssayType(sheetName) {
  var essayRow = 6;
  var essayColumn = 2;
  var essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
  var sheetRow = 6 + count;
  var sheetCell = sheet.getActiveSheet().getRange(sheetRow, 4);
  var currentValueEssayType;

  if (sheetCell.getValue() == essayTypeSheet.getRange(5,2).getValue()) {
    essayColumn = 2;
    if (checkAlreadyThere(essayColumn, sheetName) == false) {
      while ((currentValueEssayType = essayCell.getValue()) !== "") {
        essayRow++;
        essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
      }
      essayTypeSheet.getRange(essayRow,essayColumn).setValue(sheetName);
    }
  }

  else if (sheetCell.getValue() == essayTypeSheet.getRange(5,3).getValue()) {
    essayColumn = 3;
    essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
    if (checkAlreadyThere(essayColumn, sheetName) == false) {
      while ((currentValueEssayType = essayCell.getValue()) !== "") {
        essayRow++;
        essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
      }
      essayTypeSheet.getRange(essayRow,essayColumn).setValue(sheetName);
    }
  }

  else if (sheetCell.getValue() == essayTypeSheet.getRange(5,4).getValue()) {
    essayColumn = 4;
    essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
    if (checkAlreadyThere(essayColumn, sheetName) == false) {
      while ((currentValueEssayType = essayCell.getValue()) !== "") {
        essayRow++;
        essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
      }
      essayTypeSheet.getRange(essayRow,essayColumn).setValue(sheetName);
    }
  }

  else if (sheetCell.getValue() == essayTypeSheet.getRange(5,5).getValue()) {
    essayColumn = 5;
    essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
    if (checkAlreadyThere(essayColumn, sheetName) == false) {
      while ((currentValueEssayType = essayCell.getValue()) !== "") {
        essayRow++;
        essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
      }
      essayTypeSheet.getRange(essayRow,essayColumn).setValue(sheetName);
    }
  }

  else if (sheetCell.getValue() == essayTypeSheet.getRange(5,6).getValue()) {
    essayColumn = 6;
    essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
    if (checkAlreadyThere(essayColumn, sheetName) == false) {
      while ((currentValueEssayType = essayCell.getValue()) !== "") {
        essayRow++;
        essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
      }
      essayTypeSheet.getRange(essayRow,essayColumn).setValue(sheetName);
    }
  }

  else if (sheetCell.getValue() == essayTypeSheet.getRange(5,7).getValue()) {
    essayColumn = 7;
    essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
    if (checkAlreadyThere(essayColumn, sheetName) == false) {
      while ((currentValueEssayType = essayCell.getValue()) !== "") {
        essayRow++;
        essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
      }
      essayTypeSheet.getRange(essayRow,essayColumn).setValue(sheetName);
    }
  }

  else if (sheetCell.getValue() == essayTypeSheet.getRange(5,8).getValue()) {
    essayColumn = 8;
    essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
    if (checkAlreadyThere(essayColumn, sheetName) == false) {
      while ((currentValueEssayType = essayCell.getValue()) !== "") {
        essayRow++;
        essayCell = essayTypeSheet.getRange(essayRow,essayColumn);
      }
      essayTypeSheet.getRange(essayRow,essayColumn).setValue(sheetName);
    }
  }

  else {

  }
}

