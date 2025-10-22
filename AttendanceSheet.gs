/**
* @OnlyCurrentDoc
*/

let sheet = SpreadsheetApp.getActiveSpreadsheet();
var generateSheet = sheet.getSheetByName('Generate');
var attendanceSheet = sheet.getSheetByName(generateSheet.getRange(7, 10).getValue().toString());
let attendanceSheetNameCell = generateSheet.getRange(7, 10);
let attendanceSheetName = attendanceSheetNameCell.getValue().toString().trim();

function generate() {
  var daysOfWeek = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

  let startDate = generateSheet.getRange(7, 3).getValue();
  var startDayOfWeek = getDayOfWeekFromDate(startDate);
  var startIndex = daysOfWeek.indexOf(startDayOfWeek);

  let endDate = generateSheet.getRange(8, 3).getValue();

  // Check if the attendance sheet name cell is empty
  if (!attendanceSheetName) {
    SpreadsheetApp.getUi().alert('Warning: The cell for the sheet name is empty. Please enter a valid sheet name in step 3 and try again.');
    return; 
  }

  // Check if the attendance sheet exists
  if (!attendanceSheet) {
    SpreadsheetApp.getUi().alert(`Warning: The sheet "${attendanceSheetName}" does not exist. Please check if you made the sheet and try again.`);
    return;
  }

  if (endDate < startDate) {
    SpreadsheetApp.getUi().alert('Warning: The end date is earlier than the start date. Please correct the dates.');
    return; 
  }

  var sun = generateSheet.getRange(11, 3).getValue();
  var mon = generateSheet.getRange(12, 3).getValue();
  var tue = generateSheet.getRange(13, 3).getValue();
  var wed = generateSheet.getRange(14, 3).getValue();
  var thu = generateSheet.getRange(15, 3).getValue();
  var fri = generateSheet.getRange(16, 3).getValue();
  var sat = generateSheet.getRange(17, 3).getValue();

  const days = [
    { name: 'Sun', value: sun },
    { name: 'Mon', value: mon },
    { name: 'Tue', value: tue },
    { name: 'Wed', value: wed },
    { name: 'Thu', value: thu },
    { name: 'Fri', value: fri },
    { name: 'Sat', value: sat }
  ];

  //const trueVariables = days.filter(variable => variable.value === true);

  var daysToAdd = []; //find position on 0-6 of the days
  for (let i = 0; i < days.length; i++) {
    if (days[i].value === true) {
      daysToAdd.push(i); // Adding 1 to convert from zero-based index to one-based position
    }
  }

  var currentDate = new Date(startDate);
  var row = 6; // Row number where dates will be populated
  var column = 7; // Start from column G

  var dateFormatted = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'MM/dd');
  
  while (currentDate <= endDate) {
    for (var i = 0; i < daysOfWeek.length; i++) {
      var nextIndex = (startIndex + i) % 7;
      if (daysToAdd.includes(nextIndex)) {
        var daysUntilNext = (nextIndex >= startIndex) ? nextIndex - startIndex : 7 - startIndex + nextIndex;
        var nextDate = new Date(currentDate);
        nextDate.setDate(currentDate.getDate() + daysUntilNext);
        if (nextDate <= endDate && !isDateToSkip(generateSheet, nextDate)) {
          dateFormatted = Utilities.formatDate(nextDate, Session.getScriptTimeZone(), 'MM/dd');
          attendanceSheet.getRange(row, column).setValue(dateFormatted);
          attendanceSheet.getRange(row - 1, column).setValue(daysOfWeek[nextIndex]);
          column++; // Move to the next column
        }
      }
    }
    currentDate.setDate(currentDate.getDate() + 7); // Move to the next week
  }
}

function getDayOfWeekFromDate(date) {
  var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'E');
  return formattedDate;
}

function isDateToSkip(sheet, date) {
  var range = sheet.getRange('B20:C30');
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] instanceof Date && values[i][j].getTime() === date.getTime()) {
        return true;
      }
    }
  }
  return false;
}

function fixBorders() {
  var lastRow = attendanceSheet.getMaxRows();
  var lastColumn = attendanceSheet.getMaxColumns();
  
  var range = attendanceSheet.getRange(5, 7, lastRow - 4, lastColumn - 6); // Starting from row 5, all columns
  
  range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  range = attendanceSheet.getRange(2, 27, 2, lastColumn - 26);

  range.setBorder(true, null, true, null, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
}

function clearSheet() {
  var rangeToClear1 = generateSheet.getRange('C7:C8'); 
  var rangeToClear2 = generateSheet.getRange('B20:C30');
  var rangeToClear3 = generateSheet.getRange('J7');

  var checkboxRange = generateSheet.getRange('C11:C17'); // Change the range as needed

  var values = checkboxRange.getValues();

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] === true) {
        values[i][j] = false;
      }
    }
  }

  checkboxRange.setValues(values);

  // Clear the content and formatting from the specified ranges
  rangeToClear1.clear({contentsOnly: true});
  rangeToClear2.clear({contentsOnly: true});
  rangeToClear3.clear({contentsOnly: true});
}

