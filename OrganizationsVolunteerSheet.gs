/**
* @OnlyCurrentDoc
*/

let sheet = SpreadsheetApp.getActiveSpreadsheet();
var calculationSheet = sheet.getSheetByName('Calculations');
var hourLogSheet = sheet.getSheetByName('Volunteer Log'); //if you want to change the name of hour log sheet, replace the text in the single quotes. MAKE SURE THE SINGLE QUOTES ARE STILL THERE!
let numEvents = parseInt(calculationSheet.getRange(2, 13).getValue());
let numPeople = parseInt(calculationSheet.getRange(4, 13).getValue());
let numDates = parseInt(calculationSheet.getRange(6, 13).getValue());

//enter in hours button
function enterHours() {
  let sourceSheet = sheet.getSheetByName('Enter in Hours');

  var activeSheet = sheet.getActiveSheet();
  var secondSourceSheet = "Enter in Hours 2";

  var range = hourLogSheet.getDataRange();
  var values = range.getValues();

  //get the length of the amount of rows there are
  let rowLength = calculationSheet.getRange(2, 11).getValue();

  //get the row and column value to put in
  let calcRow = 2;
  let calcColumnName = 9;
  let calcColumnDate = 10;

  //get the value of the hours, source sheet
  let row = 13;
  let column = 8;

  //second sheet used, checks if sheet is active
  if (activeSheet.getName() == secondSourceSheet) {
    sourceSheet = sheet.getSheetByName('Enter in Hours 2');
    rowLength = calculationSheet.getRange(2, 17).getValue();
    calcColumnName = 15;
    calcColumnDate = 16;
  }

  //get the event and date from the dropdowns
  let event = sourceSheet.getRange(8, 3).getValue();
  let date = sourceSheet.getRange(7, 3).getValue();

  //if the cell next to hour: in source sheet is not blank and has a number
  let hourCell = sourceSheet.getRange(10, 4).getValue();
  let hours = sourceSheet.getRange(row, column).getValue();

  var col = 3; //volunteer hour log sheet
  for (let i = 0; i < numDates; i++) {
    if (values[0][col].toString() == event.toString() && values[1][col].toString() == date.toString()) {
      for (let count = 0; count < rowLength; count++) {
        let rowNumberName = calculationSheet.getRange(calcRow, calcColumnName).getValue();
        let columnNumberDate = calculationSheet.getRange(calcRow,calcColumnDate).getValue();
        hours = sourceSheet.getRange(row, column).getValue();
        if (hourCell !== "" && !isNaN(hourCell) && hours == "") {
          hours = hourCell;
        }

        hourLogSheet.getRange(parseInt(rowNumberName), parseInt(columnNumberDate)).setValue(hours);
        calcRow++;
        row++;
      }
      break;
    }
    col++;
  }

  //erase the content in this range
  if ((hourCell !== "" && !isNaN(hourCell)) || hours !== "") {
    sourceSheet.getRange('A13:F275').clearContent();
  }
}

//current time button
function setTime() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var selectedCell = sheet.getActiveRange();
  
  var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
  var currentTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'hh:mm a');
  
  var currentDateTime = currentDate + ' ' + currentTime;
  
  selectedCell.setValue(currentDateTime);
  selectedCell.setNumberFormat('hh:mm AM/PM');
}

//update hours button on team summary sheet
function teamSummaryHours() {
  var targetSheet = sheet.getSheetByName('Team Summary');

  var range = hourLogSheet.getDataRange();
  var values = range.getValues();

  let sourceRow = 6; //team summary sheet -- should change name bc confusion
  var sum = 0;

  for (let count = 0; count < numEvents; count++) {
    var specificValue = targetSheet.getRange(sourceRow, 2).getValue(); 
    var col = 3;

    for (var date = 0; date < numDates; date++) {
      if (values[0][col] === specificValue) {
        for (var i = 2; i < numPeople + 2; i++) {
          if (typeof values[i][col] === "number") {
            sum = sum + parseFloat(values[i][col]);
          }
        }
      }
      col++;
    }  

    var truncate = sum.toFixed(2);
    //Logger.log("Sum: " + truncate);
    targetSheet.getRange(sourceRow, 3).setValue(truncate);
    sum = 0;
    sourceRow++;
  }
}

function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  
  if ((sheet.getName() === 'Enter in Hours' || sheet.getName() === 'Enter in Hours 2') && (range.getColumn() === 4 || range.getColumn() === 5 ||range.getColumn() === 6 || range.getColumn() === 7) && range.getRow() >= 13) {
    var timeValue = range.getValue();
    var dateCell = sheet.getRange(7, 3); // Assuming the date is in the adjacent column (column B)
    var dateValue = dateCell.getValue();
    
    if (dateValue instanceof Date && !isNaN(dateValue)) {
      dateValue.setHours(timeValue.getHours(), timeValue.getMinutes(), timeValue.getSeconds());
      range.setValue(dateValue);
    }
  }
}
