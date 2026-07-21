const otherDataSheet = mohFunction.getSheetById(305205063);
const otherDataSheetStartingRow = 2;
const warningColor = "#ffe599";

// Local helper — mirrors the column-insertion logic inside HourMain's grantHour,
// duplicated here (not extracted from HourMain) so HourMain stays untouched.
function ghlInsertHourColumn(hourSheet, dateValue, nameValue) {
  var beforeLastCol = hourSheet.getMaxColumns();
  var offset = (hourSheet == ofSheet)
    ? mohSettings.getRange("C9").getValue()
    : mohSettings.getRange("B9").getValue();

  var targetColumn = beforeLastCol - offset + 1;

  if (offset === 0) {
    hourSheet.insertColumnAfter(beforeLastCol);
  } else if (targetColumn >= hourSheet.getMaxColumns()) {
    hourSheet.insertColumns(hourSheet.getMaxColumns(), 1);
  } else {
    hourSheet.insertColumnBefore(targetColumn);
  }

  var currentColumn = hourSheet.getMaxColumns() - offset;

  var headerRange = hourSheet.getRange(
    mohSettings.getRange("E9").getValue(),
    currentColumn,
    mohSettings.getRange("F9").getValue()
  ).merge();
  headerRange.setValue(nameValue);

  var dateRow = mohSettings.getRange("D9").getValue();
  var dateCell = hourSheet.getRange(dateRow, currentColumn);
  dateCell.setValue(dateValue);
  dateCell.setNumberFormat("m/d");

  return currentColumn;
}

function manuealHour() {
  var isReady = otherDataSheet.getRange(1, 6);
  if (isReady.getValue() == true) {
    isReady.setValue(false);
  } else {
    return;
  }

  var frHour, spHour, juHour, snHour;
  var ofHour = getData(ofSheet, peopleSheetStartingRow, officerStartingCol, 1, 3);

  var eventData = otherDataSheet.getRange(7, 6, 2).getValues().flat(); // [date, name]

  // insert + header/date columns for every sheet, keyed by hourSheets index
  var columns = [];
  for (let i = 0; i < hourSheets.length; i++) {
    columns[i] = ghlInsertHourColumn(hourSheets[i], eventData[0], eventData[1]);
  }
  var freshCol = columns[0], sophCol = columns[1], juniorCol = columns[2],
      seniorCol = columns[3], ofCol = columns[4];

  var allNames = getData(otherDataSheet, otherDataSheetStartingRow, 1, 0, 2);
  var gradeArray = otherDataSheet.getRange(otherDataSheetStartingRow, 3, otherDataSheet.getLastRow(), 2).getValues();

  for (let i = 0; i < allNames.length; i++) {
    if (allNames[i] == "") { Logger.log("name not there"); continue; }
    if (gradeArray[i][1] == "" || typeof gradeArray[i][1] != "number") {
      otherDataSheet.getRange(i + otherDataSheetStartingRow, 4, 1, 1).setBackground(warningColor);
      continue;
    }

    var officerCheck = checkForPerson(i, ofSheet, ofHour, allNames[i], gradeArray[i][1], ofCol);
    if (officerCheck) {
      otherDataSheet.getRange(i + otherDataSheetStartingRow, 3, 1, 1).setValue("officer");
      continue;
    }

    switch (gradeArray[i][0].toString().toLowerCase()) {
      case "freshman":
        if (frHour == null) frHour = getData(freshSheet, peopleSheetStartingRow, 1, 1, 3);
        checkForPerson(i, freshSheet, frHour, allNames[i], gradeArray[i][1], freshCol);
        break;
      case "sophomore":
        if (spHour == null) spHour = getData(sophSheet, peopleSheetStartingRow, 1, 1, 3);
        checkForPerson(i, sophSheet, spHour, allNames[i], gradeArray[i][1], sophCol);
        break;
      case "junior":
        if (juHour == null) juHour = getData(juniorSheet, peopleSheetStartingRow, 1, 1, 3);
        checkForPerson(i, juniorSheet, juHour, allNames[i], gradeArray[i][1], juniorCol);
        break;
      case "senior":
        if (snHour == null) snHour = getData(seniorSheet, peopleSheetStartingRow, 1, 1, 3);
        checkForPerson(i, seniorSheet, snHour, allNames[i], gradeArray[i][1], seniorCol);
        break;
      default:
        otherDataSheet.getRange(i + otherDataSheetStartingRow, 3, 1, 1).setBackground(warningColor);
        continue;
    }
  }
  isReady.setValue(true);
}

function checkForPerson(i, hourSheet, hourTable, personName, hours, column) {
  let index = hourTable.findIndex(name => name == personName);
  if (index == -1) {
    otherDataSheet.getRange(i + otherDataSheetStartingRow, 1, 1, 4).setBackground(didntFindPplColor);
    return false;
  }
  hourSheet.getRange(index + peopleSheetStartingRow, column).setValue(hours);
  otherDataSheet.getRange(i + otherDataSheetStartingRow, 1, 1, 4).setBackground(foundPplColour);
  return true;
}

//picture buttons
function button() {
  var drawings = mohFunctionData.getDrawings()
  var balls = mohFunctionData.getRange(4,19,2).getValues().flat();
  drawings[balls[1]].setWidth(1)
  drawings[balls[1]].setHeight(1)

  var drawing
  main()

  if (balls[0] == true) {
    var rand = randNum(0, drawings.length - 1)
    drawing = drawings[rand]
    mohFunctionData.getRange(5,19).setValue(rand)
  } else {
    drawing = drawings[balls[1]+1]
    mohFunctionData.getRange(5,19).setValue(balls[1]+1)
    if (balls[1]+2 == drawings.length) {
      mohFunctionData.getRange(4,19).setValue(true)
    }
  }
  
  drawing.setWidth(190)
  drawing.setHeight(116)
}

function button2() { //couldve just combined button functions imo                     //could've yes.... couldve had one function with sheet parameter and then 2 functions to set up the pic buttons
  var drawings = otherDataSheet.getDrawings()
  var balls = otherDataSheet.getRange(3,7,2).getValues().flat();
  drawings[balls[1]].setWidth(1)
  drawings[balls[1]].setHeight(1)

  var drawing
  manuealHour();

  if (balls[0] == true) {
    var rand = randNum(0, drawings.length - 1)
    drawing = drawings[rand]
    otherDataSheet.getRange(4,7).setValue(rand)
  } else {
    drawing = drawings[balls[1]+1]
    otherDataSheet.getRange(4,7).setValue(balls[1]+1)
    if (balls[1]+2 == drawings.length) {
      otherDataSheet.getRange(3,7).setValue(true)
    }
  }
  
  //drawing.setPosition(2, 5, 5, 5)
  drawing.setWidth(190)
  drawing.setHeight(116)
}

//for displaying a random picture
function randNum(min, max) {
  const minCeiled = Math.ceil(min);
  const maxFloored = Math.floor(max);
  return Math.floor(Math.random() * (maxFloored - minCeiled + 1)) + minCeiled;
}