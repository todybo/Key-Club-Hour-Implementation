/*

//alternative colors
//const primaryColor = "white"
//const secondaryColor = "Light Gray"
const otherDataSheet = css.getSheetById(305205063);

const otherDataSheetStartingRow = 2
const warningColor = "#ffe599"

//for "online idea" sheet
function manuealHour() {
  var isReady = otherDataSheet.getRange(1, 6);
  if (isReady.getValue() == true) {
    isReady.setValue(false);
  } else {
    return;
  }

  var frHour, spHour, juHour, snHour, ofHour
  ofHour = getData(ofsheet, peopleSheetStartingRow, 2, 1, 3)
  var nonsnCurrentColumn,snCurrentColumn, ofCurrentColumn
  var range
  var datePut

  var eventData = otherDataSheet.getRange(7,6,2).getValues().flat();
  console.log(eventData);
  //creates header and sets event info in new column for each sheet
  for (i = 0; i < hourSheets.length; i++){
    hourSheets[i].insertColumnBefore(hourSheets[i].getLastColumn() - 2);
    if (hourSheets[i] == ofsheet) {                                         //i remeber the ofsheet looking a little different but couldnt you format each sheet to look the same to not need to
      ofCurrentColumn = hourSheets[i].getLastColumn() - 3;                  //hard code this in thereby reducing the if statement here
      hourSheets[i].getRange(1, ofCurrentColumn, 2).merge()
      range = hourSheets[i].getRange(3, ofCurrentColumn, 4)
      datePut = hourSheets[i].getRange(7, ofCurrentColumn)
      
    } else if (hourSheets[i] == ssheet){
      snCurrentColumn = hourSheets[i].getLastColumn() - 3;
      hourSheets[i].getRange(1, snCurrentColumn, 2).merge()
      range = hourSheets[i].getRange(3, snCurrentColumn, 4)
      datePut = hourSheets[i].getRange(7, snCurrentColumn)
    } else {
      nonsnCurrentColumn = hourSheets[i].getLastColumn() - 3;
      hourSheets[i].getRange(1, nonsnCurrentColumn, 2).merge()
      range = hourSheets[i].getRange(3, nonsnCurrentColumn, 4)
      datePut = hourSheets[i].getRange(7, nonsnCurrentColumn)
    }
    range.merge()
    range.setValue(eventData[1])
    datePut.setValue(eventData[0])
    datePut.setNumberFormat("m/d")
  }

  var allNames = getData(otherDataSheet, otherDataSheetStartingRow, 1, 0, 2);
  var gradeArray = otherDataSheet.getRange(otherDataSheetStartingRow,3,otherDataSheet.getLastRow(),2).getValues();
  for (let i = 0; i < allNames.length; i++) {
    if (allNames[i] == "") {Logger.log("name not there"); continue}
    if (gradeArray[i][1] == "" || typeof gradeArray[i][1] != "number") { //hour not here                    //lovely error handling (personal preference but i don't like moving more than 4-5
      otherDataSheet.getRange(i + otherDataSheetStartingRow, 4, 1, 1).setBackground(warningColor);          //indents in, so having the success as the else of the if is very good programing style)
      continue
    }

    var officerCheck = checkForPerson(i, ofsheet, ofHour, allNames[i], gradeArray[i][1], ofCurrentColumn)
    if (officerCheck) {
      otherDataSheet.getRange(i + otherDataSheetStartingRow, 3, 1, 1).setValue("officer");
      continue
    }

    switch (gradeArray[i][0].toString().toLowerCase()) {
      case "freshman":
        if (frHour == null) {frHour = getData(fsheet, peopleSheetStartingRow, 1, 1, 3);}
        checkForPerson(i, fsheet, frHour, allNames[i], gradeArray[i][1], nonsnCurrentColumn)
        break;
      case "sophomore":
        if (spHour == null) {spHour = getData(spsheet, peopleSheetStartingRow, 1, 1, 3);}
        checkForPerson(i, spsheet, spHour, allNames[i], gradeArray[i][1], nonsnCurrentColumn)
        break;
      case "junior":
        if (juHour == null) {juHour = getData(jsheet, peopleSheetStartingRow, 1, 1, 3);}
        checkForPerson(i, jsheet, juHour, allNames[i], gradeArray[i][1], nonsnCurrentColumn)
        break;
      case "senior":
        if (snHour == null) {snHour = getData(ssheet, peopleSheetStartingRow, 1, 1, 3);}
        checkForPerson(i, ssheet, snHour, allNames[i], gradeArray[i][1], snCurrentColumn)
        break;
      default: //grade not found
        otherDataSheet.getRange(i + otherDataSheetStartingRow, 3, 1, 1).setBackground(warningColor);
        continue;
    }
  }
  isReady.setValue(true);
}

function checkForPerson(i, hourSheet, hourTable, personName, hours, column) {
  let index = hourTable.findIndex(name => name == personName)
  if (index == -1) {
    otherDataSheet.getRange(i + otherDataSheetStartingRow, 1, 1, 4).setBackground(didntFindPplColor);
    return false
  }
  hourSheet.getRange(index + peopleSheetStartingRow, column).setValue(hours);
  otherDataSheet.getRange(i + otherDataSheetStartingRow, 1, 1, 4).setBackground(foundPplColour);
  return true
}

//picture buttons
function button() {
  var drawings = dataSheet.getDrawings()
  var balls = dataSheet.getRange(4,19,2).getValues().flat();
  drawings[balls[1]].setWidth(1)
  drawings[balls[1]].setHeight(1)

  var drawing
  main()

  if (balls[0] == true) {
    var rand = randNum(0, drawings.length - 1)
    drawing = drawings[rand]
    dataSheet.getRange(5,19).setValue(rand)
  } else {
    drawing = drawings[balls[1]+1]
    dataSheet.getRange(5,19).setValue(balls[1]+1)
    if (balls[1]+2 == drawings.length) {
      dataSheet.getRange(4,19).setValue(true)
    }
  }
  
  //drawing.setPosition(3, 16, 5, 5)
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
} */