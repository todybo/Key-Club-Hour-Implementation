const gMeetingSS = SpreadsheetApp.openById(mohSettings.getRange("A13").getValue()); // General Meeting spreadsheet
const gMeetingSheetF = gMeetingSS.getSheetById(mohSettings.getRange("B13").getValue());
const gMeetingSheetSoph = gMeetingSS.getSheetById(mohSettings.getRange("C13").getValue());
const gMeetingSheetJ = gMeetingSS.getSheetById(mohSettings.getRange("D13").getValue());
const gMeetingSheetS = gMeetingSS.getSheetById(mohSettings.getRange("E13").getValue());
const gMeetingSheetOf = gMeetingSS.getSheetById(mohSettings.getRange("F13").getValue());
const attendanceSheets = [gMeetingSheetF, gMeetingSheetSoph, gMeetingSheetJ, gMeetingSheetS, gMeetingSheetOf];

const attendanceStartingRow = mohSettings.getRange("A15").getValue();   // starting row of names
const meetingDateRow = mohSettings.getRange("B15").getValue();          // row of meeting date (title merge start)
const meetingDateRowSpan = mohSettings.getRange("C15").getValue();      // # of rows date spans
const gMeetingOfficerStartingCol = mohSettings.getRange("D15").getValue(); // starting column of officers

const generalDataSheet = mohFunction.getSheetById(1122725972);
const generalDataSheetStartingRow = 2;

function manuealAttend() {
  console.log("Function started - running without trigger check");

  var frAttend, spAttend, juAttend, snAttend, ofAttend;
  var sheetColumns = {};

  var attendanceData = generalDataSheet.getRange(7,6,2).getValues().flat();
  console.log("Attendance Data:", attendanceData);

  for (let i = 0; i < attendanceSheets.length; i++){
    var sheet = attendanceSheets[i];
    console.log("Processing sheet: " + sheet.getName());

    sheet.insertColumnAfter(sheet.getLastColumn());
    var currentColumn = sheet.getLastColumn();
    sheetColumns[sheet.getName()] = currentColumn;

    console.log("Inserted column at: " + currentColumn);

    var titleRange = sheet.getRange(meetingDateRow, currentColumn, meetingDateRowSpan, 1);
    titleRange.merge();
    titleRange.setValue(attendanceData[0]);
    titleRange.setNumberFormat("m/d");

    var datePut = sheet.getRange(meetingDateRow + meetingDateRowSpan, currentColumn);
    datePut.setValue(attendanceData[1]);
  }

  console.log("Starting to process names...");
  var allNames = getData(generalDataSheet, generalDataSheetStartingRow, 1, 0, 2);
  var gradeArray = generalDataSheet.getRange(generalDataSheetStartingRow,3,generalDataSheet.getLastRow(),1).getValues();

  console.log("Found " + allNames.length + " names to process");

  var checkboxBatch = {
    "Freshmen (9th)": { sheet: gMeetingSheetF, rows: [], column: null },
    "Sophomores (10th)": { sheet: gMeetingSheetSoph, rows: [], column: null },
    "Juniors (11th)": { sheet: gMeetingSheetJ, rows: [], column: null },
    "Seniors (12th)": { sheet: gMeetingSheetS, rows: [], column: null },
    "Officers": { sheet: gMeetingSheetOf, rows: [], column: null }
  };

  var rowColors = new Array(allNames.length).fill(null);

  for (let i = 0; i < allNames.length; i++) {
    if (allNames[i] == "") continue;

    var targetColumn, nameTable;
    var sheetName = "";

    switch (gradeArray[i][0].toString().toLowerCase()) {
      case "freshman":
        if (frAttend == null) frAttend = getData(gMeetingSheetF, attendanceStartingRow, 1, 1, 3);
        targetColumn = sheetColumns["Freshmen (9th)"];
        nameTable = frAttend;
        sheetName = "Freshmen (9th)";
        break;
      case "sophomore":
        if (spAttend == null) spAttend = getData(gMeetingSheetSoph, attendanceStartingRow, 1, 1, 3);
        targetColumn = sheetColumns["Sophomores (10th)"];
        nameTable = spAttend;
        sheetName = "Sophomores (10th)";
        break;
      case "junior":
        if (juAttend == null) juAttend = getData(gMeetingSheetJ, attendanceStartingRow, 1, 1, 3);
        targetColumn = sheetColumns["Juniors (11th)"];
        nameTable = juAttend;
        sheetName = "Juniors (11th)";
        break;
      case "senior":
        if (snAttend == null) snAttend = getData(gMeetingSheetS, attendanceStartingRow, 1, 1, 3);
        targetColumn = sheetColumns["Seniors (12th)"];
        nameTable = snAttend;
        sheetName = "Seniors (12th)";
        break;
      case "officer":
        if (ofAttend == null) ofAttend = getData(gMeetingSheetOf, attendanceStartingRow, gMeetingOfficerStartingCol, 1, 1);
        targetColumn = sheetColumns["Officers"];
        nameTable = ofAttend;
        sheetName = "Officers";
        break;
      default:
        console.log("Unknown grade: " + gradeArray[i][0]);
        continue;
    }

    let index = nameTable.findIndex(name => name == allNames[i]);

    if (index != -1) {
      checkboxBatch[sheetName].rows.push(index + attendanceStartingRow);
      checkboxBatch[sheetName].column = targetColumn;
      rowColors[i] = "#b7e1cd";
    } else {
      rowColors[i] = "#ea9999";
      console.log("NOT FOUND: " + allNames[i]);
    }
  }

  console.log("Inserting checkboxes in batches...");

  for (let sheetName in checkboxBatch) {
    let batch = checkboxBatch[sheetName];
    if (batch.rows.length === 0 || !batch.column) continue;

    let minRow = Math.min(...batch.rows);
    let maxRow = Math.max(...batch.rows);
    let rowCount = maxRow - minRow + 1;

    console.log(`${sheetName}: ${batch.rows.length} attendees across rows ${minRow}-${maxRow}`);

    let range = batch.sheet.getRange(minRow, batch.column, rowCount, 1);
    range.insertCheckboxes();

    let values = new Array(rowCount).fill(false).map(() => [false]);
    batch.rows.forEach(row => {
      values[row - minRow][0] = true;
    });

    range.setValues(values);
  }

  console.log("Updating colors...");
  for (let i = 0; i < rowColors.length; i++) {
    if (rowColors[i] !== null) {
      generalDataSheet.getRange(i + generalDataSheetStartingRow, 1, 1, 4).setBackground(rowColors[i]);
    }
  }

  console.log("Finished processing all names");
}


/* function manuealAttend() {                                                             //theres always going to be some lag because of searching
  console.log("Function started - running without trigger check");
  
  var frAttend, spAttend, juAttend, snAttend, ofAttend;
  var sheetColumns = {}; // Object to store column for each sheet

  var attendanceData = generalDataSheet.getRange(7,6,2).getValues().flat();
  console.log("Attendance Data:", attendanceData);
  
  // Process each sheet and store its column
  for (i = 0; i < attendanceSheets.length; i++){
    var sheet = attendanceSheets[i];
    console.log("Processing sheet: " + sheet.getName()); 
    
    // Insert new column
    sheet.insertColumnAfter(sheet.getLastColumn());
    var currentColumn = sheet.getLastColumn();
    sheetColumns[sheet.getName()] = currentColumn; // Store column for this sheet
    
    console.log("Inserted column at: " + currentColumn);
    
    // Merge and set event title (rows 3-6)
    var titleRange = sheet.getRange(3, currentColumn, 4, 1);
    titleRange.merge();
    titleRange.setValue(attendanceData[0]); // Event title
    titleRange.setNumberFormat("m/d");
    
    // Set date in row 7
    var datePut = sheet.getRange(7, currentColumn);
    datePut.setValue(attendanceData[1]); // Date
  }

  // Process names
  console.log("Starting to process names...");
  var allNames = getData(generalDataSheet, generalDataSheetStartingRow, 1, 0, 2);
  var gradeArray = generalDataSheet.getRange(generalDataSheetStartingRow,3,generalDataSheet.getLastRow(),1).getValues();
  
  console.log("Found " + allNames.length + " names to process");

  for (let i = 0; i < allNames.length; i++) {
    if (allNames[i] == "") continue;

    var targetSheet, targetColumn, nameTable;
    
    switch (gradeArray[i][0].toString().toLowerCase()) {
      case "freshman":
        if (frAttend == null) frAttend = getData(gMeetingSheetF, 8, 1, 1, 3);
        targetSheet = gMeetingSheetF;
        targetColumn = sheetColumns["Freshmen (9th)"]; // Use stored column
        nameTable = frAttend;
        break;
      case "sophomore":
        if (spAttend == null) spAttend = getData(gMeetingSheetSoph, 8, 1, 1, 3);
        targetSheet = gMeetingSheetSoph;
        targetColumn = sheetColumns["Sophomores (10th)"];
        nameTable = spAttend;
        break;
      case "junior":
        if (juAttend == null) juAttend = getData(gMeetingSheetJ, 8, 1, 1, 3);
        targetSheet = gMeetingSheetJ;
        targetColumn = sheetColumns["Juniors (11th)"];
        nameTable = juAttend;
        break;
      case "senior":
        if (snAttend == null) snAttend = getData(gMeetingSheetS, 8, 1, 1, 3);
        targetSheet = gMeetingSheetS;
        targetColumn = sheetColumns["Seniors (12th)"];
        nameTable = snAttend;
        break;
      case "officer": // ADDED OFFICER CASE
        if (ofAttend == null) ofAttend = getData(gMeetingSheetOf, 8, 1, 1, 3);
        targetSheet = gMeetingSheetOf;
        targetColumn = sheetColumns["Officers"];
        nameTable = ofAttend;
        break;
      default:
        console.log("Unknown grade: " + gradeArray[i][0]);
        continue;
    }
    
    markAttendance(i, targetSheet, nameTable, allNames[i], targetColumn);
  }
  
  console.log("Finished processing all names");
}

function markAttendance(i, sheet, nameTable, personName, column) {
  //console.log("Looking for: " + personName + " in " + sheet.getName() + " at column " + column);
  
  let index = nameTable.findIndex(name => name == personName);
  //console.log("Found at index: " + index);
  
  if (index != -1) {
    var cell = sheet.getRange(index + 8, column);
    //console.log("Marking attendance at: " + (index + 8) + "," + column);
    
    cell.insertCheckboxes();
    cell.setValue(true);
    
    generalDataSheet.getRange(i + generalDataSheetStartingRow, 1, 1, 4).setBackground("#b7e1cd");
    console.log("Marked present: " + personName);
  } else {
    generalDataSheet.getRange(i + generalDataSheetStartingRow, 1, 1, 4).setBackground("#ea9999");
    console.log("NOT FOUND: " + personName);
  }
} */
