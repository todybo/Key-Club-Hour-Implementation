const meetingSpreadsheet = SpreadsheetApp.openById(mohSettings.getRange("A13").getValue()); // General Meeting spreadsheet
const meetingFreshmanSheet  = meetingSpreadsheet.getSheetById(mohSettings.getRange("B13").getValue());
const meetingSophomoreSheet = meetingSpreadsheet.getSheetById(mohSettings.getRange("C13").getValue());
const meetingJuniorSheet    = meetingSpreadsheet.getSheetById(mohSettings.getRange("D13").getValue());
const meetingSeniorSheet    = meetingSpreadsheet.getSheetById(mohSettings.getRange("E13").getValue());
const meetingOfficerSheet   = meetingSpreadsheet.getSheetById(mohSettings.getRange("F13").getValue());
const attendanceSheets = [meetingFreshmanSheet, meetingSophomoreSheet, meetingJuniorSheet, meetingSeniorSheet, meetingOfficerSheet];

const attendanceStartRow = mohSettings.getRange("A15").getValue();       // starting row of names
const meetingTitleRow = mohSettings.getRange("B15").getValue();          // first row of the merged title cell
const meetingTitleRowSpan = mohSettings.getRange("C15").getValue();      // # of rows the title spans
const meetingOfficerStartCol = mohSettings.getRange("D15").getValue();   // starting column of officer names

const attendanceInputSheet = mohFunction.getSheetById(1122725972); // sheet where attendees are entered
const attendanceInputStartRow = 2;

/*
 * PERFORMANCE NOTES (manuealAttend)
 * ---------------------------------
 * 1. O(1) name lookups: each attendance roster is converted once into a Map
 *    { cleanedName -> rosterRow } with buildNameIndex() (defined in
 *    GiveHourList.js, shared global scope). This replaces the per-person
 *    findIndex() scan: O(N*M) -> O(N + M).
 * 2. Batched colors: row colors go into a 2D buffer (seeded from
 *    getBackgrounds() so skipped rows keep their color) and are written with
 *    ONE setBackgrounds() call instead of one setBackground() per row.
 * 3. Checkbox batching (one insertCheckboxes() + setValues() per sheet) is
 *    unchanged.
 */
function manuealAttend() {
  console.log("Function started - running without trigger check");

  // Name -> roster row lookups, built only when someone from that sheet appears.
  var freshmanIndex, sophomoreIndex, juniorIndex, seniorIndex, officerIndex;
  var newColumnBySheet = {};

  var [meetingTitle, meetingDate] = attendanceInputSheet.getRange(7,6,2).getValues().flat();
  console.log("Meeting info:", meetingTitle, meetingDate);

  for (let i = 0; i < attendanceSheets.length; i++){
    var sheet = attendanceSheets[i];
    console.log("Processing sheet: " + sheet.getName());

    sheet.insertColumnAfter(sheet.getLastColumn());
    var newColumn = sheet.getLastColumn();
    newColumnBySheet[sheet.getName()] = newColumn;

    console.log("Inserted column at: " + newColumn);

    var titleRange = sheet.getRange(meetingTitleRow, newColumn, meetingTitleRowSpan, 1);
    titleRange.merge();
    titleRange.setValue(meetingTitle);
    titleRange.setNumberFormat("m/d");

    var dateCell = sheet.getRange(meetingTitleRow + meetingTitleRowSpan, newColumn);
    dateCell.setValue(meetingDate);
  }

  console.log("Starting to process names...");
  var names = getData(attendanceInputSheet, attendanceInputStartRow, 1, 0, 2);
  var grades = attendanceInputSheet.getRange(attendanceInputStartRow,3,attendanceInputSheet.getLastRow(),1).getValues();

  console.log("Found " + names.length + " names to process");

  var checkboxBatch = {
    "Freshmen (9th)":    { sheet: meetingFreshmanSheet,  rows: [], column: null },
    "Sophomores (10th)": { sheet: meetingSophomoreSheet, rows: [], column: null },
    "Juniors (11th)":    { sheet: meetingJuniorSheet,    rows: [], column: null },
    "Seniors (12th)":    { sheet: meetingSeniorSheet,    rows: [], column: null },
    "Officers":          { sheet: meetingOfficerSheet,   rows: [], column: null }
  };

  // One read of current colors; matched/unmatched rows are overwritten in memory.
  var rowColors = names.length > 0
    ? attendanceInputSheet.getRange(attendanceInputStartRow, 1, names.length, 4).getBackgrounds()
    : [];

  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    if (name == "") continue;

    var nameIndex;
    var sheetName = "";

    switch (grades[i][0].toString().toLowerCase()) {
      case "freshman":
        if (freshmanIndex == null) freshmanIndex = buildNameIndex(getData(meetingFreshmanSheet, attendanceStartRow, 1, 1, 3));
        nameIndex = freshmanIndex;
        sheetName = "Freshmen (9th)";
        break;
      case "sophomore":
        if (sophomoreIndex == null) sophomoreIndex = buildNameIndex(getData(meetingSophomoreSheet, attendanceStartRow, 1, 1, 3));
        nameIndex = sophomoreIndex;
        sheetName = "Sophomores (10th)";
        break;
      case "junior":
        if (juniorIndex == null) juniorIndex = buildNameIndex(getData(meetingJuniorSheet, attendanceStartRow, 1, 1, 3));
        nameIndex = juniorIndex;
        sheetName = "Juniors (11th)";
        break;
      case "senior":
        if (seniorIndex == null) seniorIndex = buildNameIndex(getData(meetingSeniorSheet, attendanceStartRow, 1, 1, 3));
        nameIndex = seniorIndex;
        sheetName = "Seniors (12th)";
        break;
      case "officer":
        if (officerIndex == null) officerIndex = buildNameIndex(getData(meetingOfficerSheet, attendanceStartRow, meetingOfficerStartCol, 1, 1));
        nameIndex = officerIndex;
        sheetName = "Officers";
        break;
      default:
        console.log("Unknown grade: " + grades[i][0]);
        continue;
    }

    // O(1) Map lookup (was an O(M) findIndex scan).
    let rosterRow = nameIndex.has(name) ? nameIndex.get(name) : -1;

    if (rosterRow != -1) {
      checkboxBatch[sheetName].rows.push(rosterRow + attendanceStartRow);
      checkboxBatch[sheetName].column = newColumnBySheet[sheetName];
      rowColors[i] = [foundPplColour, foundPplColour, foundPplColour, foundPplColour];         // #b7e1cd
    } else {
      rowColors[i] = [didntFindPplColor, didntFindPplColor, didntFindPplColor, didntFindPplColor]; // #ea9999
      console.log("NOT FOUND: " + name);
    }
  }

  console.log("Inserting checkboxes in batches...");

  for (let sheetName in checkboxBatch) {
    let batch = checkboxBatch[sheetName];
    if (batch.rows.length === 0 || !batch.column) continue;

    let firstRow = Math.min(...batch.rows);
    let lastRow = Math.max(...batch.rows);
    let rowCount = lastRow - firstRow + 1;

    console.log(`${sheetName}: ${batch.rows.length} attendees across rows ${firstRow}-${lastRow}`);

    let checkboxRange = batch.sheet.getRange(firstRow, batch.column, rowCount, 1);
    checkboxRange.insertCheckboxes();

    let checked = new Array(rowCount).fill(false).map(() => [false]);
    batch.rows.forEach(row => {
      checked[row - firstRow][0] = true;
    });

    checkboxRange.setValues(checked);
  }

  console.log("Updating colors...");
  // Single setBackgrounds() call replaces one setBackground() per row.
  if (names.length > 0) {
    attendanceInputSheet.getRange(attendanceInputStartRow, 1, names.length, 4).setBackgrounds(rowColors);
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
