/*

//alternative colors
//const primaryColor = "white"
//const secondaryColor = "Light Gray"
const ssG = SpreadsheetApp.openById("1XGTyoQay5PMSwNGauCWZko3tdZd2dNRU108Q1XGNBj4")
const generalDataSheet = css.getSheetById(1122725972);
const gMeetingSheetF = ssG.getSheetById(1780380128);
const gMeetingSheetSoph = ssG.getSheetById(204323695);
const gMeetingSheetJ = ssG.getSheetById(769892269);
const gMeetingSheetS = ssG.getSheetById(1315619862);
const gMeetingSheetOf = ssG.getSheetById(1448445140);
const attendanceSheets = [gMeetingSheetF, gMeetingSheetSoph, gMeetingSheetJ, gMeetingSheetS, gMeetingSheetOf];    //cody no acccess to attendance gs so i write based on what i reading and assume

const generalDataSheetStartingRow = 2

var frAttend, spAttend, juAttend, snAttend

                                                                                      //this functions purpose is ambigious, its doing too much, consider breaking it up into smaller functions to
function manuealAttend() {                                                            //explain whats happening in batch of lines
  console.log("Function started - running without trigger check");
  
  var frAttend, spAttend, juAttend, snAttend, ofAttend;
  var sheetColumns = {}; // Object to store column for each sheet

  var attendanceData = generalDataSheet.getRange(7,6,2).getValues().flat();
  console.log("Attendance Data:", attendanceData);
  
  // Process each sheet and store its column                                          //a more accurate comment for this for loop would be "Makes new col and sets event data in each sheet"
  for (i = 0; i < attendanceSheets.length; i++){                                      //more of a nit-pick but if ur going to comment (which u should for future projects and shared asignments),
    var sheet = attendanceSheets[i];                                                  //comements should explain why you implimented something as you did (i only started learning cuz epic project)
    console.log("Processing sheet: " + sheet.getName()); 
    
    // Insert new column
    sheet.insertColumnAfter(sheet.getLastColumn());
    var currentColumn = sheet.getLastColumn();                                        //are these sheet columns not made in parallel? wouldn't each sheet have the same last column (currentColumn)?
    sheetColumns[sheet.getName()] = currentColumn; // Store column for this sheet     //if currentCol is the same for all the sheets can't you just get the currentCol from just one sheet?
                                                                                        //just saw in manuelHour that they are not in fact similar but good way of handling
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

  // ===== CREATE BATCH STORAGE =====
  var checkboxBatch = {                           
    "Freshmen (9th)": { sheet: gMeetingSheetF, rows: [], column: null },
    "Sophomores (10th)": { sheet: gMeetingSheetSoph, rows: [], column: null },
    "Juniors (11th)": { sheet: gMeetingSheetJ, rows: [], column: null },
    "Seniors (12th)": { sheet: gMeetingSheetS, rows: [], column: null },
    "Officers": { sheet: gMeetingSheetOf, rows: [], column: null }
  };
  
  var rowColors = new Array(allNames.length).fill(null); // Store colors for batch update

  for (let i = 0; i < allNames.length; i++) {                                       //this for loop is ginormous...
    if (allNames[i] == "") continue;
                                              
    var targetSheet, targetColumn, nameTable;                                       //targetSheet unused variable
    var sheetName = "";
    
    //catagorize person into grade
    switch (gradeArray[i][0].toString().toLowerCase()) {
      case "freshman":
        if (frAttend == null) frAttend = getData(gMeetingSheetF, 8, 1, 1, 3);
        targetSheet = gMeetingSheetF;
        targetColumn = sheetColumns["Freshmen (9th)"];
        nameTable = frAttend;
        sheetName = "Freshmen (9th)";
        break;
      case "sophomore":
        if (spAttend == null) spAttend = getData(gMeetingSheetSoph, 8, 1, 1, 3);
        targetSheet = gMeetingSheetSoph;
        targetColumn = sheetColumns["Sophomores (10th)"];
        nameTable = spAttend;
        sheetName = "Sophomores (10th)";
        break;
      case "junior":
        if (juAttend == null) juAttend = getData(gMeetingSheetJ, 8, 1, 1, 3);
        targetSheet = gMeetingSheetJ;
        targetColumn = sheetColumns["Juniors (11th)"];
        nameTable = juAttend;
        sheetName = "Juniors (11th)";
        break;
      case "senior":
        if (snAttend == null) snAttend = getData(gMeetingSheetS, 8, 1, 1, 3);
        targetSheet = gMeetingSheetS;
        targetColumn = sheetColumns["Seniors (12th)"];
        nameTable = snAttend;
        sheetName = "Seniors (12th)";
        break;
      case "officer":
        if (ofAttend == null) ofAttend = getData(gMeetingSheetOf, 8, 1, 1, 3);
        targetSheet = gMeetingSheetOf;
        targetColumn = sheetColumns["Officers"];
        nameTable = ofAttend;
        sheetName = "Officers";
        break;
      default:
        console.log("Unknown grade: " + gradeArray[i][0]);
        continue;
    }
    
    // Find the person
    let index = nameTable.findIndex(name => name == allNames[i]);
    
    if (index != -1) {
      // Store for batch checkbox insertion
      checkboxBatch[sheetName].rows.push(index + 8);
      checkboxBatch[sheetName].column = targetColumn; // Set column                     //this doesnt need to be set each person, only once outside the for loop or use the sheetCol that u've alrdy
      rowColors[i] = "#b7e1cd"; // Green                                                //specified
    } else {
      rowColors[i] = "#ea9999"; // Pink/red
      console.log("NOT FOUND: " + allNames[i]);
    }
  }
  
  // ===== OPTION 2: ONE BATCH PER SHEET =====
  console.log("Inserting checkboxes in batches...");
  
  for (let sheetName in checkboxBatch) {
    let batch = checkboxBatch[sheetName];
    if (batch.rows.length === 0 || !batch.column) continue;
    
    // Get the min and max row for this sheet                                           //does this work... never tried but just asking...
    let minRow = Math.min(...batch.rows);
    let maxRow = Math.max(...batch.rows);
    let rowCount = maxRow - minRow + 1;                                                 //isnt row count the length of the batch?
    
    console.log(`${sheetName}: ${batch.rows.length} attendees across rows ${minRow}-${maxRow}`);
    
    // Get the range covering ALL rows from min to max
    let range = batch.sheet.getRange(minRow, batch.column, rowCount, 1);                //could hard code this line with starting row and get last row minus data start row reducing the need 
                                                                                        //reducing the need to calculate min and max row
    // Insert checkboxes for the entire range at once                                   //also instead of true or false (checkbox), could just color the row, slightly more performant
    range.insertCheckboxes();
    
    // Create an array with FALSE for all rows
    let values = new Array(rowCount);                                                   //could just say .fill(false)
    for (let j = 0; j < rowCount; j++) {
      values[j] = [false];
    }
    
    // Set TRUE only for the rows that actually have attendees
    batch.rows.forEach(row => {                                                         //could've been writen as a for loop i < batch.length {values[i] = true}
      values[row - minRow][0] = true;
    });
    
    // Set all values at once
    range.setValues(values);
  }
  
  // ===== BATCH UPDATE COLORS =====
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
