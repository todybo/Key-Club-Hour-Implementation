const manualHoursSheet = mohFunction.getSheetById(305205063); // sheet where hours are entered by hand
const manualHoursStartRow = 2;
const warningColor = "#ffe599";

// Local helper — mirrors the column-insertion logic inside HourMain's grantHour,
// duplicated here (not extracted from HourMain) so HourMain stays untouched.
// Returns the column number of the newly inserted event column.
function insertEventColumn(hourSheet, eventDate, eventName) {
  var lastColumnBefore = hourSheet.getMaxColumns();
  var trailingColumns = (hourSheet == ofSheet)
    ? mohSettings.getRange("C9").getValue()
    : mohSettings.getRange("B9").getValue();

  var insertAt = lastColumnBefore - trailingColumns + 1;

  if (trailingColumns === 0) {
    hourSheet.insertColumnAfter(lastColumnBefore);
  } else if (insertAt >= hourSheet.getMaxColumns()) {
    hourSheet.insertColumns(hourSheet.getMaxColumns(), 1);
  } else {
    hourSheet.insertColumnBefore(insertAt);
  }

  var newColumn = hourSheet.getMaxColumns() - trailingColumns;

  var headerRange = hourSheet.getRange(
    mohSettings.getRange("E9").getValue(),
    newColumn,
    mohSettings.getRange("F9").getValue()
  ).merge();
  headerRange.setValue(eventName);

  var dateRow = mohSettings.getRange("D9").getValue();
  var dateCell = hourSheet.getRange(dateRow, newColumn);
  dateCell.setValue(eventDate);
  dateCell.setNumberFormat("m/d");

  return newColumn;
}

/*
 * PERFORMANCE NOTES (manuealHour / recordHours)
 * ---------------------------------------------
 * 1. O(1) name lookups: each roster returned by getData() is converted once into
 *    a Map { cleanedName -> rosterRow } via buildNameIndex(). The old per-person
 *    findIndex() scan (O(N*M) overall) is now O(N + M).
 * 2. Batched writes: no Spreadsheet API calls happen inside the name loop.
 *    - Row colors are buffered in a 2D array (seeded from getBackgrounds() so
 *      untouched cells keep their current color) and written with ONE
 *      setBackgrounds() call.
 *    - "officer" grade overrides are buffered and written with ONE setValues()
 *      over just the span of rows that changed.
 *    - Hours are buffered per grade sheet and written with ONE setValues() per
 *      sheet over the matched row span of the freshly inserted column.
 *    Write order inside the buffers mirrors the original call order, so
 *    "last write wins" behavior (e.g. officer miss → pink, then grade hit →
 *    green) is unchanged.
 */

// Builds { cleanedName -> rosterRow } in O(M). Keeps the FIRST occurrence so
// results are identical to Array.prototype.findIndex. A Map is used instead of
// a plain object so names such as "constructor" can't collide with prototype keys.
// Shared global — also used by GeneralMeetings.js.
function buildNameIndex(names) {
  const index = new Map();
  for (let row = 0; row < names.length; row++) {
    if (!index.has(names[row])) index.set(names[row], row);
  }
  return index;
}

// Writes a buffer of { sheetRow -> hours } to its column with one setValues().
// The column was just inserted by insertEventColumn, so the gap cells in the
// span are empty and writing "" there is a no-op.
function writeBufferedHours(hourBuffer) {
  var sheetRows = Object.keys(hourBuffer.hoursByRow).map(Number);
  if (sheetRows.length === 0) return;
  var firstRow = Math.min(...sheetRows);
  var lastRow = Math.max(...sheetRows);
  var columnValues = new Array(lastRow - firstRow + 1).fill(null).map(() => [""]);
  sheetRows.forEach(row => { columnValues[row - firstRow][0] = hourBuffer.hoursByRow[row]; });
  hourBuffer.sheet.getRange(firstRow, hourBuffer.column, columnValues.length, 1).setValues(columnValues);
}

function manuealHour() {
  var isReady = manualHoursSheet.getRange(1, 6);
  if (isReady.getValue() == true) {
    isReady.setValue(false);
  } else {
    return;
  }

  // Name -> roster row lookups. Grade rosters are only read if someone in that grade appears.
  var freshmanIndex, sophomoreIndex, juniorIndex, seniorIndex;
  var officerIndex = buildNameIndex(getData(ofSheet, peopleSheetStartingRow, officerStartingCol, 1, 3));

  var [eventDate, eventName] = manualHoursSheet.getRange(7, 6, 2).getValues().flat();

  // Insert the header/date column on every sheet (same order as hourSheets).
  var newColumns = hourSheets.map(sheet => insertEventColumn(sheet, eventDate, eventName));

  // Per-sheet hour buffers: { sheet, column, hoursByRow: { sheetRow: hours } }
  var freshmanHours  = { sheet: freshSheet,  column: newColumns[0], hoursByRow: {} };
  var sophomoreHours = { sheet: sophSheet,   column: newColumns[1], hoursByRow: {} };
  var juniorHours    = { sheet: juniorSheet, column: newColumns[2], hoursByRow: {} };
  var seniorHours    = { sheet: seniorSheet, column: newColumns[3], hoursByRow: {} };
  var officerHours   = { sheet: ofSheet,     column: newColumns[4], hoursByRow: {} };

  var names = getData(manualHoursSheet, manualHoursStartRow, 1, 0, 2);
  // Columns C (grade) and D (hours) for every entry row.
  var gradesAndHours = manualHoursSheet.getRange(manualHoursStartRow, 3, manualHoursSheet.getLastRow(), 2).getValues();

  // One read of current colors; all color changes are applied to this buffer.
  var rowColors = names.length > 0
    ? manualHoursSheet.getRange(manualHoursStartRow, 1, names.length, 4).getBackgrounds()
    : [];

  // Span of entry rows whose grade (column C) was changed to "officer".
  var firstOfficerRow = -1, lastOfficerRow = -1;

  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    const hours = gradesAndHours[i][1];

    if (name == "") { Logger.log("name not there"); continue; }
    if (hours == "" || typeof hours != "number") {
      rowColors[i][3] = warningColor; // column D
      continue;
    }

    if (recordHours(i, officerIndex, name, hours, officerHours, rowColors)) {
      gradesAndHours[i][0] = "officer"; // buffered; written once after the loop
      if (firstOfficerRow === -1) firstOfficerRow = i;
      lastOfficerRow = i;
      continue;
    }

    switch (gradesAndHours[i][0].toString().toLowerCase()) {
      case "freshman":
        if (freshmanIndex == null) freshmanIndex = buildNameIndex(getData(freshSheet, peopleSheetStartingRow, 1, 1, 3));
        recordHours(i, freshmanIndex, name, hours, freshmanHours, rowColors);
        break;
      case "sophomore":
        if (sophomoreIndex == null) sophomoreIndex = buildNameIndex(getData(sophSheet, peopleSheetStartingRow, 1, 1, 3));
        recordHours(i, sophomoreIndex, name, hours, sophomoreHours, rowColors);
        break;
      case "junior":
        if (juniorIndex == null) juniorIndex = buildNameIndex(getData(juniorSheet, peopleSheetStartingRow, 1, 1, 3));
        recordHours(i, juniorIndex, name, hours, juniorHours, rowColors);
        break;
      case "senior":
        if (seniorIndex == null) seniorIndex = buildNameIndex(getData(seniorSheet, peopleSheetStartingRow, 1, 1, 3));
        recordHours(i, seniorIndex, name, hours, seniorHours, rowColors);
        break;
      default:
        rowColors[i][2] = warningColor; // column C
        continue;
    }
  }

  // ---- Batched writes (replaces ~2 API calls per person) ----

  // Hours: one setValues() per grade sheet that had matches.
  [freshmanHours, sophomoreHours, juniorHours, seniorHours, officerHours].forEach(writeBufferedHours);

  // Grade column: one setValues() covering only the rows that changed.
  if (firstOfficerRow !== -1) {
    var gradeColumn = gradesAndHours.slice(firstOfficerRow, lastOfficerRow + 1).map(row => [row[0]]);
    manualHoursSheet.getRange(firstOfficerRow + manualHoursStartRow, 3, gradeColumn.length, 1).setValues(gradeColumn);
  }

  // Colors: one setBackgrounds() for the whole block.
  if (names.length > 0) {
    manualHoursSheet.getRange(manualHoursStartRow, 1, names.length, 4).setBackgrounds(rowColors);
  }

  isReady.setValue(true);
}

// O(1) lookup; records the hours and the entry row's color in memory instead of
// calling the Spreadsheet API. Returns true if the person was found.
function recordHours(entryRow, nameIndex, name, hours, hourBuffer, rowColors) {
  const rosterRow = nameIndex.has(name) ? nameIndex.get(name) : -1;
  const color = rosterRow == -1 ? didntFindPplColor : foundPplColour;
  for (let col = 0; col < 4; col++) rowColors[entryRow][col] = color;
  if (rosterRow == -1) return false;
  hourBuffer.hoursByRow[rosterRow + peopleSheetStartingRow] = hours;
  return true;
}

//picture buttons
// Settings cells hold [pickRandomly, currentPictureIndex].
function button() {
  var pictures = mohFunctionData.getDrawings()
  var [pickRandomly, currentPicture] = mohFunctionData.getRange(4,19,2).getValues().flat();
  pictures[currentPicture].setWidth(1)
  pictures[currentPicture].setHeight(1)

  var nextPicture
  main()

  if (pickRandomly == true) {
    var randomPicture = randomInt(0, pictures.length - 1)
    nextPicture = pictures[randomPicture]
    mohFunctionData.getRange(5,19).setValue(randomPicture)
  } else {
    nextPicture = pictures[currentPicture+1]
    mohFunctionData.getRange(5,19).setValue(currentPicture+1)
    if (currentPicture+2 == pictures.length) {
      mohFunctionData.getRange(4,19).setValue(true)
    }
  }

  nextPicture.setWidth(190)
  nextPicture.setHeight(116)
}

function button2() { //couldve just combined button functions imo                     //could've yes.... couldve had one function with sheet parameter and then 2 functions to set up the pic buttons
  var pictures = manualHoursSheet.getDrawings()
  var [pickRandomly, currentPicture] = manualHoursSheet.getRange(3,7,2).getValues().flat();
  pictures[currentPicture].setWidth(1)
  pictures[currentPicture].setHeight(1)

  var nextPicture
  manuealHour();

  if (pickRandomly == true) {
    var randomPicture = randomInt(0, pictures.length - 1)
    nextPicture = pictures[randomPicture]
    manualHoursSheet.getRange(4,7).setValue(randomPicture)
  } else {
    nextPicture = pictures[currentPicture+1]
    manualHoursSheet.getRange(4,7).setValue(currentPicture+1)
    if (currentPicture+2 == pictures.length) {
      manualHoursSheet.getRange(3,7).setValue(true)
    }
  }

  //nextPicture.setPosition(2, 5, 5, 5)
  nextPicture.setWidth(190)
  nextPicture.setHeight(116)
}

//for displaying a random picture (inclusive of min and max)
function randomInt(min, max) {
  const low = Math.ceil(min);
  const high = Math.floor(max);
  return Math.floor(Math.random() * (high - low + 1)) + low;
}
