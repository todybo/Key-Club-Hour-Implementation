const functionSpreadsheet = SpreadsheetApp.getActiveSpreadsheet() // "Function" spreadsheet (this one)
const settingsSheet = functionSpreadsheet.getSheetById(1908298872); // Settings page of Function
const eventDataSheet = functionSpreadsheet.getSheetById(891029086); // Event data page of Function
const masterOfHours = SpreadsheetApp.openById(settingsSheet.getRange("A7").getValue()); // Master of Hours spreadsheet
const freshmanHourSheet  = masterOfHours.getSheetById(settingsSheet.getRange("B7").getValue()); // Freshman page of Master of Hours
const sophomoreHourSheet = masterOfHours.getSheetById(settingsSheet.getRange("C7").getValue()); // Sophomores page of Master of Hours
const juniorHourSheet    = masterOfHours.getSheetById(settingsSheet.getRange("D7").getValue()); // Juniors page of Master of Hours
const seniorHourSheet    = masterOfHours.getSheetById(settingsSheet.getRange("E7").getValue()); // Seniors page of Master of Hours
const officerHourSheet   = masterOfHours.getSheetById(settingsSheet.getRange("F7").getValue()); // Officers page of Master of Hours
const officerNameCol = settingsSheet.getRange("A11").getValue(); // column where officer names start

const hourSheets = [freshmanHourSheet, sophomoreHourSheet, juniorHourSheet, seniorHourSheet, officerHourSheet];
const eventDataColumns = [1, 4, 7, 10, 13]; // first column of each grade's block on the event data sheet (same order as hourSheets)
const eventInfoRow = 2;                     // row holding [hours, date, event name] on the event data sheet

const foundColor = "#b7e1cd";
const notFoundColor = "#ea9999";

/**
 * Dynamically finds the starting row and column(s) for names in a sheet
 * by searching for header cells like "First Name", "First", "Last Name", "Last", or "Name".
 * Returns an object: { startRow, firstCol, lastCol }
 */
function getNameStartInfo(sheet, searchCol) {
  var defaultCol = searchCol || 1;
  var rowsToSearch = Math.min(sheet.getLastRow(), 20);
  if (rowsToSearch < 1) {
    return { startRow: 2, firstCol: defaultCol, lastCol: defaultCol + 1 };
  }

  var cells = sheet.getRange(1, 1, rowsToSearch, Math.min(sheet.getLastColumn(), 15)).getValues();

  var headerRow = -1;
  var firstNameCol = -1;
  var lastNameCol = -1;

  for (let r = 0; r < cells.length; r++) {
    for (let c = 0; c < cells[r].length; c++) {
      let cellText = (cells[r][c] || "").toString().trim().toLowerCase();

      if (cellText === "first name" || cellText === "first" || cellText === "first name / last name" || cellText === "name" || cellText === "first/last name") {
        headerRow = r + 1;    // 1-based row
        firstNameCol = c + 1; // 1-based col

        if (c + 1 < cells[r].length) {
          let nextCellText = (cells[r][c + 1] || "").toString().trim().toLowerCase();
          if (nextCellText === "last name" || nextCellText === "last" || nextCellText === "") {
            lastNameCol = c + 2;
          }
        }
        break;
      }
    }
    if (headerRow !== -1) break;
  }

  // Fallback if no explicit "First Name" header cell is found
  if (headerRow === -1) {
    var fallbackRow = 2;
    if (sheet.getLastRow() >= 5) {
      let row1Text = (sheet.getRange(1, defaultCol).getValue() || "").toString();
      let row5Text = (sheet.getRange(5, defaultCol).getValue() || "").toString();
      if (row1Text !== "" && row5Text !== "") {
        fallbackRow = 5;
      }
    }
    return { startRow: fallbackRow, firstCol: defaultCol, lastCol: defaultCol + 1 };
  }

  return {
    startRow: headerRow + 1, // Names start on the row directly UNDER "First Name"
    firstCol: firstNameCol,
    lastCol: lastNameCol !== -1 ? lastNameCol : firstNameCol + 1
  };
}

function getNameStartRow(sheet, searchCol) {
  return getNameStartInfo(sheet, searchCol).startRow;
}

function main() {
  console.log("=== MAIN STARTED ===");

  var isReady = eventDataSheet.getRange(5, 18);
  var readyValue = isReady.getValue();
  console.log("isReady value:", readyValue, "Type:", typeof readyValue);

  if (readyValue == true) {
    console.log("isReady is true - proceeding");
    isReady.setValue(false);
  } else {
    console.log("isReady is NOT true - EXITING EARLY");
    return;
  }

  console.log("Processing", hourSheets.length, "sheets");
  for (let i = 0; i < hourSheets.length; i++) {
    console.log("Processing sheet", i, hourSheets[i].getName());
    matchAndGrantHours(hourSheets[i], eventDataColumns[i]);
  }

  console.log("Setting isReady back to true");
  isReady.setValue(true);
  console.log("=== MAIN COMPLETE ===");
}

// Lowercases one part of a name and strips "(nicknames)". Callers then remove
// spaces/punctuation from the combined result.
function cleanNamePart(text) {
  return text.toLowerCase().replace(/\([^)]*\)/g, "");
}

/**
 * Reads a block of names and returns them cleaned for matching.
 *   columnCount 1 → one "full name" column
 *   columnCount 2 → first name + last name in adjacent columns
 *   columnCount 3 → first name in col 1, last name in col 3
 * rowsToSkipAtEnd drops that many rows from the bottom (e.g. a totals row).
 * If startRow is missing/0 it is detected from the "First Name" header.
 */
function readCleanNames(sheet, startRow, startCol, rowsToSkipAtEnd, columnCount) {
  if (!startRow || startRow <= 0) {
    startRow = getNameStartRow(sheet, startCol);
  }

  console.log(`\n🔍 [readCleanNames ENTER] Sheet: "${sheet.getName()}"`);
  console.log(`   Parameters: startRow=${startRow}, startCol=${startCol}, rowsToSkipAtEnd=${rowsToSkipAtEnd}, columnCount=${columnCount}`);

  var rowCount = sheet.getLastRow() - (rowsToSkipAtEnd + startRow - 1);
  console.log(`   Sheet.getLastRow() = ${sheet.getLastRow()}`);
  console.log(`   Calculated rowCount = ${rowCount}`);

  if (rowCount <= 0) {
    console.log(`   ⚠️ No rows to read, returning empty array`);
    return [];
  }

  var range = sheet.getRange(startRow, startCol, rowCount, columnCount);
  console.log(`   Reading range: ${range.getA1Notation()}`);

  var rawRows = range.getValues();
  console.log(`   ✅ Read ${rawRows.length} rows`);

  var cleanedNames = [];

  if (columnCount == 1) {
    console.log(`   Processing mode: Single column`);
    for (let i = 0; i < rawRows.length; i++) {
      let original = rawRows[i].toString();
      let cleaned = cleanNamePart(original).replace(/[\W_]+/g, "");
      cleanedNames[i] = cleaned;
      if (i < 3) console.log(`   Row ${i}: "${original}" → "${cleaned}"`);
    }
  } else if (columnCount == 2) {
    console.log(`   Processing mode: Two columns (col1 + col2)`);
    for (let i = 0; i < rawRows.length; i++) {
      let firstName = (rawRows[i][0] || "").toString();
      let lastName = (rawRows[i][1] || "").toString();
      let cleaned = (cleanNamePart(firstName) + cleanNamePart(lastName)).replace(/[\W_]+/g, "");
      cleanedNames[i] = cleaned;
      if (i < 3) console.log(`   Row ${i}: "${firstName} + ${lastName}" → "${cleaned}"`);
    }
  } else {
    console.log(`   Processing mode: Three columns (col1 + col3)`);
    for (let i = 0; i < rawRows.length; i++) {
      let firstName = (rawRows[i][0] || "").toString();
      let lastName = (rawRows[i][2] || "").toString();
      let cleaned = (cleanNamePart(firstName) + cleanNamePart(lastName)).replace(/[\W_]+/g, "");
      cleanedNames[i] = cleaned;
      if (i < 3) console.log(`   Row ${i}: "${firstName} + ${lastName}" → "${cleaned}"`);
    }
  }

  console.log(`   ✅ Processed ${cleanedNames.length} items`);
  console.log(`🔚 [readCleanNames EXIT]\n`);
  return cleanedNames;
}

// Matches the names entered for one grade on the event data sheet against that
// grade's roster in Master of Hours, colors each entry, then grants hours.
function matchAndGrantHours(hourSheet, eventDataColumn) {
  console.log(`\n========== 🔄 MATCH START ==========`);
  console.log(`Hour Sheet: "${hourSheet.getName()}"`);
  console.log(`Event Data Column: ${eventDataColumn}`);

  var rosterNameCol = 1;
  var rosterColumnCount = 3;
  if (hourSheet == officerHourSheet) {
    rosterNameCol = officerNameCol;
    rosterColumnCount = 1;
    console.log(`   ℹ️ Officer sheet detected: startCol=${officerNameCol}, columnCount=1`);
  }

  // Dynamically locate where names start under "First Name"
  var rosterStartRow = getNameStartRow(hourSheet, rosterNameCol);
  console.log(`   ℹ️ Dynamic name start row for roster "${hourSheet.getName()}": Row ${rosterStartRow}`);

  // Get roster names
  console.log(`\n📚 Getting ROSTER names from "${hourSheet.getName()}"...`);
  var rosterNames = readCleanNames(hourSheet, rosterStartRow, rosterNameCol, 1, rosterColumnCount);
  console.log(`   Roster count: ${rosterNames.length}`);

  // name -> roster row index (last occurrence wins, as before)
  console.log(`\n🔨 Creating roster index...`);
  var rosterIndex = {};
  for (let j = 0; j < rosterNames.length; j++) {
    rosterIndex[rosterNames[j]] = j;
  }

  // For Officer sheet, also populate first-name aliases from raw roster entries
  if (hourSheet == officerHourSheet && rosterNames.length > 0) {
    var rawOfficerNames = hourSheet.getRange(rosterStartRow, rosterNameCol, rosterNames.length, 1).getValues();
    for (let j = 0; j < rawOfficerNames.length; j++) {
      let fullName = (rawOfficerNames[j][0] || "").toString().trim();
      if (fullName) {
        let firstName = fullName.split(/\s+/)[0].toLowerCase().replace(/\([^)]*\)/g, "").replace(/[\W_]+/g, "");
        if (firstName && rosterIndex[firstName] === undefined) {
          rosterIndex[firstName] = j;
          console.log(`   ℹ️ Officer first-name alias added: "${firstName}" -> index ${j}`);
        }
      }
    }
  }

  console.log(`   Roster index created with ${Object.keys(rosterIndex).length} entries`);

  // Get entered names dynamically
  var entryStartRow = getNameStartRow(eventDataSheet, eventDataColumn);
  console.log(`\n📋 Getting ENTERED names from event data column ${eventDataColumn} starting at row ${entryStartRow}...`);
  var enteredNames = readCleanNames(eventDataSheet, entryStartRow, eventDataColumn, 0, 2);
  console.log(`   Entered names count: ${enteredNames.length}`);

  // Batch get reduction flags (truthy → person gets -2 instead of the event's hours)
  console.log(`\n📊 Getting reduction flags from column ${eventDataColumn+2}...`);
  var reductionFlags = eventDataSheet.getRange(entryStartRow, eventDataColumn+2, enteredNames.length, 1).getValues().flat();
  console.log(`   Reduction flags count: ${reductionFlags.length}`);
  if (reductionFlags.length > 0) {
    console.log(`   First 5 reduction flags: [${reductionFlags.slice(0, 5).join(', ')}]`);
  }

  // Keep only rows that actually have a name
  console.log(`\n🔍 Filtering out empty rows...`);
  var filledRows = [];       // offsets from entryStartRow
  var filledNames = [];
  var filledReductions = [];

  for (let i = 0; i < enteredNames.length; i++) {
    if (enteredNames[i] != "") {
      filledRows.push(i);
      filledNames.push(enteredNames[i]);
      filledReductions.push(reductionFlags[i]);
    }
  }
  console.log(`   Found ${filledRows.length} rows with data out of ${enteredNames.length} total`);

  // Match each entered name against the roster
  console.log(`\n✅ Processing ${filledNames.length} valid names...`);
  var matchedRosterRows = [];
  var matchedReductions = [];
  var rowColors = [];

  var matchCount = 0;
  var noMatchCount = 0;

  for (let i = 0; i < filledNames.length; i++) {
    var name = filledNames[i];
    var rosterRow = rosterIndex[name];

    if (rosterRow === undefined) {
      noMatchCount++;
      rowColors.push([notFoundColor]);
      if (noMatchCount <= 5) {
        console.log(`   ❌ Row ${filledRows[i] + entryStartRow}: "${name}" NOT FOUND in roster`);
      }
    } else {
      matchCount++;
      matchedRosterRows.push(rosterRow);
      matchedReductions.push(filledReductions[i]);
      rowColors.push([foundColor]);
      if (matchCount <= 5) {
        console.log(`   ✅ Row ${filledRows[i] + entryStartRow}: "${name}" FOUND at index ${rosterRow} (reduction: ${filledReductions[i]})`);
      }
    }
  }

  console.log(`\n📊 MATCH SUMMARY:`);
  console.log(`   Total matches: ${matchCount}`);
  console.log(`   Total non-matches: ${noMatchCount}`);
  console.log(`   Match rate: ${Math.round(matchCount/filledNames.length*100)}%`);

  if (matchedRosterRows.length > 0) {
    console.log(`\n📝 Matched roster rows: [${matchedRosterRows.join(', ')}]`);
    console.log(`📝 Reduction flags: [${matchedReductions.join(', ')}]`);
  }

  // Color each run of consecutive filled rows with one setBackgrounds() call
  console.log(`\n🎨 Batch updating colors...`);
  if (filledRows.length > 0) {
    var i = 0;
    var runCount = 0;
    while (i < filledRows.length) {
      var runStart = i;

      while (i + 1 < filledRows.length && filledRows[i + 1] == filledRows[i] + 1) {
        i++;
      }

      var runEnd = i;
      var runLength = runEnd - runStart + 1;
      var firstSheetRow = filledRows[runStart] + entryStartRow;

      console.log(`   Run ${++runCount}: Coloring rows ${firstSheetRow}-${firstSheetRow + runLength - 1} (${runLength} rows)`);

      var runColors = rowColors.slice(runStart, runEnd + 1);
      eventDataSheet.getRange(firstSheetRow, eventDataColumn, runLength, 1).setBackgrounds(runColors);

      i++;
    }
    console.log(`   ✅ Applied ${runCount} color batches`);
  }

  console.log(`\n📞 Calling grantHour with ${matchedRosterRows.length} people...`);
  grantHour(matchedRosterRows, matchedReductions, hourSheet, rosterStartRow);

  console.log(`========== 🔚 MATCH END ==========\n`);
}

// Inserts a new event column on hourSheet and writes each matched person's hours.
function grantHour(rosterRows, reductionFlags, hourSheet, rosterStartRow) {
  console.log(`\n=== GRANTHOUR for ${hourSheet.getName()} ===`);
  console.log(`rosterRows:`, rosterRows);
  console.log(`reductionFlags:`, reductionFlags);

  if (!rosterStartRow) {
    let isOfficer = (hourSheet == officerHourSheet);
    rosterStartRow = getNameStartRow(hourSheet, isOfficer ? officerNameCol : 1);
  }

  var lastColumnBefore = hourSheet.getMaxColumns();
  console.log(`BEFORE insertion - Last column: ${lastColumnBefore}`);
  var trailingColumns; // # of columns kept to the right of the newest event (from settings)

  if (hourSheet == officerHourSheet) {
    trailingColumns = settingsSheet.getRange("C9").getValue();
    console.log(`Officer trailing columns (C9): ${trailingColumns}`);
  } else {
    trailingColumns = settingsSheet.getRange("B9").getValue();
    console.log(`Regular trailing columns (B9): ${trailingColumns}`);
  }

  var insertAt = lastColumnBefore - trailingColumns + 1;

  if (trailingColumns === 0) {
    hourSheet.insertColumnAfter(lastColumnBefore);
  } else if (insertAt >= hourSheet.getMaxColumns()) {
    hourSheet.insertColumns(hourSheet.getMaxColumns(), 1);
  } else {
    hourSheet.insertColumnBefore(insertAt);
  }

  var newColumn = hourSheet.getMaxColumns() - trailingColumns;

  console.log(`New column inserted at: ${newColumn}`);
  console.log(`Column ${newColumn} is now between columns ${newColumn-1} and ${newColumn+1}`);

  var [eventHours, eventDate, eventName] = eventDataSheet.getRange(eventInfoRow, eventDataColumns[4]+3, 1, 3).getValues()[0];
  console.log(`event info:`, eventHours, eventDate, eventName);

  var headerRange = hourSheet.getRange(settingsSheet.getRange("E9").getValue(), newColumn, settingsSheet.getRange("F9").getValue()).merge();
  headerRange.setValue(eventName);

  var dateRow = settingsSheet.getRange("D9").getValue();
  console.log(`Date row from settings D9: ${dateRow}`);

  var dateCell = hourSheet.getRange(dateRow, newColumn);
  dateCell.setValue(eventDate);
  dateCell.setNumberFormat("m/d");

  if (rosterRows.length > 0) {
    var sheetRows = [];
    var hourValues = [];

    for (let i = 0; i < rosterRows.length; i++) {
      var sheetRow = rosterRows[i] + rosterStartRow;
      var hours = reductionFlags[i] ? -2 : eventHours;
      sheetRows.push(sheetRow);
      hourValues.push(hours);
      console.log(`Person ${i}: roster index ${rosterRows[i]} → sheet row ${sheetRow}, hour = ${hours}`);
    }

    var firstRow = Math.min(...sheetRows);
    var lastRow = Math.max(...sheetRows);
    var rowCount = lastRow - firstRow + 1;

    console.log(`Writing to rows ${firstRow}-${lastRow} (${rowCount} rows total)`);

    var columnValues = new Array(rowCount);
    for (let i = 0; i < rowCount; i++) {
      columnValues[i] = [null];
    }

    for (let i = 0; i < sheetRows.length; i++) {
      columnValues[sheetRows[i] - firstRow][0] = hourValues[i];
      console.log(`  Placing ${hourValues[i]} at offset ${sheetRows[i] - firstRow}`);
    }

    hourSheet.getRange(firstRow, newColumn, rowCount, 1).setValues(columnValues);
    console.log(`✅ Wrote ${sheetRows.length} values to column ${newColumn}`);
  } else {
    console.log(`No values to write for ${hourSheet.getName()}`);
  }

  console.log(`=== END GRANTHOUR for ${hourSheet.getName()} ===\n`);
}
