const mohFunction = SpreadsheetApp.getActiveSpreadsheet() // Function
const mohSettings = mohFunction.getSheetById(1908298872); // Settings page of Function
const mohFunctionData = mohFunction.getSheetById(891029086); // Event data page of Function
const moh = SpreadsheetApp.openById(mohSettings.getRange("A7").getValue()); // Master of Hours
const freshSheet = moh.getSheetById(mohSettings.getRange("B7").getValue()); // Freshman page of Master of Hours
const sophSheet = moh.getSheetById(mohSettings.getRange("C7").getValue()); // Sophomores page of Master of Hours
const juniorSheet = moh.getSheetById(mohSettings.getRange("D7").getValue()); // Juniors page of Master of Hours
const seniorSheet = moh.getSheetById(mohSettings.getRange("E7").getValue()); // Seniors page of Master of Hours
const ofSheet = moh.getSheetById(mohSettings.getRange("F7").getValue()); // Officers page of Master of Hours 
const officerStartingCol = mohSettings.getRange("A11").getValue(); // officer starting column

const hourSheets = [freshSheet, sophSheet, juniorSheet, seniorSheet, ofSheet];
const hourSheetColumnDataPosition = [1, 4, 7, 10, 13];
const dataSheetStartingRow = 2;

const foundPplColour = "#b7e1cd";
const didntFindPplColor = "#ea9999";

/**
 * Dynamically finds the starting row and column(s) for names in a sheet
 * by searching for header cells like "First Name", "First", "Last Name", "Last", or "Name".
 * Returns an object: { startRow, firstCol, lastCol }
 */
function getNameStartInfo(sheet, searchCol) {
  var defaultCol = searchCol || 1;
  var maxSearchRows = Math.min(sheet.getLastRow(), 20);
  if (maxSearchRows < 1) {
    return { startRow: 2, firstCol: defaultCol, lastCol: defaultCol + 1 };
  }

  var searchRange = sheet.getRange(1, 1, maxSearchRows, Math.min(sheet.getLastColumn(), 15));
  var values = searchRange.getValues();

  var foundHeaderRow = -1;
  var foundFirstCol = -1;
  var foundLastCol = -1;

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      let cellText = (values[r][c] || "").toString().trim().toLowerCase();
      
      if (cellText === "first name" || cellText === "first" || cellText === "first name / last name" || cellText === "name" || cellText === "first/last name") {
        foundHeaderRow = r + 1; // 1-based row
        foundFirstCol = c + 1;  // 1-based col
        
        if (c + 1 < values[r].length) {
          let nextCell = (values[r][c + 1] || "").toString().trim().toLowerCase();
          if (nextCell === "last name" || nextCell === "last" || nextCell === "") {
            foundLastCol = c + 2;
          }
        }
        break;
      }
    }
    if (foundHeaderRow !== -1) break;
  }

  // Fallback if no explicit "First Name" header cell is found
  if (foundHeaderRow === -1) {
    var fallbackRow = 2;
    if (sheet.getLastRow() >= 5) {
      let row1Val = (sheet.getRange(1, defaultCol).getValue() || "").toString();
      let row5Val = (sheet.getRange(5, defaultCol).getValue() || "").toString();
      if (row1Val !== "" && row5Val !== "") {
        fallbackRow = 5;
      }
    }
    return { startRow: fallbackRow, firstCol: defaultCol, lastCol: defaultCol + 1 };
  }

  return {
    startRow: foundHeaderRow + 1, // Names start on the row directly UNDER "First Name"
    firstCol: foundFirstCol,
    lastCol: foundLastCol !== -1 ? foundLastCol : foundFirstCol + 1
  };
}

function getNameStartRow(sheet, searchCol) {
  return getNameStartInfo(sheet, searchCol).startRow;
}

function main() {
  console.log("=== MAIN STARTED ===");
  
  var isReady = mohFunctionData.getRange(5, 18);
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
    comparer(hourSheets[i], hourSheetColumnDataPosition[i]);
  }

  console.log("Setting isReady back to true");
  isReady.setValue(true);
  console.log("=== MAIN COMPLETE ===");
}

function getData(sheet, startingRow, startingColumn, lastRowOffset, isOneColumn) {
  if (!startingRow || startingRow <= 0) {
    startingRow = getNameStartRow(sheet, startingColumn);
  }
  
  console.log(`\n🔍 [getData ENTER] Sheet: "${sheet.getName()}"`);
  console.log(`   Parameters: startingRow=${startingRow}, startingColumn=${startingColumn}, lastRowOffset=${lastRowOffset}, isOneColumn=${isOneColumn}`);
  
  var numRows = sheet.getLastRow() - (lastRowOffset + startingRow - 1);
  console.log(`   Sheet.getLastRow() = ${sheet.getLastRow()}`);
  console.log(`   Calculated numRows = ${numRows}`);
  
  if (numRows <= 0) {
    console.log(`   ⚠️ No rows to read, returning empty array`);
    return [];
  }
  
  var range = sheet.getRange(startingRow, startingColumn, numRows, isOneColumn);
  console.log(`   Reading range: ${range.getA1Notation()}`);
  
  var rawData = range.getValues();
  console.log(`   ✅ Read ${rawData.length} rows`);
  
  var newData = [];
  
  if (isOneColumn == 1) {
    console.log(`   Processing mode: Single column`);
    for (let i = 0; i < rawData.length; i++) {
      let original = rawData[i].toString();
      let bob = original.toLowerCase().replace(/\([^)]*\)/g, "");
      let cleaned = bob.replace(/[\W_]+/g, "");
      newData[i] = cleaned;
      if (i < 3) console.log(`   Row ${i}: "${original}" → "${cleaned}"`);
    }
  } else if (isOneColumn == 2) {
    console.log(`   Processing mode: Two columns (col1 + col2)`);
    for (let i = 0; i < rawData.length; i++) {
      let val1 = (rawData[i][0] || "").toString();
      let val2 = (rawData[i][1] || "").toString();
      let original = val1 + " + " + val2;
      let bob = val1.toLowerCase().replace(/\([^)]*\)/g, "") + 
                val2.toLowerCase().replace(/\([^)]*\)/g, "");
      let cleaned = bob.replace(/[\W_]+/g, "");
      newData[i] = cleaned;
      if (i < 3) console.log(`   Row ${i}: "${original}" → "${cleaned}"`);
    }
  } else {
    console.log(`   Processing mode: Three columns (col1 + col3)`);
    for (let i = 0; i < rawData.length; i++) {
      let val1 = (rawData[i][0] || "").toString();
      let val3 = (rawData[i][2] || "").toString();
      let original = val1 + " + " + val3;
      let bob = val1.toLowerCase().replace(/\([^)]*\)/g, "") + 
                val3.toLowerCase().replace(/\([^)]*\)/g, "");
      let cleaned = bob.replace(/[\W_]+/g, "");
      newData[i] = cleaned;
      if (i < 3) console.log(`   Row ${i}: "${original}" → "${cleaned}"`);
    }
  }
  
  console.log(`   ✅ Processed ${newData.length} items`);
  console.log(`🔚 [getData EXIT]\n`);
  return newData;
}

function comparer(referenceSheet, dataSheetPosition) {
  console.log(`\n========== 🔄 COMPARER START ==========`);
  console.log(`Reference Sheet: "${referenceSheet.getName()}"`);
  console.log(`Data Sheet Position: ${dataSheetPosition}`);
  
  var isOfficerSheet = 1;
  var amtColumns = 3;
  if (referenceSheet == ofSheet) {
    isOfficerSheet = officerStartingCol;
    amtColumns = 1;
    console.log(`   ℹ️ Officer sheet detected: startingColumn=${officerStartingCol}, amtColumns=1`);
  }
  
  // Dynamically locate where names start under "First Name"
  var refStartRow = getNameStartRow(referenceSheet, isOfficerSheet);
  console.log(`   ℹ️ Dynamic name start row for reference sheet "${referenceSheet.getName()}": Row ${refStartRow}`);

  // Get reference data
  console.log(`\n📚 Getting REFERENCE data from "${referenceSheet.getName()}"...`);
  var referenceData = getData(referenceSheet, refStartRow, isOfficerSheet, 1, amtColumns);
  console.log(`   Reference data count: ${referenceData.length}`);
  
  // Create lookup object
  console.log(`\n🔨 Creating lookup object...`);
  var referenceLookup = {};
  for (let j = 0; j < referenceData.length; j++) {
    referenceLookup[referenceData[j]] = j;
  }

  // For Officer sheet, also populate first-name aliases from raw roster entries
  if (referenceSheet == ofSheet && referenceData.length > 0) {
    var rawOfficerRange = referenceSheet.getRange(refStartRow, isOfficerSheet, referenceData.length, 1).getValues();
    for (let j = 0; j < rawOfficerRange.length; j++) {
      let rawStr = (rawOfficerRange[j][0] || "").toString().trim();
      if (rawStr) {
        let firstNameOnly = rawStr.split(/\s+/)[0].toLowerCase().replace(/\([^)]*\)/g, "").replace(/[\W_]+/g, "");
        if (firstNameOnly && referenceLookup[firstNameOnly] === undefined) {
          referenceLookup[firstNameOnly] = j;
          console.log(`   ℹ️ Officer first-name alias added: "${firstNameOnly}" -> index ${j}`);
        }
      }
    }
  }

  console.log(`   Lookup object created with ${Object.keys(referenceLookup).length} entries`);
  
  // Get lookup data dynamically
  var dataStartRow = getNameStartRow(mohFunctionData, dataSheetPosition);
  console.log(`\n📋 Getting LOOKUP data from main data sheet at position ${dataSheetPosition} starting at row ${dataStartRow}...`);
  var lookupData = getData(mohFunctionData, dataStartRow, dataSheetPosition, 0, 2);
  console.log(`   Lookup data count: ${lookupData.length}`);
  
  // Batch get reduction values
  console.log(`\n📊 Getting reduction values from column ${dataSheetPosition+2}...`);
  var reductionRange = mohFunctionData.getRange(dataStartRow, dataSheetPosition+2, lookupData.length, 1);
  var reductionValues = reductionRange.getValues().flat();
  console.log(`   Reduction values count: ${reductionValues.length}`);
  if (reductionValues.length > 0) {
    console.log(`   First 5 reduction values: [${reductionValues.slice(0, 5).join(', ')}]`);
  }
  
  // Find only rows with actual data
  console.log(`\n🔍 Filtering out empty rows...`);
  var validRows = [];
  var validLookupData = [];
  var validReductionValues = [];
  
  for (let i = 0; i < lookupData.length; i++) {
    if (lookupData[i] != "") {
      validRows.push(i);
      validLookupData.push(lookupData[i]);
      validReductionValues.push(reductionValues[i]);
    }
  }
  console.log(`   Found ${validRows.length} rows with data out of ${lookupData.length} total`);
  
  // Process only valid rows
  console.log(`\n✅ Processing ${validLookupData.length} valid names...`);
  var onBothIndexs = [];
  var getHrReduce = [];
  var validRowColors = [];
  
  var matchCount = 0;
  var noMatchCount = 0;
  
  for (let i = 0; i < validLookupData.length; i++) {
    var name = validLookupData[i];
    var pp = referenceLookup[name];
    
    if (pp === undefined) {
      noMatchCount++;
      validRowColors.push([didntFindPplColor]);
      if (noMatchCount <= 5) {
        console.log(`   ❌ Row ${validRows[i] + dataStartRow}: "${name}" NOT FOUND in reference`);
      }
    } else {
      matchCount++;
      onBothIndexs.push(pp);
      getHrReduce.push(validReductionValues[i]);
      validRowColors.push([foundPplColour]);
      if (matchCount <= 5) {
        console.log(`   ✅ Row ${validRows[i] + dataStartRow}: "${name}" FOUND at index ${pp} (reduction: ${validReductionValues[i]})`);
      }
    }
  }
  
  console.log(`\n📊 MATCH SUMMARY:`);
  console.log(`   Total matches: ${matchCount}`);
  console.log(`   Total non-matches: ${noMatchCount}`);
  console.log(`   Match rate: ${Math.round(matchCount/validLookupData.length*100)}%`);
  
  if (onBothIndexs.length > 0) {
    console.log(`\n📝 Matched indices: [${onBothIndexs.join(', ')}]`);
    console.log(`📝 Reduction values: [${getHrReduce.join(', ')}]`);
  }
  
  // Smart batch update colors
  console.log(`\n🎨 Batch updating colors...`);
  if (validRows.length > 0) {
    var i = 0;
    var batchCount = 0;
    while (i < validRows.length) {
      var startIdx = i;
      
      while (i + 1 < validRows.length && validRows[i + 1] == validRows[i] + 1) {
        i++;
      }
      
      var endIdx = i;
      var batchLength = endIdx - startIdx + 1;
      var startRow = validRows[startIdx] + dataStartRow;
      
      console.log(`   Batch ${++batchCount}: Coloring rows ${startRow}-${startRow + batchLength - 1} (${batchLength} rows)`);
      
      var batchColors = validRowColors.slice(startIdx, endIdx + 1);
      mohFunctionData.getRange(startRow, dataSheetPosition, batchLength, 1).setBackgrounds(batchColors);
      
      i++;
    }
    console.log(`   ✅ Applied ${batchCount} color batches`);
  }
  
  console.log(`\n📞 Calling grantHour with ${onBothIndexs.length} people...`);
  grantHour(onBothIndexs, getHrReduce, referenceSheet, refStartRow);
  
  console.log(`========== 🔚 COMPARER END ==========\n`);
}

function grantHour(indexArray, hrReductionArray, hourSheet, refStartRow) {
  console.log(`\n=== GRANTHOUR for ${hourSheet.getName()} ===`);
  console.log(`indexArray:`, indexArray);
  console.log(`hrReductionArray:`, hrReductionArray);
  
  if (!refStartRow) {
    let isOfficer = (hourSheet == ofSheet);
    refStartRow = getNameStartRow(hourSheet, isOfficer ? officerStartingCol : 1);
  }

  var beforeLastCol = hourSheet.getMaxColumns();
  console.log(`BEFORE insertion - Last column: ${beforeLastCol}`);
  var currentColumn = 0;
  var offset;

  if (hourSheet == ofSheet) {
    offset = mohSettings.getRange("C9").getValue();
    console.log(`Officer offset (C9): ${offset}`);
  } else {
    offset = mohSettings.getRange("B9").getValue();
    console.log(`Regular offset (B9): ${offset}`);
  }
  
  var targetColumn = beforeLastCol - offset + 1;

  if (offset === 0) {
    hourSheet.insertColumnAfter(beforeLastCol);
  } else if (targetColumn >= hourSheet.getMaxColumns()) {
    hourSheet.insertColumns(hourSheet.getMaxColumns(), 1);
  } else {
    hourSheet.insertColumnBefore(targetColumn);
  }
  
  currentColumn = hourSheet.getMaxColumns() - offset;
  
  console.log(`New column inserted at: ${currentColumn}`);
  console.log(`Column ${currentColumn} is now between columns ${currentColumn-1} and ${currentColumn+1}`);
  
  var eventData = mohFunctionData.getRange(dataSheetStartingRow, hourSheetColumnDataPosition[4]+3, 1, 3).getValues()[0];
  console.log(`eventData:`, eventData);
  
  hourSheet.getRange(1, currentColumn, 2);
  var headerRange = hourSheet.getRange(mohSettings.getRange("E9").getValue(), currentColumn, mohSettings.getRange("F9").getValue()).merge();
  headerRange.setValue(eventData[2]);
  
  var dateRow = mohSettings.getRange("D9").getValue();
  console.log(`Date row from settings D9: ${dateRow}`);
  
  var dateCell = hourSheet.getRange(dateRow, currentColumn);
  dateCell.setValue(eventData[1]);
  dateCell.setNumberFormat("m/d");
  
  if (indexArray.length > 0) {
    var rows = [];
    var values = [];
    
    for (let i = 0; i < indexArray.length; i++) {
      var rowNum = indexArray[i] + refStartRow;
      var hourValue = hrReductionArray[i] ? -2 : eventData[0];
      rows.push(rowNum);
      values.push([hourValue]);
      console.log(`Person ${i}: ref index ${indexArray[i]} → sheet row ${rowNum}, hour = ${hourValue}`);
    }
    
    var minRow = Math.min(...rows);
    var maxRow = Math.max(...rows);
    var rowCount = maxRow - minRow + 1;
    
    console.log(`Writing to rows ${minRow}-${maxRow} (${rowCount} rows total)`);
    
    var allValues = new Array(rowCount);
    for (let i = 0; i < rowCount; i++) {
      allValues[i] = new Array(1).fill(null);
    }
    
    for (let i = 0; i < rows.length; i++) {
      allValues[rows[i] - minRow][0] = values[i][0];
      console.log(`  Placing ${values[i][0]} at offset ${rows[i] - minRow}`);
    }
    
    hourSheet.getRange(minRow, currentColumn, rowCount, 1).setValues(allValues);
    console.log(`✅ Wrote ${rows.length} values to column ${currentColumn}`);
  } else {
    console.log(`No values to write for ${hourSheet.getName()}`);
  }
  
  console.log(`=== END GRANTHOUR for ${hourSheet.getName()} ===\n`);
}
