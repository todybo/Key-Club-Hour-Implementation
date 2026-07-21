const mohFunction = SpreadsheetApp.getActiveSpreadsheet() // Function
const mohSettings = mohFunction.getSheetById(1908298872); // Settings page of Function
const mohFunctionData = mohFunction.getSheetById(891029086); // Event data page of Function
const moh = SpreadsheetApp.openById(mohSettings.getRange("A7").getValue()); // Master of Hours
const freshSheet = moh.getSheetById(mohSettings.getRange("B7").getValue()); // Freshman page of Master of Hours
const sophSheet = moh.getSheetById(mohSettings.getRange("C7").getValue()); // Sophomores page of Master of Hours
const juniorSheet = moh.getSheetById(mohSettings.getRange("D7").getValue()); // Juniors page of Master of Hours
const seniorSheet = moh.getSheetById(mohSettings.getRange("E7").getValue()); // Seniors page of Master of Hours
const ofSheet = moh.getSheetById(mohSettings.getRange("F7").getValue()); // Officers page of Master of Hours 
const officerStartingCol = mohSettings.getRange("A11").getValue();


const hourSheets = [freshSheet, sophSheet, juniorSheet, seniorSheet, ofSheet];
const hourSheetColumnDataPosition = [1, 4, 7, 10, 13];
const peopleSheetStartingRow = mohSettings.getRange("A9").getValue();
const dataSheetStartingRow = 2;

const foundPplColour = "#b7e1cd";
const didntFindPplColor = "#ea9999";

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
      let val3 = (rawData[i][2] || "").toString(); // if merged like in 2025-2026, moh, then use 2. else use 1
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
    console.log(`   ℹ️ Officer sheet detected: startingColumn=2, amtColumns=1`);
  }
  
  // Get reference data
  console.log(`\n📚 Getting REFERENCE data from "${referenceSheet.getName()}"...`);
  var referenceData = getData(referenceSheet, peopleSheetStartingRow, isOfficerSheet, 1, amtColumns);
  console.log(`   Reference data count: ${referenceData.length}`);
  
  // Create lookup object
  console.log(`\n🔨 Creating lookup object...`);
  var referenceLookup = {};
  for (let j = 0; j < referenceData.length; j++) {
    referenceLookup[referenceData[j]] = j;
  }
  console.log(`   Lookup object created with ${Object.keys(referenceLookup).length} entries`);
  
  // Get lookup data
  console.log(`\n📋 Getting LOOKUP data from main data sheet at position ${dataSheetPosition}...`);
  var lookupData = getData(mohFunctionData, dataSheetStartingRow, dataSheetPosition, 0, 2);
  console.log(`   Lookup data count: ${lookupData.length}`);
  
  // Batch get reduction values
  console.log(`\n📊 Getting reduction values from column ${dataSheetPosition+2}...`);
  var reductionRange = mohFunctionData.getRange(dataSheetStartingRow, dataSheetPosition+2, lookupData.length, 1);
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
        console.log(`   ❌ Row ${validRows[i] + dataSheetStartingRow}: "${name}" NOT FOUND in reference`);
      }
    } else {
      matchCount++;
      onBothIndexs.push(pp);
      getHrReduce.push(validReductionValues[i]);
      validRowColors.push([foundPplColour]);
      if (matchCount <= 5) {
        console.log(`   ✅ Row ${validRows[i] + dataSheetStartingRow}: "${name}" FOUND at index ${pp} (reduction: ${validReductionValues[i]})`);
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
      var startRow = validRows[startIdx] + dataSheetStartingRow;
      
      console.log(`   Batch ${++batchCount}: Coloring rows ${startRow}-${startRow + batchLength - 1} (${batchLength} rows)`);
      
      var batchColors = validRowColors.slice(startIdx, endIdx + 1);
      mohFunctionData.getRange(startRow, dataSheetPosition, batchLength, 1).setBackgrounds(batchColors);
      
      i++;
    }
    console.log(`   ✅ Applied ${batchCount} color batches`);
  }
  
  console.log(`\n📞 Calling grantHour with ${onBothIndexs.length} people...`);
  grantHour(onBothIndexs, getHrReduce, referenceSheet);
  
  console.log(`========== 🔚 COMPARER END ==========\n`);
}

function grantHour(indexArray, hrReductionArray, hourSheet) { //CP VIBE CODE
  console.log(`\n=== GRANTHOUR for ${hourSheet.getName()} ===`);
  console.log(`indexArray:`, indexArray);
  console.log(`hrReductionArray:`, hrReductionArray);
  
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
  
// Insert the column safely without going out of bounds
  var targetColumn = beforeLastCol - offset + 1;

  if (offset === 0) 
  {
    hourSheet.insertColumnAfter(beforeLastCol);
  }
  else if (targetColumn >= hourSheet.getMaxColumns()) {
  // If we are pushing past the absolute edge, append it directly to the end
    hourSheet.insertColumns(hourSheet.getMaxColumns(), 1);
  } else {
  // Otherwise, use your standard relative insertion
    hourSheet.insertColumnBefore(targetColumn);
  }
  
  // After insertion, new column is at (lastColumn - offset)
  currentColumn = hourSheet.getMaxColumns() - offset;
  
  console.log(`New column inserted at: ${currentColumn}`);
  console.log(`Column ${currentColumn} is now between columns ${currentColumn-1} and ${currentColumn+1}`);
  
  // Get all event data at once
  var eventData = mohFunctionData.getRange(dataSheetStartingRow, hourSheetColumnDataPosition[4]+3, 1, 3).getValues()[0];
  console.log(`eventData:`, eventData);
  
  // Setup headers
  hourSheet.getRange(1, currentColumn, 2);
  var headerRange = hourSheet.getRange(mohSettings.getRange("E9").getValue(), currentColumn, mohSettings.getRange("F9").getValue()).merge(); // at what # row does event name start? how many rows does it span?
  headerRange.setValue(eventData[2]);
  
  var dateRow = mohSettings.getRange("D9").getValue();
  console.log(`Date row from settings D9: ${dateRow}`);
  
  var dateCell = hourSheet.getRange(dateRow, currentColumn);
  dateCell.setValue(eventData[1]);
  dateCell.setNumberFormat("m/d");
  
  // Batch write all hour values at once
  if (indexArray.length > 0) {
    var rows = [];
    var values = [];
    
    for (let i = 0; i < indexArray.length; i++) {
      var rowNum = indexArray[i] + peopleSheetStartingRow;
      var hourValue = hrReductionArray[i] ? -2 : eventData[0];
      rows.push(rowNum);
      values.push([hourValue]);
      console.log(`Person ${i}: ref index ${indexArray[i]} → sheet row ${rowNum}, hour = ${hourValue}`);
    }
    
    // Find the range that covers all rows
    var minRow = Math.min(...rows);
    var maxRow = Math.max(...rows);
    var rowCount = maxRow - minRow + 1;
    
    console.log(`Writing to rows ${minRow}-${maxRow} (${rowCount} rows total)`);
    
    // Create a 2D array for all rows
    var allValues = new Array(rowCount);
    for (let i = 0; i < rowCount; i++) {
      allValues[i] = new Array(1).fill(null);
    }
    
    // Place values at correct positions
    for (let i = 0; i < rows.length; i++) {
      allValues[rows[i] - minRow][0] = values[i][0];
      console.log(`  Placing ${values[i][0]} at offset ${rows[i] - minRow}`);
    }
    
    // Batch update the entire column
    hourSheet.getRange(minRow, currentColumn, rowCount, 1).setValues(allValues);
    console.log(`✅ Wrote ${rows.length} values to column ${currentColumn}`);
  } else {
    console.log(`No values to write for ${hourSheet.getName()}`);
  }
  
  console.log(`=== END GRANTHOUR for ${hourSheet.getName()} ===\n`);
}