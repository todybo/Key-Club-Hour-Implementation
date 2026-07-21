/* const ss = SpreadsheetApp.openById("1sxydhYCKzykkD-KokAy88ofrLbB3PDY7GU6b2WzwW-M")
const css = SpreadsheetApp.getActiveSpreadsheet()
const dataSheet = css.getSheetById(891029086);
const fsheet = ss.getSheetById(1446750787);
const spsheet = ss.getSheetById(968806907);
const jsheet = ss.getSheetById(895219470);
const ssheet = ss.getSheetById(664476988);
const ofsheet = ss.getSheetById(1779276002);

const hourSheets = [fsheet, spsheet, jsheet, ssheet, ofsheet];
const hourSheetColumnDataPosition = [1, 4, 7, 10, 13];
const peopleSheetStartingRow = 8;
const dataSheetStartingRow = 2;

const foundPplColour = "#b7e1cd";
const didntFindPplColor = "#ea9999";

function main() {
  var isReady = dataSheet.getRange(5, 18);
  if (isReady.getValue() == true) {
    isReady.setValue(false);
  } else {
    return;
  }

  for (let i = 0; i < hourSheets.length; i++) {
    comparer(hourSheets[i], hourSheetColumnDataPosition[i]);
  }

  isReady.setValue(true);
}

function getData(sheet, startingRow, startingColumn, lastRowOffset, isOneColumn) {
  var newData = [];
  var rawData = sheet.getRange(startingRow, startingColumn, sheet.getLastRow() - (lastRowOffset + startingRow - 1), isOneColumn).getValues();

  if (isOneColumn == 1) {
    for (let i = 0; i < rawData.length; i++) {
      let bob = rawData[i].toString().toLowerCase().replace(/\([^)]*\)/g, "");
      newData[i] = bob.replace(/[\W_]+/g, "");
    }
  } else if (isOneColumn == 2) {
    for (let i = 0; i < rawData.length; i++) {
      let bob = rawData[i][0].toString().toLowerCase().replace(/\([^)]*\)/g, "") + rawData[i][1].toString().toLowerCase().replace(/\([^)]*\)/g, "");
      newData[i] = bob.replace(/[\W_]+/g, "");
    }
  } else {
    for (let i = 0; i < rawData.length; i++) {
      let bob = rawData[i][0].toString().toLowerCase().replace(/\([^)]*\)/g, "") + rawData[i][2].toString().toLowerCase().replace(/\([^)]*\)/g, "");
      newData[i] = bob.replace(/[\W_]+/g, "");
    }
  }
  return newData;
}

function comparer(referenceSheet, dataSheetPosition) { //CP VIBECODE
  var isOfficerSheet = 1;
  var amtColumns = 3
  if (referenceSheet == ofsheet) {
    isOfficerSheet = 2;
    amtColumns = 1
  }
  
  // Get reference data
  var referenceData = getData(referenceSheet, peopleSheetStartingRow, isOfficerSheet, 1, amtColumns);
  
  // Create lookup object (O(1) access instead of O(n) findIndex)
  var referenceLookup = {};
  for (let j = 0; j < referenceData.length; j++) {
    referenceLookup[referenceData[j]] = j;
  }
  
  var lookupData = getData(dataSheet, dataSheetStartingRow, dataSheetPosition, 0, 2);
  
  // Batch get reduction values
  var reductionRange = dataSheet.getRange(dataSheetStartingRow, dataSheetPosition+2, lookupData.length, 1);
  var reductionValues = reductionRange.getValues().flat();
  
  // ===== FIND ONLY ROWS WITH ACTUAL DATA =====
  var validRows = [];           // Original indices of rows with data
  var validLookupData = [];     // Names for valid rows
  var validReductionValues = []; // Reduction values for valid rows
  
  for (let i = 0; i < lookupData.length; i++) {
    if (lookupData[i] != "") {
      validRows.push(i);
      validLookupData.push(lookupData[i]);
      validReductionValues.push(reductionValues[i]);
    }
  }
  
  // ===== PROCESS ONLY VALID ROWS =====
  var onBothIndexs = [];
  var getHrReduce = [];
  var validRowColors = []; // Colors only for valid rows
  
  for (let i = 0; i < validLookupData.length; i++) {
    var pp = referenceLookup[validLookupData[i]];
    
    if (pp === undefined) {
      validRowColors.push([didntFindPplColor]);
    } else {
      onBothIndexs.push(pp);
      getHrReduce.push(validReductionValues[i]);
      validRowColors.push([foundPplColour]);
    }
  }
  
  // ===== SMART BATCH UPDATE COLORS (only for valid rows) =====
  if (validRows.length > 0) {
    var i = 0;
    while (i < validRows.length) {
      var startIdx = i;
      
      // Find consecutive rows
      while (i + 1 < validRows.length && validRows[i + 1] == validRows[i] + 1) {
        i++;
      }
      
      var endIdx = i;
      var batchLength = endIdx - startIdx + 1;
      
      // Batch update this consecutive block
      var startRow = validRows[startIdx] + dataSheetStartingRow;
      var batchColors = validRowColors.slice(startIdx, endIdx + 1);
      
      dataSheet.getRange(startRow, dataSheetPosition, batchLength, 1).setBackgrounds(batchColors);
      
      i++;
    }
  }
  
  console.log(referenceSheet.getName()+": ", onBothIndexs, ", ", getHrReduce);
  grantHour(onBothIndexs, getHrReduce, referenceSheet);
}

/* function comparer(referenceSheet, dataSheetPosition) { // CODY BUI CODE!!!! -cp
  var isOfficerSheet = 1;
  var amtColumns = 3
  if (referenceSheet == ofsheet) {
    isOfficerSheet = 2;
    amtColumns = 1
  }
  var referenceData = getData(referenceSheet, peopleSheetStartingRow, isOfficerSheet, 1, amtColumns);
  var lookupData = getData(dataSheet, dataSheetStartingRow, dataSheetPosition, 0, 2);

  var onBothIndexs = [];
  var getHrReduce = []; //true is reduction
  for (let i = 0; i < lookupData.length; i++) {
    if (lookupData[i] == "") {continue}

    let pp = referenceData.findIndex(name => name == lookupData[i])
    if (pp == -1) {
      dataSheet.getRange(i + dataSheetStartingRow, dataSheetPosition, 1, 3).setBackground(didntFindPplColor);
    } else {
      onBothIndexs.push(pp);
      getHrReduce.push(dataSheet.getRange(i + dataSheetStartingRow,dataSheetPosition+2).getValue());
      dataSheet.getRange(i + dataSheetStartingRow, dataSheetPosition, 1, 3).setBackground(foundPplColour);
    }
  }
  console.log(referenceSheet.getName()+": ", onBothIndexs, ", ", getHrReduce)
  grantHour(onBothIndexs, getHrReduce, referenceSheet);
} */

/* function grantHour(indexArray, hrReductionArray, hourSheet) { // CODY BUI CODE!!!! -cp
  hourSheet.insertColumnBefore(hourSheet.getLastColumn() - 2);
  var currentColumn = hourSheet.getLastColumn() - 3;

  hourSheet.getRange(1, currentColumn, 2).merge()

  var eventData = dataSheet.getRange(dataSheetStartingRow, hourSheetColumnDataPosition[4]+3, 1, 3).getValues().flat();
  var range = hourSheet.getRange(3, currentColumn, 4)
  range.merge()
  range.setValue(eventData[2])
  var datePut = hourSheet.getRange(7, currentColumn)
  datePut.setValue(eventData[1])
  datePut.setNumberFormat("m/d")

  for (let i = 0; i < indexArray.length; i++) {
    let hourValue = eventData[0]
    if (hrReductionArray[i]) {hourValue = -2}
    hourSheet.getRange(indexArray[i] + peopleSheetStartingRow, currentColumn).setValue(hourValue);
  }
} */

/* function grantHour(indexArray, hrReductionArray, hourSheet) { //CP VIBE CODE
  hourSheet.insertColumnBefore(hourSheet.getLastColumn() - 2);
  var currentColumn = hourSheet.getLastColumn() - 3;
  
  // Get all event data at once
  var eventData = dataSheet.getRange(dataSheetStartingRow, hourSheetColumnDataPosition[4]+3, 1, 3).getValues()[0];
  
  // Setup headers (single operations)
  hourSheet.getRange(1, currentColumn, 2).merge();
  var headerRange = hourSheet.getRange(3, currentColumn, 4);
  headerRange.merge();
  headerRange.setValue(eventData[2]);
  
  var dateCell = hourSheet.getRange(7, currentColumn);
  dateCell.setValue(eventData[1]);
  dateCell.setNumberFormat("m/d");
  
  // Batch write all hour values at once
  if (indexArray.length > 0) {
    var rows = [];
    var values = [];
    
    for (let i = 0; i < indexArray.length; i++) {
      rows.push(indexArray[i] + peopleSheetStartingRow);
      let hourValue = hrReductionArray[i] ? -2 : eventData[0];
      values.push([hourValue]);
    }
    
    // Find the range that covers all rows
    var minRow = Math.min(...rows);
    var maxRow = Math.max(...rows);
    var rowCount = maxRow - minRow + 1;
    
    // Create a 2D array for all rows
    var allValues = new Array(rowCount);
    for (let i = 0; i < rowCount; i++) {
      allValues[i] = new Array(1).fill(null); // Fill with null
    }
    
    // Place values at correct positions
    for (let i = 0; i < rows.length; i++) {
      allValues[rows[i] - minRow][0] = values[i][0];
    }
    
    // Batch update the entire column
    hourSheet.getRange(minRow, currentColumn, rowCount, 1).setValues(allValues);
  }
}
*/
