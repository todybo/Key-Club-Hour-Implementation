/*

//cody's hehehaha testing/expirementing with ui elements -- pretty much unused page of functions (and i think it doesnt work cuz of the 100% error rate in triggers)
function onOpen() {
  return
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Clean up tables")
    .addItem('Reset colors', 'cleanColors')
    .addItem('Reset data', 'removeData')
    .addItem('Reset to base (clean data + colors)', 'removeBoth')
    .addToUi()
}

function cleanColors() { 
  var rows = dataSheet.getRange(dataSheetStartingRow,1,50,15).getBandings();
  rows.forEach(banding => banding.remove())
  dataSheet.getRange(dataSheetStartingRow,1,50,12).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  dataSheet.getRange(dataSheetStartingRow,13,10,3).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
}

//const hourSheetColumnDataPosition = [1, 4, 7, 10, 13];
function removeData() {
  for (i = 0; i < hourSheetColumnDataPosition.length; i++) {
    dataSheet.getRange(dataSheetStartingRow, hourSheetColumnDataPosition[i], 50, 2).setValue("");
    dataSheet.getRange(dataSheetStartingRow, hourSheetColumnDataPosition[i]+2, 50, 1).setValue(false);
  }
}

function removeBoth() {
  removeData();
  cleanColors();
}

//resets colors but i'm pretty sure it doesnt work.... (has a 100% error rate trigger)
function onEdit2(e) { 
  return
  if (css.getActiveSheet().getName() !== otherDataSheet.getName()) {return}
  if (e.range.getA1Notation() == dataSheet.getRange(otherDataSheetStartingRow,1).getA1Notation()) {
    var rows = otherDataSheet.getRange(otherDataSheetStartingRow,1,100,4);
    rows.getBandings().forEach(banding => banding.remove())
    rows.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  }
} */