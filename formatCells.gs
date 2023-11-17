var formatRangeTxt;

function formatFullyExecuted(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.EXECUTED);
}

function formatReported(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.REPORTED);
}

function formatWaitingSig(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.SIG_WAITING);  
}

function formatUnder18(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.UNDER_18);  
}

function formatReadyLease(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.LEASE_READY);
}

function formatRdyLseWait(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.LSE_RDY_WAITING);
}

function formatRoomPref(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.ROOM_PREF);
}

function formatLeaseDetails(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.LEASE_DETAILS);
}

function formatApproved(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.APPROVED);
}

function formatCancelled(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.CANCELLED);
}

function formatDenied(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.DENIED);
}

function formatManual(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.MANUAL);
}

function formatScreenReady(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.SCREEN_RDY);
}

function formatScrRdyWait(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.SCR_RDY_WAITING);
}

function formatWaitScreening(){
  formatCells(SpreadsheetApp.getActiveRange(), AppCellTypes.WAITING_SCREEN);
}

function formatCells(cellRange, formatType) {
  if(cellRange == null){
    return;
  }
  //Get the properties object and the sheet with the formatting template
  scriptProperties = PropertiesService.getScriptProperties();
  GetSheet(scriptProperties, "trackerSheetName");

  formatRangeTxt = scriptProperties.getProperty("appsFormatRange");

  var formatRange = currSheet.getRange(formatRangeTxt);

  //get the different format objects from our template range
  var formatFontColorObjs = formatRange.getFontColorObjects();
  var formatFontFamilies = formatRange.getFontFamilies();
  var formatFontStyles = formatRange.getFontStyles();
  var formatFontWeight = formatRange.getFontWeights();
  var formatBgObjs = formatRange.getBackgroundObjects();

  //select the correct format data
  var selectedFormat = [formatFontColorObjs[formatType][0],
                        formatFontFamilies[formatType],
                        formatFontStyles[formatType],
                        formatFontWeight[formatType],
                        formatBgObjs[formatType][0]];

  //apply the format to the cell range
  cellRange.setFontColorObject(selectedFormat[FormatFields.FONT_COLOR]);
  cellRange.setFontFamily(selectedFormat[FormatFields.FONT_FAMILY]);
  cellRange.setFontStyle(selectedFormat[FormatFields.FONT_STYLE]);
  cellRange.setFontWeight(selectedFormat[FormatFields.FONT_WEIGHT]);
  cellRange.setBackgroundObject(selectedFormat[FormatFields.BG_OBJ]);
}
