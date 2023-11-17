function GetSheet(scriptProperties, sheetName){
  try{
    spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    if(sheetName === ""){
      currSheet = spreadSheet.getActiveSheet();
    }else{
      currSheet = spreadSheet.getSheetByName(scriptProperties.getProperty(sheetName));
    }
  }
  catch (err){
    console.log(err.message);
  }
}

