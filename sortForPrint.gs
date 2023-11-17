function sortForPrinting() {
    scriptProperties = PropertiesService.getScriptProperties();
    GetSheet(scriptProperties, "");
  
    const startRow = 18;
    const startColumn = 1;
    const numRows = currSheet.getLastRow() - (startRow-1);
    const numColumns = 5;
  
    const sortRange = currSheet.getRange(startRow, startColumn, numRows, numColumns);
    const sortRangeBgObjs = sortRange.getBackgroundObjects();
  
    let sortColorItr;
    if(currSheet.getName() == scriptProperties.getProperty("trackerSheetName")){
      sortColorItr = AppCellTypesIterator;
    }else{
      sortColorItr = GCCellTypesIterator;
    }
  
    const sortColorOrder = sortColorItr.map(e => AppCellTypesHexColor[e]);
  
    // 2. Create the request body for using the batchUpdate method of Sheets API.
    const backgroundColorObj = sortRangeBgObjs.reduce((o, [a]) => {
        const rgb = a.asRgbColor();
        return Object.assign(o, {[rgb.asHexString()]: {red: rgb.getRed() / 255, green: rgb.getGreen() / 255, blue: rgb.getBlue() / 255}})
      }, {});
    const backgroundColors = sortColorOrder.map(e => backgroundColorObj[e]);
  
    sstartRow = startRow -1;
    sstartColumn = startColumn -1;
    const srange = {
      sheetId: currSheet.getSheetId(),
      startRowIndex: sstartRow,
      endRowIndex: sstartRow + numRows,
      startColumnIndex: sstartColumn,
      endColumnIndex: sstartColumn + numColumns
    };
  
    const temp = backgroundColors.map(rgb => ({backgroundColor: rgb}));
    const requests = [
      {sortRange: {range: srange, sortSpecs: [{dimensionIndex: AppsColumnNames.DATE_RECEIVED, sortOrder: "ASCENDING"}]}},
      {sortRange: {range: srange, sortSpecs: backgroundColors.map(rgb => ({backgroundColor: rgb}))}}
    ];
    
    // 3. Request to Sheets API using the request body.
    Sheets.Spreadsheets.batchUpdate({requests: requests}, spreadSheet.getId());
  }
  