var agentNamesRange;
var waitLeaseColorsRange;
var agentList;
var waitLeaseColors;

function GenerateLeasingNumbers() {
  //get the script properties object and the relevant ranges
  scriptProperties = PropertiesService.getScriptProperties();
  agentNamesRange = scriptProperties.getProperty("agentNamesRange");
  waitLeaseColorsRange = scriptProperties.getProperty("waitLeaseColorsRange");

  //get the application tracker sheet
  GetSheet(scriptProperties, "trackerSheetName");

  //get the colors of active applications
  waitLeaseColors = GetWaitLeaseColors(waitLeaseColorsRange);

  //populate the list of agents from the list of agent names
  agentList = GetAgentNamesFromSheet(agentNamesRange).map(e => (new Agent(e)));

  //populate the agent object with data from the application sheet
  GetNameCounts(waitLeaseColors);
  GetCoSignSubAmt();

  //fill in the counts on the spreadsheet
  WriteCounts();
}

function GetAgentNamesFromSheet(range){
  const aNamesRange = currSheet.getRange(range).getValues();
  let aNames = [];

  for(var i = 0; i < aNamesRange.length; i++){
    const currValue = aNamesRange[i][0];

    if(currValue != ''){
      aNames[i] = currValue;
    }
  }

  return aNames;
}

function GetWaitLeaseColors(range){
  //get the 1d array of backround objects arrays from the range list
  const sheetColRangeBgObjs = [].concat(...currSheet.getRangeList(range.split(", ")).getRanges().map(e => e.getBackgroundObjects()));
  
  //get the 1d array of the actual backround objects 
  const sheetBgObjs = [].concat(...sheetColRangeBgObjs);

  //get the hex strings of the colors from the background objects
  const sheetColorVals = sheetBgObjs.map(e => e.asRgbColor().asHexString());

  return sheetColorVals;
}

function getPrimaryName(name){
  let primName = "";

  if(name.includes("co-signer for")){
    return primName;
  }

  const nameList = name.split(" ");
  primName = nameList[0] + " " + nameList[nameList.length-1];

  return primName;
}

function GetNameCounts(waitLeaseColors){
  const startRow = 18;
  const startColumn = 4;
  const appNameOffset = -3;

  const numRows = currSheet.getLastRow() - (startRow-1);
  const numColumns = 1;

  if(numRows <= 0){
    console.log("No data. Skipping counts");
    return;
  
  }

  const assignNames = [].concat(...currSheet.getRange(startRow, startColumn, numRows, numColumns).getValues());
  const appBgObjs = [].concat(...currSheet.getRange(startRow, startColumn + appNameOffset, numRows, numColumns).getBackgroundObjects());
  const appNames = [].concat(...currSheet.getRange(startRow, startColumn + appNameOffset, numRows, numColumns).getValues());
  const agentNames = agentList.map(e => e.name);

  for(let i = 0; i < assignNames.length; i++){
    const currAssignName = assignNames[i];
    const agentIndex = agentNames.indexOf(currAssignName);
    
    if(agentIndex >= 0){
      const currCellBgObj = appBgObjs[i];
      const currCellColor = currCellBgObj.asRgbColor().asHexString();

      if(waitLeaseColors.indexOf(currCellColor) >= 0){
        agentList[agentIndex].wait_for_lease.push(appNames[i].toLowerCase());
        agentList[agentIndex].wait_lease_count++;

        const primeName = getPrimaryName(appNames[i].toLowerCase());

        if(primeName.length != 0){
          agentList[agentIndex].wait_for_lease_prim.push(primeName.toLowerCase());
        }
      }
      
      agentList[agentIndex].name_count++;
    }
  }
}

function HasCoSigner(name, nameList){
  for(let i = 0; i < nameList.length; i++){
    currName = nameList[i];

    if(nameList[i].includes(name) && nameList[i].includes("co-signer for")){
      return true;
    }
  }

  return false;
}

function GetCoSignSubAmt(){
  for(let i = 0; i < agentList.length; i++){
    const currAgent = agentList[i];

    for(let j = 0; j < currAgent.wait_for_lease_prim.length; j++){
      currName = currAgent.wait_for_lease_prim[j]; 

      if(HasCoSigner(currName, currAgent.wait_for_lease)){
        agentList[i].wait_co_sign_sub++;
      }
    }
  }
}

function WriteCounts(){
  for(let i = 0; i < agentList.length; i++){
    const currAgent = agentList[i];

    values = [[currAgent.wait_lease_count - currAgent.wait_co_sign_sub,
              currAgent.wait_lease_count,
              currAgent.name_count]];
    
    const startRow = 4 + i;
    const startColumn = 6;
    const numRows = 1;
    const numColumns = 3;

    const outputRange = currSheet.getRange(startRow, startColumn, numRows, numColumns);

    outputRange.setValues(values);
  }
}