function generateGCNumbers() {
  //get the script properties object and the relevant ranges
  scriptProperties = PropertiesService.getScriptProperties();
  agentNamesRange = scriptProperties.getProperty("agentNamesRange");
  waitLeaseColorsRange = scriptProperties.getProperty("waitGCColorsRange");

  //get the guest card sheet
  GetSheet(scriptProperties, "guestCardsSheetName");

  //build the agent list from the list of agent names
  agentList = GetAgentNamesFromSheet(agentNamesRange).map(e => (new Agent(e)));

  //get the colors of the active guest cards
  waitLeaseColors = GetWaitLeaseColors(waitLeaseColorsRange);

  //populate the data in the agent objects
  GetNameCounts();

  //write this data
  writeGuestCardCounts();
}

function writeGuestCardCounts(){
  for(let i = 0; i < agentList.length; i++){
    const currAgent = agentList[i];

    values = [[currAgent.wait_lease_count,
              currAgent.name_count]];
    
    const startRow = 4 + i;
    const startColumn = 6;
    const numRows = 1;
    const numColumns = 2;

    const outputRange = currSheet.getRange(startRow, startColumn, numRows, numColumns);

    outputRange.setValues(values);
  }
}
