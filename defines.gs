//Globals for use in all scripts
var scriptProperties;
var currSheet;
var spreadSheet;

//Agent class for counting up performance numbers
class Agent{
  constructor(name){
    this.name = name;
    this.name_count = 0;
    this.wait_lease_count = 0;
    this.wait_co_sign_sub = 0;
    this.wait_for_lease = [];
    this.wait_for_lease_prim = [];
  }
}

//Different progress steps in the application tracker
var AppCellTypes = {
  EXECUTED: 0,
  REPORTED: 1,
  SIG_WAITING: 2,
  UNDER_18: 3,
  LEASE_READY: 4,
  LSE_RDY_WAITING: 5,
  ROOM_PREF: 6,
  LEASE_DETAILS: 7,
  APPROVED: 8,
  CANCELLED: 9,
  DENIED: 10,
  MANUAL: 11, 
  SCREEN_RDY: 12,
  SCR_RDY_WAITING: 13,
  WAITING_SCREEN: 14
};

var AppCellTypesIterator = Object.values(AppCellTypes);

//Different progress steps in the guest card tracker
var GCCellTypes = {
  APP_COMPLETE: AppCellTypes.EXECUTED,
  APP_SENT: AppCellTypes.SIG_WAITING,
  COLD: AppCellTypes.CANCELLED,
  APP_NOT_SENT: AppCellTypes.WAITING_SCREEN
};

var GCCellTypesIterator = Object.values(GCCellTypes);

//All colors of cells that can be used 
var AppCellTypesHexColor = [
  "#70ad47",
  "#a8d08d",
  "#ffc000",
  "#ffffff",
  "#4472c4",
  "#8eaadb",
  "#f2f2f2",
  "#ed7d31",
  "#c6efce",
  "#a5a5a5",
  "#ffc7ce",
  "#ffcc99",
  "#5b9bd5",
  "#9cc2e5",
  "#ffeb9c"
];

var FormatFields = {
  FONT_COLOR: 0,
  FONT_FAMILY: 1,
  FONT_STYLE: 2,
  FONT_WEIGHT: 3,
  BG_OBJ: 4, 

  NUM_FORMAT_FIELDS: 5
};

var AppsColumnNames = {
  APPLICANTS: 0,
  DATE_RECEIVED: 1,
  DATE_MOVE_IN: 2,
  ASSIGNED_AGENT: 3,
  NOTES: 4,

  NUM_COLUMN_NAMES: 5
}