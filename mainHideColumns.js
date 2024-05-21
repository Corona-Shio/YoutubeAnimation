const SHEET_HIDE_COLUMNS = "hideColumns";

// 今日の日付を探して、指定列から今日の日付までの列を非表示にする
function mainHideColumns() {

  let infoList = getInfoList();

  for(let i=0; i<infoList.length; i++){
    let nameArray = getDateRange(infoList[i]);
    let targetNumber = findToday(nameArray);
    hideColumuns(infoList[i], targetNumber);
  }

}

function getInfoList() {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(SHEET_HIDE_COLUMNS);  

  const START_COL = 1;
  const START_ROW = 2;
  const END_ROW = targetSheet.getLastRow();
  const COUNT_DATA = END_ROW - START_ROW + 1;

  let infoList = targetSheet.getRange(START_ROW,START_COL,COUNT_DATA,2).getValues();

  return infoList

}

function getDateRange(infoList) {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(infoList[0]);

  const startColumn = infoList[1];
  const lastColumn  = targetSheet.getLastColumn();
  const countColumn = lastColumn - startColumn -1;
  const TARGET_ROW = 7;

  const nameArray = targetSheet.getRange(TARGET_ROW,startColumn,1,countColumn).getValues().flat();

  return nameArray

}

function findToday(nameArray) {

  let targetNumber;

  let date = new Date();
  let today = Utilities.formatDate(date, "JST", "yyyy/MM/dd")

  for(let i=0; i<nameArray.length; i++){
    
    const isExisted = Utilities.formatDate(nameArray[i], "JST", "yyyy/MM/dd").indexOf(today);

    if(isExisted != -1){
      targetNumber = i;
      break
    }

  }

  return targetNumber
  
}

function hideColumuns(infoList,targetNumber) {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(infoList[0]);

  const startColumn = infoList[1];

  targetSheet.hideColumns(startColumn,targetNumber);

}
