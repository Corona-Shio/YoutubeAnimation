const SHEET_GANTT_CHART = "ガントチャート";
const SHEET_DONE_TASK   = "完了タスク";
const SHEET_MASTER      = "Master";
const SHEET_PEKETTSU    = "ペケッツ";

const SORT_TYPE_PERSON  = "PERSON";
const SORT_TYPE_CHANNEL = "CHANNEL";
const HEADING_STRING    = "見出し";

const STATUS_COL_NUMBER      = 1;
const CHANNEL_COL_NUMBER     = 2;
const PERSON_COL_NUMBER      = 3;
const NUMBER_COL_NUMBER      = 4;
const TASK_TYPE_COL_NUMBER   = 5;
const TASK_DETAIL_COL_NUMBER = 6;
const START_DAY_COL_NUMBER   = 7;
const END_DAY_COL_NUMBER     = 8;

const STATUS_ARRAY_NUMBER      = STATUS_COL_NUMBER - 1;
const CHANNEL_ARRAY_NUMBER     = CHANNEL_COL_NUMBER - 1;
const PERSON_ARRAY_NUMBER      = PERSON_COL_NUMBER - 1;
const NUMBER_ARRAY_NUMBER      = NUMBER_COL_NUMBER - 1;
const TASK_TYPE_ARRAY_NUMBER   = TASK_TYPE_COL_NUMBER - 1;
const TASK_DETAIL_ARRAY_NUMBER = TASK_DETAIL_COL_NUMBER - 1;
const START_DAY_ARRAY_NUMBER   = START_DAY_COL_NUMBER - 1;
const END_DAY_ARRAY_NUMBER     = END_DAY_COL_NUMBER - 1;

function mainSortChannel() {

  let targetRange = getTargetRange();

  // 完了したタスクを別シートにコピー
  copyDoneTask(targetRange);

  let todoTaskListAddedHeading = makeData(targetRange,SORT_TYPE_CHANNEL);

  // 並び替え ここだけやることが分岐する
  let sortedValues = sortValuesInChannel(todoTaskListAddedHeading);

  setSortedValues(targetRange,sortedValues)

}

function mainSortPerson() {

  let targetRange = getTargetRange();

  // 完了したタスクを別シートにコピー
  copyDoneTask(targetRange);

  let todoTaskListAddedHeading = makeData(targetRange,SORT_TYPE_PERSON);

  // 並び替え ここだけやることが分岐する
  let sortedValues = sortValuesInPerson(todoTaskListAddedHeading);

  setSortedValues(targetRange,sortedValues)

}


/**
 * 完了したタスク・見出しを除外
 * 見出し行の作成
 * タスクリストと見出し行を結合
 * そのリストを返す
 * */
function makeData(targetRange,sortType){

  let headingList = [];

  // 完了したタスク・見出しを除外
  let listExceptHeadingAndDoneTask = targetRange.getValues().filter(
    record => record[STATUS_ARRAY_NUMBER] == false
           && record[TASK_TYPE_ARRAY_NUMBER] !== HEADING_STRING
  );

  // 見出し行の作成
  if(sortType == SORT_TYPE_CHANNEL){
    // チャンネル優先ソート
    headingList = createChannelHeadingList(targetRange);
  } else {
    // 担当優先ソート
    headingList = createPersonHeadingList(targetRange);
  }

  // タスクリストと見出し行を結合
  let todoTaskListAddedHeading = listExceptHeadingAndDoneTask.concat(headingList)

  return todoTaskListAddedHeading

}

function setSortedValues(targetRange,sortedValues){

  // 貼り付けのための範囲を調整
  let filterdRange = targetRange.offset(0,0,sortedValues.length)
  
  // 範囲の値をクリア
  targetRange.clearContent()

  // 最終的なセット
  filterdRange.setValues(sortedValues);

}


function createChannelHeadingList(targetRange){

  let channelList = getChannelList();
  let distinctAllList = [];
  let headingList = []

  let values = targetRange.getValues();

  // 重複のないチャンネルと脚本番号のリストを作る
  for(i=0;i<channelList.length;i++){
    
    let filteredValuesByChannel = values.filter(record => record[CHANNEL_ARRAY_NUMBER] == channelList[i]);
    let scriptNumberList = [];

    for(j=0;j<filteredValuesByChannel.length;j++){
      scriptNumberList.push(filteredValuesByChannel[j][NUMBER_ARRAY_NUMBER]);
    }

    let distinctList = Array.from(new Set(scriptNumberList));

    for(j=0;j<distinctList.length;j++){
      distinctAllList.push([channelList[i],distinctList[j]]);
    }

  }

  // 見出しの配列を作る
  for(i=0;i<distinctAllList.length;i++){
    headingList.push(["",distinctAllList[i][0],"",distinctAllList[i][1],HEADING_STRING,"","","",""])
  }

  return headingList

}

function createPersonHeadingList(targetRange){

  let personList = [];
  let distinctAllList = [];
  let headingList = []

  let values = targetRange.getValues();

  for(i=0;i<values.length;i++){
    personList.push(values[i][PERSON_ARRAY_NUMBER])
  }

  distinctAllList = Array.from(new Set(personList));

  // 見出しの配列を作る
  for(i=0;i<distinctAllList.length;i++){
    headingList.push(["","",distinctAllList[i],"",HEADING_STRING,"","","",""])
  }

  return headingList

}

/* 完了済みタスクを別シートに移動させる */
function copyDoneTask(targetRange){

  let doneTaskList = filterDoneTask(targetRange);

  if(doneTaskList.length > 0){
    pasteDoneTask(doneTaskList);  
  }

}

function filterDoneTask(targetRange){

  // タスクが完了してるものだけを残す
  let doneTaskList = targetRange.getValues().filter(record => record[0] == true);

  return doneTaskList

}

function pasteDoneTask(doneTaskList){

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let doneTaskSheet = ss.getSheetByName(SHEET_DONE_TASK);  

  const START_COL = 1;
  const WRITE_START_ROW = doneTaskSheet.getLastRow() + 1;

  doneTaskSheet.getRange(WRITE_START_ROW,START_COL,doneTaskList.length,doneTaskList[0].length).setValues(doneTaskList);

}

function sortValuesInChannel(targetValues) {

  let channelList  = getChannelList();
  let taskTypeList = getTaskTypeList();

  targetValues.sort((a,b) => 
        compareArrayElements(b,a,STATUS_ARRAY_NUMBER)
     || compareArrayNumbers(a,b,channelList,CHANNEL_ARRAY_NUMBER)
     || compareArrayElements(a,b,NUMBER_ARRAY_NUMBER)
     || compareArrayNumbers(a,b,taskTypeList,TASK_TYPE_ARRAY_NUMBER)
  );

  return targetValues

}

function sortValuesInPerson(targetValues) {

  let personList   = getPersonList();
  let taskTypeList = getTaskTypeList();

  targetValues.sort((a,b) => 
        compareArrayElements(b,a,STATUS_ARRAY_NUMBER)
     || compareArrayNumbers(a,b,personList,PERSON_ARRAY_NUMBER)
     || compareArrayElements(a,b,END_DAY_ARRAY_NUMBER)
     || compareArrayNumbers(a,b,taskTypeList,TASK_TYPE_ARRAY_NUMBER)
  );

  return targetValues

}

function getTargetRange() {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(SHEET_GANTT_CHART);  

  const START_ROW = 9;
  const START_COL = 1;
  const END_ROW   = targetSheet.getLastRow();
  const END_COL   = 9;
  const COUNT_DATA = END_ROW - START_ROW + 1;

  let targetRange = targetSheet.getRange(START_ROW,START_COL,COUNT_DATA,END_COL);

  return targetRange

}

function getChannelList() {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(SHEET_MASTER);  

  const CHANNEL_COL  = 1;
  const START_ROW = 2;
  const END_ROW   = targetSheet.getRange(1,CHANNEL_COL).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const COUNT_DATA = END_ROW - START_ROW + 2; // 空白セルも取得するために「+2」

  let channelList = targetSheet.getRange(START_ROW,CHANNEL_COL,COUNT_DATA).getValues();

  return channelList.flat()

}

function getPersonList() {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(SHEET_MASTER);  

  const PERSON_COL  = 2;
  const START_ROW = 2;
  const END_ROW   = targetSheet.getRange(1,PERSON_COL).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const COUNT_DATA = END_ROW - START_ROW + 2; // 空白セルも取得するために「+2」

  let personList = targetSheet.getRange(START_ROW,PERSON_COL,COUNT_DATA).getValues();

  return personList.flat()

}

function getTaskTypeList() {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(SHEET_MASTER);  

  const TASK_TYPE_COL  = 3;
  const START_ROW = 2;
  const END_ROW   = targetSheet.getRange(1,TASK_TYPE_COL).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const COUNT_DATA = END_ROW - START_ROW + 2; // 空白セルも取得するために「+2」

  let taskTypeList = targetSheet.getRange(START_ROW,TASK_TYPE_COL,COUNT_DATA).getValues();

  return taskTypeList.flat()

}

// 引数をa,bで渡すと昇順、b,aで渡すと降順
function compareArrayElements(a,b,arrayNumber){
  return a[arrayNumber] - b[arrayNumber]
}

// 引数をa,bで渡すと昇順、b,aで渡すと降順
// function compareArrayElements(a,b,arrayNumber){

//   Logger.log("a:"+a[arrayNumber]+" b:"+b[arrayNumber])

//   a = a[arrayNumber]
//   b = b[arrayNumber]

//   const sa = String(a).replace(/(\d+)/g, m => m.padStart(30, '0'));
//   const sb = String(b).replace(/(\d+)/g, m => m.padStart(30, '0'));

//   Logger.log(sa < sb ? -1 : sa > sb ? 1 : 0)

//   return sa < sb ? -1 : sa > sb ? 1 : 0;

// }



// 引数をa,bで渡すと昇順、b,aで渡すと降順
function compareArrayNumbers(a,b,list,arrayNumber){
  return list.indexOf(a[arrayNumber]) - list.indexOf(b[arrayNumber])
}
