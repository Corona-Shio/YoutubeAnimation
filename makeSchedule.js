const HIROTAMA = "ヒロたま";
const PEKETTSU = "ペケッツ";

function onOpen() {

  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu('日程生成');
  menu.addItem("生成したいところの行を選択してね（ペケッツとヒロたまのみ対応）","formatRequestSchedule");
  menu.addToUi();

}

function formatRequestSchedule() {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_GANTT_CHART);
  
  // アクティブなセルから行番号を取得する
  let activeCell = getActiveCell(ss);
  let rowNumber = activeCell.getRow();

  let record = getRecord(sheet, rowNumber);

  displayFinalMessage(record);

}

function getActiveCell(ss) {
  
  var sheet = ss.getActiveSheet();
  var range = sheet.getActiveCell();

  return range

}

function getRecord(sheet, rowNumber) {

  let rowData = sheet.getRange(rowNumber, 1, 1, 8).getValues()[0]; // A列からH列までを取得
  
  let channel   = rowData[1];
  let person    = rowData[2];
  let number    = rowData[3];
  let taskType  = rowData[4];
  let title     = rowData[5];
  let startDate = formatDate(rowData[6]);
  let endDate   = formatDate(rowData[7]);

  // 公開日を取得
  let releaseDate = formatDate(searchAndFetchData(number, channel));

  // URLを取得
  let targetRange = getRangeFromSheetPekettsu(number);
  let url = targetRange.getRichTextValue().getLinkUrl();

  return { 
    "Channel"    : channel, 
    "Person"     : person, 
    "Number"     : number, 
    "TaskType"   : taskType,
    "Title"      : title, 
    "StartDate"  : startDate, 
    "EndDate"    : endDate, 
    "ReleaseDate": releaseDate,
    "Url"        : url
  };

}

function formatDate(date) {
  Logger.log(date)
  const WEEK_LIST = new Array('日', '月', '火', '水', '木', '金', '土');
  const week = WEEK_LIST[date.getDay()];
  
  // フォーマットを設定
  let formattedDate = Utilities.formatDate(date, "JST", "M月dd日");

  return formattedDate + " (" + week + ")";

}

function searchAndFetchData(targetNumber, targetChannel) {

  const HIROTAMA_RELEASE_SCHEDULE_RANGE = "4:4";
  const PEKETTSU_RELEASE_SCHEDULE_RANGE = "5:5";
  
  let targetRange = "";

  switch(targetChannel){
    case HIROTAMA:
      targetRange = HIROTAMA_RELEASE_SCHEDULE_RANGE;
      break;
    case PEKETTSU:
      targetRange = PEKETTSU_RELEASE_SCHEDULE_RANGE;
      break;
    default:
      return null;
  }

  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_GANTT_CHART);

  // 対象の行からデータを取得
  let rowData = sheet.getRange(targetRange).getValues()[0];
  
  // 指定した番号を検索
  for (let col = 0; col < rowData.length; col++) {
    if (rowData[col] == targetNumber) {
      // 一致した列の7行目のデータを取得
      let result = sheet.getRange(7, col + 1).getValue();
      return result;
    }
  }

  // 一致するデータがない場合はnullを返す
  return null;

}

function getRangeFromSheetPekettsu(number) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PEKETTSU);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == number) {
      return sheet.getRange(i+1,2)
    }
  }

  return 'Number not found'; // numberが見つからなかった場合

}

function displayFinalMessage(record) {

    Browser.msgBox(
    "#" + record.Number+ "_" + record.Title + "\\n" +
    "編集開始：" + record.StartDate + "\\n" +
    "初稿予定：" + record.EndDate + "\\n" + 
    "公開予定：" + record.ReleaseDate + "\\n" + 
    "\\n" +
    record.Url
  );

}
