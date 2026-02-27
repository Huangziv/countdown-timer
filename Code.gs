// *****************************************************************
// ** Code.gs 程式碼 (已更新欄位索引以匹配 A欄 開始日期) **
// *****************************************************************

/**
 * 網頁應用程式入口函式 (不變)
 */
function doGet() {
  const htmlOutput = HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('演講排程計時器');
  
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  return htmlOutput;
}

// *****************************************************************
// ** 試算表 ID 和設定 **
// *****************************************************************
// 您的 Google Sheet ID
const SPREADSHEET_ID = '1TGkgxqP1litJyu0U4VA0v2paDprZtCOQK5RtQ0QofWY'; 
const SHEET_NAME = 'Sheet1'; // 假設是 Sheet1 (請檢查是否正確)
const START_ROW = 3; // 資料從第 2 行開始
const DATA_RANGE = 'A:G'; // 讀取 A 欄到 G 欄

// 定義欄位索引 (相對於讀取的範圍 A:G，A=0)
const COL_START_DATE = 0; // A 欄: 開始日期
const COL_START_TIME = 1; // B 欄: 開始時間
const COL_END_DATE = 2;   // C 欄: 結束日期
const COL_END_TIME = 3;   // D 欄: 結束時間
const COL_DURATION = 4;   // E 欄: 課程長度 (分)
const COL_NAME = 5;       // F 欄: 課程名稱
const COL_SPEAKER = 6;    // G 欄: 分享人員

/**
 * 將試算表中讀取到的 Date 物件（日期或時間）合併成一個完整的 DateTime 物件。
 * (此函式與前一版本相同)
 */
function combineDateTime(datePart, timePart) {
  if (!(datePart instanceof Date) || !(timePart instanceof Date)) {
    return null;
  }
  
  const year = datePart.getFullYear();
  const month = datePart.getMonth();
  const date = datePart.getDate();
  
  const hours = timePart.getHours();
  const minutes = timePart.getMinutes();
  const seconds = timePart.getSeconds();
  
  return new Date(year, month, date, hours, minutes, seconds);
}


/**
 * 根據目前的伺服器時間，從試算表中計算出當前課程和下一堂課程。
 * (核心邏輯與前一版本相同，但使用新的索引)
 */
function getScheduleInfo() {
  const now = new Date(); 
  const scriptTimeZone = Session.getScriptTimeZone();
  const nowTimeString = Utilities.formatDate(now, scriptTimeZone, "HH:mm:ss");
  
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`找不到工作表: ${SHEET_NAME}`);
    }

    const dataRange = sheet.getRange(`${DATA_RANGE.split(':')[0]}${START_ROW}:${DATA_RANGE.split(':')[1]}${sheet.getLastRow()}`);
    const data = dataRange.getValues();
    
    let currentCourse = null;
    let nextCourse = null;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      const startDateValue = row[COL_START_DATE];
      const startTimeValue = row[COL_START_TIME];
      const endDateValue = row[COL_END_DATE];
      const endTimeValue = row[COL_END_TIME];
      const durationMin = row[COL_DURATION]; 
      const name = row[COL_NAME];
      const speaker = row[COL_SPEAKER];

      if (!name || !startDateValue || !startTimeValue) continue;

      const startTime = combineDateTime(startDateValue, startTimeValue);
      const endTime = combineDateTime(endDateValue, endTimeValue);
      
      if (!startTime || !endTime || startTime.getTime() >= endTime.getTime()) {
          continue; 
      }
      
      const duration = `${Math.round(durationMin).toString()} 分鐘`; // 取整數分鐘
      
      const courseInfo = {
        name: name,
        speaker: speaker,
        // 格式化為 HH:mm 顯示
        start: Utilities.formatDate(startTime, scriptTimeZone, "HH:mm"),
        end: Utilities.formatDate(endTime, scriptTimeZone, "HH:mm"),
        duration: duration,
        remainingSeconds: 0 
      };

      // 判斷當前課程
      if (now >= startTime && now < endTime) {
        currentCourse = courseInfo;
        currentCourse.remainingSeconds = Math.floor((endTime.getTime() - now.getTime()) / 1000);
      } 
      
      // 判斷下一堂課程
      if (now < startTime) {
        if (!nextCourse) { 
            nextCourse = courseInfo;
        }
        if (currentCourse && nextCourse) break;
      }
    }

    // 返回結果給前端
    return {
      currentTime: nowTimeString,
      current: currentCourse,
      next: nextCourse,
      status: currentCourse ? 'ONGOING' : (nextCourse ? 'UPCOMING' : 'FINISHED')
    };
    
  } catch (e) {
    Logger.log("執行錯誤: " + e.toString());
    return {
      currentTime: nowTimeString,
      current: null,
      next: null,
      error: e.toString()
    };
  }
}
