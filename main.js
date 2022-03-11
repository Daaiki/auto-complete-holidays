// 実行ボタンをスプレッドシート側に作成
function onOpen() {
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  //スプレッドシートのメニューにカスタムメニュー「カレンダー連携 > 実行」を作成
  const subMenus = [];
  subMenus.push({
    name: '実行',
    functionName: 'createSchedule'
  });
  sheet.addMenu('カレンダー連携', subMenus);
}

function createSchedule() {

  // 連携するアカウント
  const calendarId = CALENDAR_ID

  // シートを取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const initRow = 2
  const initCol = 1
  const rowRange = 2
  const colRange = sheet.getLastColumn()

  const schedules = sheet.getRange(initRow, initCol, rowRange, colRange).getValues()

  // カレンダーの取得
  const calendar = CalendarApp.getCalendarById(calendarId)

  // 日付の取得
  const days = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()

  for (let i = 0; i < sheet.getLastColumn(); i++) {
    // 稼働しない日を除く
    if (schedules[0][i] === '') continue

    // いつ稼働するかを取得
    const day = new Date(days[0][i])
    const startTime = schedules[0][i]
    const endTime = schedules[1][i]

    // 開始日時をフォーマット
    const startDate = new Date(day)
    startDate.setHours(startTime.getHours())
    startDate.setMinutes(startTime.getMinutes());

    // 終了日時をフォーマット
    const endDate = new Date(day);
    endDate.setHours(endTime.getHours())
    endDate.setMinutes(endTime.getMinutes());

    // 予定を作成
    calendar.createEvent(
      '対応可能',
      startDate,
      endDate
    );
  }
}

