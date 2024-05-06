// mainメソッド実行用トリガーを設定する（毎週月曜AM7:00に起動、祝日に実行しないよう回避）
function setTrigger() {
  const date = new Date();
  let dayOfWeek = 0;
  dayOfWeek = date.getDay();

  // 実行日が祝日にならないようループ
  while(isHoliday(date)) {
    date.setDate(date.getDate() + 1);
  }

  // 設定した日付の12:00分に実行するトリガーを設定する
  const time = date;
  time.setHours(12);
  time.setMinutes(00);
  ScriptApp.newTrigger('notifyCleaningPosition').timeBased().at(time).create();
}

// Slackへのメッセージ投稿メソッドおよび担当個所のローテーションメソッドを呼び出す
function notifyCleaningPosition() {
  const sh = sheet();
  // メンバーが変わったらここの範囲を変更する。詳細は実シート参照
  const ranges = sh.getRange('A2:B7').getValues();
  // Slackに投稿するメッセージ本文の作成
  const message = getMessage(ranges);
  // Slackへ投稿する
  postSlackbot(message);
  // 次回実行用に担当個所を変更する
  updatePosition(sh, ranges);

  // 実行後にトリガーを削除
  delTrigger();
}

// 既存のトリガーを削除
function delTrigger() {
  // 設定されているトリガー一覧を取得
  const triggers = ScriptApp.getProjectTriggers();
  // notifyCleaningPosition実行用トリガーのみ削除
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "notifyCleaningPosition"){
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

// 掃除機（メイン）担当を取得
function getVacuum(ranges) {
  const vacuum = ranges.find(range => range[1] === 'vacuum');
  return vacuum[0];
}

// 掃除機（サブ）担当を取得
function getVacuumSub(ranges) {
  const vacuumSub = ranges.find(range => range[1] === 'vacuumSub');
  return vacuumSub[0];
}

// コーヒー（メイン）担当を取得
function getCoffee(ranges) {
  const coffee = ranges.find(range => range[1] === 'coffee');
  return coffee[0];
}

// コーヒー（サブ）担当を取得
function getCoffeeSub(ranges) {
  const coffeeSub = ranges.find(range => range[1] === 'coffeeSub');
  return coffeeSub[0];
}

// クイックルワイパー担当を取得
function getMop(ranges) {
  const mop = ranges.find(range => range[1] === 'mop');
  return mop[0];
}

// スーパーサブ担当を取得
function getSuperSub(ranges) {
  const superSub = ranges.find(range => range[1] === 'superSub');
  return superSub[0];
}

// 各担当の担当作業をローテーションする
function updatePosition(sh, ranges) {
  const position = ranges.map(range => [range[1]]);
  // 掃除機→掃除機（サブ）→コーヒー→コーヒー（サブ）→モップ→スーパーサブの順番で巡回
  const lastPosition = position.pop();
  position.unshift(lastPosition);
  // メンバーが変わったらここの範囲を変更する
  sh.getRange('B2:B7').setValues(position); 
}

// シート情報を取得
function sheet() {
  const spreadsheetID = PropertiesService.getScriptProperties().getProperty("SpreadSheetID")
  const ss = SpreadsheetApp.openById(spreadsheetID);
  console.log(spreadsheetID)
  const sh = ss.getSheets()[0];
  return sh;
}

// 各担当を取得し、Slackにとばすメッセージを作成する
function getMessage(ranges) {
  // 最新の各担当を取得
  const vacuum = getVacuum(ranges);
  const vacuumSub = getVacuumSub(ranges);
  const coffee = getCoffee(ranges);
  const coffeeSub = getCoffeeSub(ranges);
  const mop = getMop(ranges);
  const superSub = getSuperSub(ranges);
  // メッセージ本文を作成してreturn
  return `今週の掃除当番は以下の通りです。\n掃除機（メイン）: ${vacuum},\n掃除機（サブ）: ${vacuumSub},\nコーヒー（メイン）: ${coffee},\nコーヒー（サブ）: ${coffeeSub},\nクイックルワイパー: ${mop},\n スーパーサブ: ${superSub}`;
}

// 作成したメッセージをSlackAPIに渡す
function postSlackbot(message) {
  //トークン設定
  const token = PropertiesService.getScriptProperties().getProperty("BotToken");
  const slackApp = SlackApp.create(token);
  //投稿先のチャンネルを設定
  const channelId = '#わたなべてすと';
  //Slackへ投稿
  slackApp.postMessage(channelId, message);
}

// 祝日判定メソッド（実行日が祝日の場合、trueを返す）
function isHoliday(targetDate) {
  // 日本の祝日カレンダーのIDを取得
  const holidayCalendarId = 'ja.japanese#holiday@group.v.calendar.google.com';
  // カレンダーIDを使用してカレンダーを取得
  const calendar = CalendarApp.getCalendarById(holidayCalendarId);
  // ターゲットの日付のイベント（祝日）を取得
  const events = calendar.getEventsForDay(targetDate);
  // 上記で何らかのイベントを取得している場合、祝日と判定
  return events.length > 0;
}
