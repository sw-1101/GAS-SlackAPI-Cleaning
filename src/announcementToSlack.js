/**
 * Copyright 2024 Sho Watanabe
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

// トリガーを設定する（毎週月曜AM7:00に起動）
function setTrigger() {
  let date = new Date(2024, 4, 4);

  let dayOfWeek = 0;
  dayOfWeek = date.getDay();

  if(dayOfWeek == 1){
    dayOfWeek ++;
  }

  while(dayOfWeek != 1) {
    date.setDate(date.getDate() + 1);
    dayOfWeek = date.getDay();
  }

  if(isHoliday(date)) {
    console.log('しゅくじつだよ')
    date.setDate(date.getDate() + 1);
    console.log('実行日：', date)
  } else {
    console.log('しゅくじつじゃないよ')
    console.log('実行日：', date)
  }

  const time = date;
  time.setHours(12);
  time.setMinutes(00);
  ScriptApp.newTrigger('main').timeBased().at(time).create();
}

// Slackへのメッセージ投稿メソッドおよび担当個所のローテーションメソッドを呼び出すメインメソッド
function main() {
  const sh = sheet();
  // メンバーが変わったらここの範囲を変更する。詳細は実シート参照
  const ranges = sh.getRange('A2:B7').getValues();
  const message = getMessage(ranges);

  // postSlackbot(message);
  // updatePosition(sh, ranges);

  // 実行後にトリガーを削除
  delTrigger();
}

// 既存のトリガーを削除
function delTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for(const trigger of triggers){
    if(trigger.getHandlerFunction() == "main"){
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
  sh.getRange('B2:B7').setValues(position); // メンバーが変わったらここの範囲を変更する
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
  //SlackAPIで登録したボットのトークンを設定する
  const token = PropertiesService.getScriptPropertied().getProperty("BotToken");
  //ライブラリから導入したSlackAppを定義し、トークンを設定する
  const slackApp = SlackApp.create(token);
  //Slackボットがメッセージを投稿するチャンネルを定義する
  const channelId = '#わたなべてすと';
  //SlackAppオブジェクトのpostMessageメソッドでボット投稿を行う
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
  // イベントが存在するかどうかをチェック（存在すれば祝日、存在しなければ非祝日）
  return events.length > 0;
}
