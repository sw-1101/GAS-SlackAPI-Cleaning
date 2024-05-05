
// Slack上でのメッセージ応答用メソッド
function doPost(e) {
  // Events APIからのPOSTを取得
  const json = JSON.parse(e.postData.getDataAsString());

   // Events APIを使用する初回、URL Verificationのための記述
  if (json.type == "url_verification") {
    return ContentService.createTextOutput(json.challenge);
  }

  const event = json.event;

  // スプレッドシートにアクセス
  const spreadsheetID = PropertiesService.getScriptProperties().getProperty("SpreadSheetID")
  const spreadsheet = SpreadsheetApp.openById(spreadsheetID);
  const sheet = spreadsheet.getSheetByName('シート3');

  // 現在の日付と時刻
  const timestamp = new Date();

  // スプレッドシートに書き込むデータ（例: 日付、イベントの種類、イベントの内容）
  const data = [timestamp, event.type, JSON.stringify(event)];

  // スプレッドシートの最後にデータを追加
  sheet.appendRow(data);

  // 無限ループ防止のため条件に合ったメッセージを受信時のみ返信する
  if (event && event.type === "message" && event.bot_id === undefined) {
    // 現在の担当区分をメッセージにして送信する
    main()
  }
}

// Slackへのメッセージ投稿メソッドおよび担当個所のローテーションメソッドを呼び出す
function main() {
  const sh = sheet();
  // メンバーが変わったらここの範囲を変更する。詳細は実シート参照
  const ranges = sh.getRange('A2:B7').getValues();
  const message = getMessage(ranges);

  // メッセージ投稿
  postSlackbot(message);
}

// 作成したメッセージをSlackAPIに渡す
function postSlackbot(message) {
  //SlackAPIで登録したボットのトークンを設定する
  const token = PropertiesService.getScriptProperties().getProperty("BotToken");
  //ライブラリから導入したSlackAppを定義し、トークンを設定する
  const slackApp = SlackApp.create(token);
  //Slackボットがメッセージを投稿するチャンネルを定義する
  const channelId = '#わたなべてすと';
  //SlackAppオブジェクトのpostMessageメソッドでボット投稿を行う
  slackApp.postMessage(channelId, message);
}

// シート情報を取得
function sheet() {
  const spreadsheetID = PropertiesService.getScriptProperties().getProperty("SpreadSheetID")
  const ss = SpreadsheetApp.openById(spreadsheetID);
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