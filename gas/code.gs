// ====================================
// ★ 以下を自分の情報に変更してください
// ====================================
var ADMIN_EMAIL = 'hirokazusugisugi@gmail.com,stevenm.kawasaki@gmail.com'; // 通知を受け取るメールアドレス
var STUDY_GROUP_NAME = 'UX/UI研究会';

var VENUE_INFO = [
  '【開催情報】',
  '日時: 毎月第2水曜日 19:30〜21:00',
  '※ 日時は変更となる場合がございます。変更の際は別途ご連絡いたします。',
  '',
  '【会場参加の場合】',
  '場所: ナレッジサロン（グランフロント大阪 北館7F）',
  ''
].join('\n');

var ZOOM_INFO = '【Zoom参加の場合】\nZoomのURLは開催日が近くなりましたら、別途メールにてお送りいたします。\n';

// ====================================
// 初回セットアップ用（1回だけ実行）
// Apps Scriptエディタで setupSheets を選択して▶実行
// ====================================
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 見学申込シート
  var visitSheet = ss.getSheetByName('見学申込');
  if (!visitSheet) {
    visitSheet = ss.insertSheet('見学申込');
  }
  visitSheet.getRange('A1:F1').setValues([['タイムスタンプ', 'お名前', 'ふりがな', 'メールアドレス', '協会登録', '参加方法']]);
  visitSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#E3F2FD');
  visitSheet.setColumnWidth(1, 180);
  visitSheet.setColumnWidth(2, 120);
  visitSheet.setColumnWidth(3, 120);
  visitSheet.setColumnWidth(4, 240);
  visitSheet.setColumnWidth(5, 100);
  visitSheet.setColumnWidth(6, 100);

  // お問い合わせシート
  var contactSheet = ss.getSheetByName('お問い合わせ');
  if (!contactSheet) {
    contactSheet = ss.insertSheet('お問い合わせ');
  }
  contactSheet.getRange('A1:E1').setValues([['タイムスタンプ', 'お名前', 'ふりがな', 'メールアドレス', 'お問い合わせ内容']]);
  contactSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#E3F2FD');
  contactSheet.setColumnWidth(1, 180);
  contactSheet.setColumnWidth(2, 120);
  contactSheet.setColumnWidth(3, 120);
  contactSheet.setColumnWidth(4, 240);
  contactSheet.setColumnWidth(5, 400);

  // デフォルトの Sheet1 を削除
  var defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  SpreadsheetApp.getUi().alert('セットアップ完了！\\n「見学申込」「お問い合わせ」シートを作成しました。');
}

// ====================================
// メイン処理（変更不要）
// ====================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var type = data.type; // 'visit' or 'contact'

    if (type === 'visit') {
      handleVisit(data);
    } else if (type === 'contact') {
      handleContact(data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 見学申込 ──
function handleVisit(data) {
  // スプレッドシートに記録
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('見学申込');
  sheet.appendRow([
    new Date(),
    data.name,
    data.furigana,
    data.email,
    data.membership,
    data.method
  ]);

  // 参加方法に応じた情報
  var methodLabel = data.method === 'venue' ? '会場参加' : 'Zoom参加';
  var participationInfo = '';
  if (data.method === 'venue') {
    participationInfo = VENUE_INFO + '\n' + ZOOM_INFO;
  } else {
    participationInfo = ZOOM_INFO + '\n' + VENUE_INFO;
  }

  // 申込者への自動返信メール
  var replySubject = '【' + STUDY_GROUP_NAME + '】見学のお申し込みありがとうございます';
  var replyBody = [
    data.name + ' 様',
    '',
    STUDY_GROUP_NAME + 'の見学にお申し込みいただき、ありがとうございます。',
    '',
    '以下の内容で受け付けました。',
    '─────────────────────────',
    'お名前: ' + data.name + '（' + data.furigana + '）',
    'メールアドレス: ' + data.email,
    '大阪府中小企業診断士協会: ' + data.membership,
    '参加方法: ' + methodLabel,
    '─────────────────────────',
    '',
    '本研究会はハイブリッド（会場＋Zoom）開催となっております。',
    '',
    participationInfo,
    '※ Zoom参加をご希望の方には、開催日が近くなりましたら',
    '  別途ZoomのURLをメールにてお送りいたします。',
    '',
    '当日のご参加をお待ちしております。',
    'ご不明な点がございましたら、このメールにご返信ください。',
    '',
    '─────────────────────────',
    '大阪府中小企業診断士協会',
    STUDY_GROUP_NAME,
    '─────────────────────────'
  ].join('\n');

  MailApp.sendEmail({
    to: data.email,
    subject: replySubject,
    body: replyBody,
    replyTo: ADMIN_EMAIL
  });

  // 管理者への通知メール
  var adminSubject = '【見学申込】' + data.name + ' 様（' + methodLabel + '）';
  var adminBody = [
    '見学の申込がありました。',
    '',
    'お名前: ' + data.name + '（' + data.furigana + '）',
    'メールアドレス: ' + data.email,
    '大阪府中小企業診断士協会: ' + data.membership,
    '参加方法: ' + methodLabel,
    '申込日時: ' + new Date().toLocaleString('ja-JP'),
    '',
    'スプレッドシートで確認:',
    SpreadsheetApp.getActiveSpreadsheet().getUrl()
  ].join('\n');

  MailApp.sendEmail(ADMIN_EMAIL, adminSubject, adminBody);
}

// ── お問い合わせ ──
function handleContact(data) {
  // スプレッドシートに記録
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('お問い合わせ');
  sheet.appendRow([
    new Date(),
    data.name,
    data.furigana,
    data.email,
    data.message
  ]);

  // 申込者への自動返信メール
  var replySubject = '【' + STUDY_GROUP_NAME + '】お問い合わせを受け付けました';
  var replyBody = [
    data.name + ' 様',
    '',
    'お問い合わせいただき、ありがとうございます。',
    '以下の内容で受け付けました。',
    '',
    '─────────────────────────',
    'お問い合わせ内容:',
    data.message,
    '─────────────────────────',
    '',
    '担当者より折り返しご連絡いたしますので、',
    '今しばらくお待ちください。',
    '',
    '─────────────────────────',
    '大阪府中小企業診断士協会',
    STUDY_GROUP_NAME,
    '─────────────────────────'
  ].join('\n');

  MailApp.sendEmail({
    to: data.email,
    subject: replySubject,
    body: replyBody,
    replyTo: ADMIN_EMAIL
  });

  // 管理者への通知メール
  var adminSubject = '【お問い合わせ】' + data.name + ' 様';
  var adminBody = [
    'お問い合わせがありました。',
    '',
    'お名前: ' + data.name + '（' + data.furigana + '）',
    'メールアドレス: ' + data.email,
    '日時: ' + new Date().toLocaleString('ja-JP'),
    '',
    '内容:',
    data.message,
    '',
    '返信先: ' + data.email
  ].join('\n');

  MailApp.sendEmail(ADMIN_EMAIL, adminSubject, adminBody);
}

// ── CORS対応（プリフライトリクエスト） ──
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ result: 'ok', message: 'UX/UI研究会フォームAPI' }))
    .setMimeType(ContentService.MimeType.JSON);
}
