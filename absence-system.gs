// 改修版 absence-system.gs
const SPREADSHEET_ID = '1ZL3LtXDaApR-6IbNJ3n1JtpeOob-Lv_RPWuk-pL6ZaA';
const ROSTER_SHEET_NAME = '名簿';

function doGet(e) {
  const email = Session.getActiveUser().getEmail();

  // 学生アカウント判定
  if (!email.endsWith('@ktc.ac.jp') || !isUserInRoster(email)) {
    return HtmlService.createHtmlOutput('<h2>アクセスが許可されていません。</h2>');
  }

  const studentInfo = getStudentInfo(email);
  if (!studentInfo) {
    return HtmlService.createHtmlOutput('<h2>名簿に登録されていません。</h2>');
  }

  // URLパラメータによるテンプレート分岐
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';

  const template = HtmlService.createTemplateFromFile(page);

  // index.htmlのみ student変数を渡す
  if (page === 'index') {
    template.student = studentInfo;
  }

  return template.evaluate()
    .setTitle(page === 'status' ? '申請状況' : '公欠届フォーム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}



function isUserInRoster(email) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ROSTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues().slice(1);
  Logger.log("isUserInRoster: email=" + email);
  data.forEach(row => Logger.log("名簿row=" + row[0]));
  if (email === 'nyasui@ktc.ac.jp') return true;
  return data.some(row => row[0] === email);
}

function getStudentInfo(email) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ROSTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const row = data.find(r => r[0] === email);
  if (!row) return null;
  return {
    email: row[0],
    id: row[1],
    grade: row[2],
    course: row[3],
    number: row[4],
    name: row[5],
  };
}

function getLoginUser() {
  const email = Session.getActiveUser().getEmail();
  const student = getStudentInfo(email); // 既存の関数と連携して取得
  if (!student) throw new Error("ログインユーザーが名簿に存在しません");
  return student;
}
