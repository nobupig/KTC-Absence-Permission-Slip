 function doGet(e) {
  // テンプレートにサーバー側で取得した entries を渡す
  const tmpl = HtmlService.createTemplateFromFile('status');
  tmpl.entries = getMyStatusByEmail_();  // 内部関数から同期取得
  return tmpl.evaluate()
    .setTitle('申請状況の確認');
 }

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getMyStatusByEmail_() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) throw new Error("Googleアカウントにログインしていません");
    const lowerEmail = email.trim().toLowerCase();

    // 1. 名簿シート：メールアドレス列だけを抽出（列名が「メールアドレス」で1列目なら colNum=1 ）
    const rosterSheet = SpreadsheetApp.openById("1ZL3LtXDaApR-6IbNJ3n1JtpeOob-Lv_RPWuk-pL6ZaA").getSheetByName("名簿");
    const headers = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
    const emailColNum = headers.indexOf("メールアドレス") + 1; // GASは1始まり
    if (emailColNum < 1) throw new Error("名簿シートに「メールアドレス」列が見つかりません");

    const emailList = rosterSheet.getRange(2, emailColNum, rosterSheet.getLastRow() - 1, 1).getValues();
    const isStudent = emailList.some(row =>
      String(row[0]).trim().toLowerCase() === lowerEmail
    );
    if (!isStudent) throw new Error("あなたは学生名簿に登録されていません");

    // 2. 申請状況シート全体から該当履歴を抽出
    const statusSheet = SpreadsheetApp.openById("1QfYNeYzAtbNwVm5rUl9cKixXKBbrCkCdu91p1ZUkJ-I").getSheetByName("申請状況");

      const lastRow = statusSheet.getLastRow();
      const statusHeaders = statusSheet.getRange(1, 1, 1, statusSheet.getLastColumn()).getValues()[0];
      const lastCol = statusHeaders.length;
      if (lastRow < 2) return [];
      const statusData = statusSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();


    const statusEmailIdx = statusHeaders.indexOf("メールアドレス");
    if (statusEmailIdx < 0) throw new Error("申請状況シートに「メールアドレス」列が見つかりません");
    const idx = col => statusHeaders.indexOf(col);

    const myRows = statusData.filter(row =>
      String(row[statusEmailIdx]).trim().toLowerCase() === lowerEmail
    );

    return myRows.map(row => ({
      date: row[idx("欠席期間")],
      reason: row[idx("理由")],
      decision: row[idx("裁定")],
      rejectReason: row[idx("却下理由")],
      name: row[idx("氏名")]
    }));
  } catch(e) {
    throw new Error("GASエラー: " + e.message);
  }
}

