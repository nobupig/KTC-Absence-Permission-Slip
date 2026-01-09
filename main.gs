// function doGet() {
//   const t = HtmlService.createTemplateFromFile('index');
//   t.student = getLoginUser();
//   return t.evaluate()
//           .setTitle('公欠・忌引届')
//           .addMetaTag('viewport', 'width=device-width, initial-scale=1');
// }


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function base64ToBlob(base64Data, mimeType, name) {
  const decoded = Utilities.base64Decode(base64Data);
  return Utilities.newBlob(decoded, mimeType, name);
}

function submitEntries(entries, fileDataArray) {
  // ──────────────── ① ロック取得 ────────────────
  const lock = LockService.getScriptLock();
  try {
    // 最大 30 秒待ってロックを獲得
    lock.waitLock(30000);
  } catch (lockErr) {
    // ロック失敗時はユーザーに「混雑中」を返す
    throw new Error("現在混雑しています。数秒後に再度お試しください。");
  }

  try {
    Logger.log("=== 🔵 submitEntries START ===");
    Logger.log("📥 entries: " + JSON.stringify(entries));
    Logger.log("✅ typeof fileDataArray: " + typeof fileDataArray);
    Logger.log("✅ fileDataArray isArray: " + Array.isArray(fileDataArray));

    const email = Session.getActiveUser().getEmail();
    Logger.log("email: " + email);

    const student = getStudentInfo(email);
    Logger.log("student: " + JSON.stringify(student));
    if (!student) throw new Error("名簿に登録されていません");

    const sheet = SpreadsheetApp.openById('1PolwIbf2e3ebcleGMuUkl_86alI-WwxMR1TfkT7BgCQ').getSheetByName('申請ログ');
    const statusSheet = SpreadsheetApp.openById("1QfYNeYzAtbNwVm5rUl9cKixXKBbrCkCdu91p1ZUkJ-I").getSheetByName("申請状況");
    const now = new Date();
    const fiscalYear = getFiscalYear_(now);
    const month = ('0' + (now.getMonth() + 1)).slice(-2);
    const yearFolder = getOrCreateFolderByName(BASE_FOLDER_ID, `${fiscalYear}年度`);
    const monthFolder = getOrCreateFolder(yearFolder, `${month}月`);
    const studentFolder = getOrCreateFolder(monthFolder, `${student.id}_${student.name}`);

    const ADMIN_EMAIL = 'nyasui@ktc.ac.jp';
    [yearFolder, monthFolder, studentFolder].forEach(folder => {
      try {
        folder.addEditor(ADMIN_EMAIL);
      } catch (e) {
        Logger.log('addEditor skipped: ' + e.message);
      }
    });

    // ▼ 日付書式: yyyy/MM/dd（曜日）
    function formatJPDate(d) {
      if (!d) return "";
      const jsDate = (typeof d === "string") ? new Date(d) : d;
      const w = ["日", "月", "火", "水", "木", "金", "土"];
      return Utilities.formatDate(jsDate, "JST", "yyyy/MM/dd") + "（" + w[jsDate.getDay()] + "）";
    }

  for (let i = 0; i < entries.length; i++) {
  const entry = entries[i];
  const fileDataList = fileDataArray[i];
  const entryForLog = {
    reason: entry.reason,
    other: entry.otherReason || entry.other || "",
    from: entry.dateFrom || entry.from,
    periodFrom: entry.periodFrom,
    to: entry.dateTo || entry.to,
    periodTo: entry.periodTo
  };

  const folderName = formatDate(now, 'yyyy-MM-dd') + `_申請${i + 1}`;
  const entryFolder = getOrCreateFolder(studentFolder, folderName);

  const fileInfoList = saveFilesToFolder(entryFolder, fileDataList, i + 1, student.name);
  const weekCode = getWeekCode(now);

  // ▼ここが重要!! appendLogRowにentryForLogを渡してください
  appendLogRow(sheet, student, entryForLog, now, fileInfoList, entryFolder.getId(), weekCode);

  // --- 申請状況シートへの記録（こっちは今のentryでOK）---
  const absencePeriod =
  `${formatJPDate(entryForLog.from)} ${entryForLog.periodFrom} ～ ${formatJPDate(entryForLog.to)} ${entryForLog.periodTo}`;
  
  const logReason =
    entry.reason + (entry.otherReason ? `（${entry.otherReason}）` : '');

  statusSheet.appendRow([
    email,
    student.name,
    absencePeriod,
    logReason,
    "未処理", // 裁定
    ""        // 却下理由
  ]);
}

    Logger.log("submitEntries正常終了");
    return "申請を受け付けました";
  } catch (e) {
    Logger.log("submitEntriesエラー: " + e.message);
    throw e;
   } finally {
    lock.releaseLock();
   }
}


function getFiscalYear_(date) {
  const d = date || new Date();
  const y = d.getFullYear();
  const m = d.getMonth() + 1; // 1-12
  return (m <= 3) ? (y - 1) : y; // 1〜3月は前年の年度
}


// 申請状況シートから自分の申請状況だけ返す
function getMyStatusByEmail() {
  const email = Session.getActiveUser().getEmail().trim().toLowerCase();
  const statusSheet = SpreadsheetApp.openById("1QfYNeYzAtbNwVm5rUl9cKixXKBbrCkCdu91p1ZUkJ-I").getSheetByName("申請状況");
  const data = statusSheet.getDataRange().getValues();
  const headers = data[0];
  const idx = col => headers.indexOf(col);

  const myRows = data.slice(1).filter(row => {
    // 念のため両方trim + toLowerCaseで完全一致判定
    return String(row[idx("メールアドレス")]).trim().toLowerCase() === email;
  });

  return myRows.map(row => ({
    date: row[idx("欠席期間")],
    reason: row[idx("理由")],
    decision: row[idx("裁定")],
    rejectReason: row[idx("却下理由")],
    name: row[idx("氏名")]
  }));
}


// 1回のみ実行：申請ログにヘッダーを追加
function setAbsenceLogHeader() {
    const sheetId = '1PolwIbf2e3ebcleGMuUkl_86alI-WwxMR1TfkT7BgCQ';
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName('申請ログ'); // ← シート名はそのままでOK

  const headers = [
    'タイムスタンプ',
    'メールアドレス',
    '学籍番号',
    '氏名',
    '学年',
    '組・コース',
    '申請理由',
    'その他記述',
    '欠席開始日',
    '何限目から',
    '欠席終了日',
    '何限目まで',
    '添付ファイル一覧',
    '添付ファイルフォルダURL',
    '申請日（表示用）',
    '事務員処理',
    '事務員却下理由',
    '事務員処理日時',
    '教務主事裁定',
    '教務主事却下理由',
    '教務主事裁定日',
    '学生通知済み',
    '裁定確認URL'
  ];
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function testEmailDebug() {
  const email = Session.getActiveUser().getEmail().trim().toLowerCase();
  Logger.log('ログイン中のメール: ' + email);

  const ss = SpreadsheetApp.openById("1QfYNeYzAtbNwVm5rUl9cKixXKBbrCkCdu91p1ZUkJ-I");
  const sheet = ss.getSheetByName("申請状況");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = col => headers.indexOf(col);

  data.slice(1).forEach(row => {
    Logger.log(
      'A列: [' + String(row[idx("メールアドレス")]) + '] / 判定: ' +
      (String(row[idx("メールアドレス")]).trim().toLowerCase() === email ? '一致！' : 'NO')
    );
  });
}

