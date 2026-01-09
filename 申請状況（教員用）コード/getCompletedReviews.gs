function getCompletedReviews() {
  var email = Session.getActiveUser().getEmail();
  var role = getRole(email);

  const sheet = SpreadsheetApp.openById("1PolwIbf2e3ebcleGMuUkl_86alI-WwxMR1TfkT7BgCQ")
                             .getSheetByName("申請ログ");
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const result = [];

  // カラム位置（今まで通りでOK）
  const dateCol = header.indexOf("申請日（表示用）");
  const nameCol = header.indexOf("氏名");
  const gradeCol = header.indexOf("学年");
  const courseCol = header.indexOf("組・コース");
  const studentIdCol = header.indexOf("学籍番号");
  const reasonCol = header.indexOf("申請理由");
  const principalStatusCol = header.indexOf("教務主事裁定");
  const decisionDateCol = header.indexOf("教務主事裁定日");
  const rejectReasonCol = header.indexOf("教務主事却下理由");
  const periodYearCol = header.indexOf("欠席開始日");
  const periodFromCol = header.indexOf("何限目から");
  const periodToCol   = header.indexOf("何限目まで");

  data.slice(1).forEach(row => {
    if (row[principalStatusCol]) {
      // 一般教員は「許可」だけ表示
      if (role === "一般教員" && row[principalStatusCol] !== "許可") return;

    
      // 欠席期間periodを日本語表記で組み立て（欠席終了日・何限目まで対応版）
let period = "";
const startDateRaw = row[periodYearCol]; // 欠席開始日
const startPeriodRaw = row[periodFromCol]; // 何限目から
const endDateRaw = row[header.indexOf("欠席終了日")]; // 欠席終了日
const endPeriodRaw = row[periodToCol]; // 何限目まで
const youbiArr = ["日","月","火","水","木","金","土"];

// 時限文字列正規化（例：限限/限目/SHR対応）
const cleanPeriod = val => {
  if (!val) return "";
  if (typeof val === "string" && val.match(/^[A-Za-z]+$/)) return val;
  return String(val).replace(/限目|限限|限/g, "") + "限";
};
const startPeriod = cleanPeriod(startPeriodRaw);
const endPeriod = cleanPeriod(endPeriodRaw);

if (startDateRaw && startPeriod && endDateRaw && endPeriod) {
  const startDate = new Date(startDateRaw);
  const endDate = new Date(endDateRaw);
  period =
    Utilities.formatDate(startDate, "Asia/Tokyo", "yyyy/MM/dd") +
    "（" + youbiArr[startDate.getDay()] + "）" +
    startPeriod + " ～ " +
    Utilities.formatDate(endDate, "Asia/Tokyo", "yyyy/MM/dd") +
    "（" + youbiArr[endDate.getDay()] + "）" +
    endPeriod;
} else if (startDateRaw && startPeriod) {
  const startDate = new Date(startDateRaw);
  period =
    Utilities.formatDate(startDate, "Asia/Tokyo", "yyyy/MM/dd") +
    "（" + youbiArr[startDate.getDay()] + "）" +
    startPeriod;
}

      // 返却内容を役割で分岐
      let entry = {
        date: row[dateCol],
        name: row[nameCol],
        grade: row[gradeCol],
        course: row[courseCol],
        studentId: row[studentIdCol],
        reason: row[reasonCol],
        period: period
      };
      if (role === "事務員" || role === "教務主事") {
        entry.result = row[principalStatusCol];
        entry.decisionDate = (row[decisionDateCol] instanceof Date)
          ? Utilities.formatDate(row[decisionDateCol], "Asia/Tokyo", "yyyy/MM/dd")
          : (row[decisionDateCol] || "");
        entry.rejectReason = row[rejectReasonCol] || "";
      }
      result.push(entry);
    }
  });

  // 必ず role も一緒に返す
  return { entries: result, role: role };
}
