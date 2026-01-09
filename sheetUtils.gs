// 改修版 sheetUtils.gs
function appendLogRow(sheet, student, entry, timestamp, fileInfoList, folderId, weekCode) {
  sheet.appendRow([
    timestamp,
    student.email,
    student.id,
    student.name,
    student.grade,
    student.course,
    entry.reason,
    entry.other || '',
    entry.from,
    entry.periodFrom,
    entry.to,
    entry.periodTo,
    fileInfoList.length > 0 ? "有" : "無",         // 添付ファイル有無
    fileInfoList.map(f => f.name).join(','),      // 添付ファイル名リスト
    fileInfoList.map(f => f.url).join(','),       // 添付ファイルURLリスト
    fileInfoList.map(f => f.id).join(','),        // 添付ファイルIDリスト
    `https://drive.google.com/drive/folders/${folderId}`,
    weekCode,
    '', '', '', '', '', '', '', ''
  ]);
}


function formatDate(date, pattern) {
  return Utilities.formatDate(date, 'Asia/Tokyo', pattern);
}

function getWeekCode(date) {
  const wd = ['日','月','火','水','木','金','土'][date.getDay()];
  return formatDate(date, 'yyyy-MM-dd') + `(${wd})`;
}