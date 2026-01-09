function doGet(e) {
  return HtmlService
    .createTemplateFromFile('completed')
    .evaluate()
    .setTitle('裁定済申請一覧');
}