function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('小数のしくみ 習熟ツール')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
