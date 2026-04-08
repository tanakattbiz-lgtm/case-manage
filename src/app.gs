function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('案件管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
