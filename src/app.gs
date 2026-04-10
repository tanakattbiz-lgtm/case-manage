function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle(APP_INFO.title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover');
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}
