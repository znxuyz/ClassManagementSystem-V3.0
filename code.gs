function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("班級管理系統")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
