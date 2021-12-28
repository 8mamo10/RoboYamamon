function doGet(e) {
  var output = ContentService.createTextOutput("hoge");
  output.setMimeType(ContentService.MimeType.TEXT);
  return output;
}
