function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
