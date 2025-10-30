function showDashboard() {
  const template = HtmlService.createTemplateFromFile('dashboard');
  return template.evaluate()
    .setTitle('Travel Orders Dashboard')
    .setWidth(1200)
    .setHeight(700);
}

function showNewToForm() {
  const template = HtmlService.createTemplateFromFile('newTo');
  return template.evaluate()
    .setTitle('New Travel Order')
    .setWidth(800)
    .setHeight(700);
}

function showUpdateForm(toId) {
  const template = HtmlService.createTemplateFromFile('update');
  template.toId = toId;
  return template.evaluate()
    .setTitle('Update Travel Order')
    .setWidth(800)
    .setHeight(700);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
