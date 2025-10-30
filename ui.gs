function showDashboard() {
  try {
    const template = HtmlService.createTemplateFromFile('dashboard');
    return template.evaluate()
      .setTitle('Travel Orders Dashboard')
      .setWidth(1200)
      .setHeight(700);
  } catch (error) {
    Logger.log('Error in showDashboard: ' + error.message);
    const html = HtmlService.createHtmlOutput(
      '<h2>Error Loading Dashboard</h2><p>' + error.message + '</p>' +
      '<p>Please check that the HTML files exist and are properly formatted.</p>'
    );
    return html.setTitle('Error');
  }
}

function showNewToForm() {
  try {
    const template = HtmlService.createTemplateFromFile('newTo');
    return template.evaluate()
      .setTitle('New Travel Order')
      .setWidth(800)
      .setHeight(700);
  } catch (error) {
    Logger.log('Error in showNewToForm: ' + error.message);
    const html = HtmlService.createHtmlOutput(
      '<h2>Error Loading Form</h2><p>' + error.message + '</p>'
    );
    return html.setTitle('Error');
  }
}

function showUpdateForm(toId) {
  try {
    const template = HtmlService.createTemplateFromFile('update');
    template.toId = toId;
    return template.evaluate()
      .setTitle('Update Travel Order')
      .setWidth(800)
      .setHeight(700);
  } catch (error) {
    Logger.log('Error in showUpdateForm: ' + error.message);
    const html = HtmlService.createHtmlOutput(
      '<h2>Error Loading Update Form</h2><p>' + error.message + '</p>'
    );
    return html.setTitle('Error');
  }
}

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    Logger.log('Error in include(' + filename + '): ' + error.message);
    return '<p>Error loading ' + filename + ': ' + error.message + '</p>';
  }
}
