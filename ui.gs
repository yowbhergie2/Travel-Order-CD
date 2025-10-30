/**
 * Shows the Dashboard view
 */
function showDashboard() {
  try {
    const template = HtmlService.createTemplateFromFile('dashboard');
    return template.evaluate()
      .setTitle('Travel Orders Dashboard')
      .setWidth(1200)
      .setHeight(700);
  } catch (error) {
    Logger.log('Error in showDashboard: ' + error.message);
    return HtmlService.createHtmlOutput(
      '<html><body><h2>Error Loading Dashboard</h2>' +
      '<p>' + error.message + '</p>' +
      '<p>Please check that dashboard.html exists and is properly formatted.</p></body></html>'
    ).setTitle('Error');
  }
}

/**
 * Shows the New Travel Order form
 */
function showNewToForm() {
  try {
    const template = HtmlService.createTemplateFromFile('newTo');
    return template.evaluate()
      .setTitle('New Travel Order')
      .setWidth(800)
      .setHeight(700);
  } catch (error) {
    Logger.log('Error in showNewToForm: ' + error.message);
    return HtmlService.createHtmlOutput(
      '<html><body><h2>Error Loading Form</h2>' +
      '<p>' + error.message + '</p></body></html>'
    ).setTitle('Error');
  }
}

/**
 * Shows the Update Travel Order form
 */
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
    return HtmlService.createHtmlOutput(
      '<html><body><h2>Error Loading Update Form</h2>' +
      '<p>' + error.message + '</p></body></html>'
    ).setTitle('Error');
  }
}

/**
 * Includes HTML files (for partials like nav.html)
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    Logger.log('Error in include(' + filename + '): ' + error.message);
    return '<div style="color: red;">Error loading ' + filename + ': ' + error.message + '</div>';
  }
}
