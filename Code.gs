/**
 * --------------------------------------------------------------------------
 * Palia Fish Assistant
 * Code.gs
 *
 * Application entry point.
 *
 * Responsibilities:
 *   • Create custom menu
 *   • Open sidebar
 *   • Load HTML partials
 *   • Store user preferences
 *
 * No spreadsheet searching is performed here.
 * --------------------------------------------------------------------------
 */

/**
 * Runs automatically whenever the spreadsheet opens.
 */
function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu(APP_SETTINGS.APP_NAME)
    .addItem("Open Assistant", "showSidebar")
    .addSeparator()
    .addItem("Refresh Assistant", "refreshAssistant")
    .addSeparator()
    .addItem("Settings", "showSettings")
    .addToUi();

}


/**
 * Opens the main sidebar.
 */
function showSidebar() {

  try {

    const html =
      HtmlService
        .createTemplateFromFile("Sidebar")
        .evaluate()
        .setTitle(APP_SETTINGS.APP_NAME);

    SpreadsheetApp
      .getUi()
      .showSidebar(html);

  }

  catch (err) {

    SpreadsheetApp
      .getUi()
      .alert(err.message);

  }

}


/**
 * Placeholder.
 *
 * Later this will rebuild any runtime information
 * such as recipe discovery.
 *
 * Because the spreadsheet is the source of truth,
 * no permanent cache is maintained.
 */
function refreshAssistant() {

  SpreadsheetApp.getActiveSpreadsheet()
    .toast(
      "Assistant refreshed.",
      APP_SETTINGS.APP_NAME,
      3
    );

}


/**
 * Opens the settings page.
 *
 * Initially this opens the same sidebar.
 * Later it will switch to the Settings tab.
 */
function showSettings() {

  showSidebar();

}


/**
 * Includes HTML partials.
 *
 * Example:
 *
 * <?!= include('Styles'); ?>
 */
function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}

function onInstall() {
  onOpen();
}
