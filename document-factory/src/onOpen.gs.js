/**
 * Create a Scripts menu for easy access
 */

MENU_DOCUMENT_FACTORY = LF.i18n(['Document factory', 'Document factory']);

MENU_RUN = LF.i18n(['Execute actions', 'Effectuer les actions']);

MENU_INSERT = LF.i18n(['Insert a table', 'Insérer une table']);
MENU_INSERT_SMALL = LF.i18n(['with the most common columns', 'avec les colonnes les plus courantes']);
MENU_INSERT_LARGE = LF.i18n(['with all the columns', 'avec toutes les colonnes']);

MENU_ADD = LF.i18n(['Add to a table', 'Ajouter à une table']);
MENU_ADD_ALL = LF.i18n(['all the columns', 'toutes les colonnes']);
MENU_ADD_SAMPLE = LF.i18n(['a sample row', 'un example de ligne']);

MENU_SELECT = LF.i18n(['Select a table', 'Sélectionner une table']);

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu(MENU_DOCUMENT_FACTORY)
      .addItem(MENU_RUN, 'DocApp.run')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu(MENU_INSERT)
          .addItem(MENU_INSERT_SMALL, 'DocApp.insertHeaderMini')
          .addItem(MENU_INSERT_LARGE, 'DocApp.insertHeaderMaxi'))
      .addSubMenu(SpreadsheetApp.getUi().createMenu(MENU_ADD)
          .addItem(MENU_ADD_ALL, 'DocApp.completeHeader')
          .addItem(MENU_ADD_SAMPLE, 'DocApp.sample'))
      .addItem(MENU_SELECT, 'DocApp.select')
      .addToUi();
};

function insertHeadert() {
  DocApp.insertHeader(false);
}

function insertHeaderMini() {
  DocApp.insertHeader(true);
}