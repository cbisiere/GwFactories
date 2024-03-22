/**
 * Create a Scripts menu for easy access
 */

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
  {
    name : LF.i18n(["Run the folder creation tool CreateFolders","Lancer l'outil de création de dossiers CreateFolders"]),
    functionName : "FolderApp.create"
  },
  null,
  {
    name : LF.i18n(["Run the class load tool LoadFolders","Lancer l'outil de lecture de classes LoadFolders"]),
    functionName : "FolderApp.load"
  },
  {
    name : LF.i18n(["Refresh a folder data table","Rafraîchir les données d'une table de dossiers"]),
    functionName : "FolderApp.refresh"
  },
  {
    name : LF.i18n(["Insert a folder data table","Insérer une table de dossiers"]),
    functionName : "FolderApp.insert"
  },
  {
    name : LF.i18n(["Add a sample folder","Ajouter un example de dossier"]),
    functionName : "FolderApp.sample"
  },
  {
    name : LF.i18n(["TEST FolderApp","TEST FolderApp"]),
    functionName : "FolderApp.test"
  }
  ];
  spreadsheet.addMenu(LF.i18n(["Folder factory","Folder factory"]), entries);
};
