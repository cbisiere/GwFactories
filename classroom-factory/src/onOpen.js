/**
 * Create a Scripts menu for easy access 
 */

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
  {
    name : LF.i18n(["Run the class creation tool CreateClasses","Lancer l'outil de création de classes CreateClasses"]),
    functionName : "CourseApp.create"
  },
  {
    name : LF.i18n(["Run the class load tool LoadClasses","Lancer l'outil de lecture de classes LoadClasses"]),
    functionName : "CourseApp.load"
  },
  {
    name : LF.i18n(["Refresh a Classroom data table","Rafraîchir les données d'une table Classroom"]),
    functionName : "CourseApp.refresh"
  },
  {
    name : LF.i18n(["Insert a Classroom data table","Insérer une table Classroom"]),
    functionName : "CourseApp.insert"
  },
  {
    name : LF.i18n(["Add a sample course","Ajouter un example de cours"]),
    functionName : "CourseApp.sample"
  },
  {
    name : LF.i18n(["TEST CourseApp","TEST CourseApp"]),
    functionName : "CourseApp.test"
  },
  ];
  spreadsheet.addMenu(LF.i18n(["Classroom Factory","Classroom Factory"]), entries);
};
