/**
 * Class FolderApp
 *
 * An application to manage folders from a spreadsheet.
 *
 * Christophe Bisière
 *
 * version 2022-04-18
 *
 */

/**
 * The folder add-on.
 *
 */
class FolderApp {

  /**
   * Apply a function on the folder table present in the current worksheet
   *
   * TODO: if many tables, select the one intersecting with the cursor, or
   *  take the first one
   *
   */
  static applyOnFolderData(f, r) {
    try {
      if (r == undefined) {
        /* no range provided: lookup for a table */
        r = FolderTable.locate(SpreadsheetApp.getActiveSheet());
        if (r == null) {
          throw new Error(
              [
                'Cannot find any folder data table in this worksheet.',
                'Aucune table de données de dossiers dans cette feuille.',
              ]);
        }
      }
      Logger.log('Folder table at coordinates %s', r.getA1Notation());

      const cTable = new FolderTable(r);

      f(cTable);

    } catch (e) {
      LF.SheetHelper.alertUser(LF.i18n(
          [
            'Error: ' + e,
            'Erreur : ' + e,
          ]));
    }
  }

  /**
  * Create folders
  *
  */
  static doCreateFolders(cTable) {

    let creationCount = 0;
    let errorCount = 0;

    [creationCount, errorCount] = cTable.create();

    LF.SheetHelper.alertUser(LF.i18n(
        [
          'Folder creation done: ' + creationCount +
              ' folder(s) created successfully, ' + errorCount + ' error(s).',
          'Création des dossiers terminée : ' + creationCount +
              ' dossier(s) créé(s) avec succès, ' + errorCount + ' erreur(s).',
        ]));
  }

  /**
   * Refresh folder data
   *
   */
  static doRefreshFolders(cTable) {
    if (cTable.dt.getNumRows() == 0) {
      throw new Error(LF.i18n([
        'The folder table is empty.',
        'la table des dossiers est vide.',
      ]));
    }

    if (!cTable.dt.has(COL_FOLDER_ID)) {
      throw new Error(LF.i18n([
        'Missing column:\'' + COL_FOLDER_ID + '\'.',
        'colonne manquante: \'' + COL_FOLDER_ID + '\'.',
      ]));
    }
    cTable.refresh();
  }

  /**
   * Load folder data
   *
   * do NOT delete row of courses that are not updated
   */
  static doLoadFolders(cTable) {
    cTable.update();
  }

  /**
   * Append a sample row
   *
   */
  static doAppendSample(cTable) {
    cTable.sample();
  }

  /**
   * Add missing columns
   *
   */
  static doComplete(cTable) {
    cTable.complete();
  }

  /**
   * Test
   *
   */
  static doTest(cTable) {
    //SpreadsheetApp.getUi().alert(cTable.getRange().getA1Notation());
  }

  static insertHeaderAt(c) {
    c.setValue(COL_FOLDER_NAME);
    this.applyOnFolderData(this.doComplete, c);
  }


  static insert() {
    let c = SpreadsheetApp.getActiveSheet().getSelection().getCurrentCell();
    this.insertHeaderAt(c);
  }

  static create() {
    Logger.log("CREATE CALL");
    this.applyOnFolderData(this.doCreateFolders);
  }

  static refresh() {
    Logger.log("REFRESH CALL");
    this.applyOnFolderData(this.doRefreshFolders);
  }

  static load() {
    Logger.log("LOAD CALL");
    this.applyOnFolderData(this.doLoadFolders);
  }

  static sample() {
    Logger.log("SAMPLE CALL");
    this.applyOnFolderData(this.doAppendSample);
  }

  static test() {
    Logger.log("TEST CALL");
    this.applyOnFolderData(this.doTest);
  }

}
