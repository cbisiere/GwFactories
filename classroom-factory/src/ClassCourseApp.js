/**
 * Class CourseApp
 *
 * An application to manage classrooms from a spreadsheet.
 *
 * Christophe Bisière
 *
 * version 2020-12-13
 *
 */

/**
 * The course add-on.
 *
 */
class CourseApp {
  /**
   * Apply a function on the class table present in the current worksheet
   *
   * TODO: if many tables, select the one intersecting with the cursor, or
   *  take the first one
   *
   */
  static applyOnClassData(f, r) {
    try {
      if (r == undefined) {
        /* no range provided: lookup for a table */
        r = CourseTable.locate(SpreadsheetApp.getActiveSheet());
        if (r == null) {
          throw new Error(
              [
                'Cannot find any course data table in this worksheet.',
                'Aucune table de données de cours dans cette feuille.',
              ]);
        }
      }
      Logger.log("Course table at coordinates %s", r.getA1Notation());

      const cTable = new CourseTable(r);

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
  * Create Classes
  *
  */
  static doCreateClasses(cTable) {

    let creationCount = 0;
    let errorCount = 0;

    [creationCount, errorCount] = cTable.create();

    LF.SheetHelper.alertUser(LF.i18n(
        [
          'Class creation done: ' + creationCount +
              ' classe(s) created successfully, ' + errorCount + ' error(s).',
          'Création des classes terminée : ' + creationCount +
              ' classe(s) créé(s) avec succès, ' + errorCount + ' erreur(s).',
        ]));
  }

  /**
   * Refresh class data
   *
   */
  static doRefreshClasses(cTable) {

    if (cTable.getNumCourses() == 0) {
      throw new Error(LF.i18n([
        'The course table is empty.',
        'la table des cours est vide.',
      ]));
    }

    if (!cTable.has(COL_CLASS_ID)) {
      throw new Error(LF.i18n([
        'colonne manquante: \'' + COL_CLASS_ID + '\'.',
        'colonne manquante: \'' + COL_CLASS_ID + '\'.',
      ]));
    }
    cTable.refresh();
  }

  /**
   * Load class data
   *
   * do NOT delete row of courses that are not updated
   */
  static doLoadClasses(cTable) {
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
    testGetTeachers();
  }

  static insertHeaderAt(c) {
    c.setValue(COL_CLASS_NAME);
    this.applyOnClassData(this.doComplete, c);
  }


  static insert() {
    let c = SpreadsheetApp.getActiveSheet().getSelection().getCurrentCell();
    this.insertHeaderAt(c);
  }

  static create() {
    Logger.log("CREATE CALL");
    this.applyOnClassData(this.doCreateClasses);
  }

  static refresh() {
    Logger.log("REFRESH CALL");
    this.applyOnClassData(this.doRefreshClasses);
  }

  static load() {
    Logger.log("LOAD CALL");
    this.applyOnClassData(this.doLoadClasses);
  }

  static sample() {
    Logger.log("SAMPLE CALL");
    this.applyOnClassData(this.doAppendSample);
  }

  static test() {
    Logger.log("TEST CALL");
    this.applyOnClassData(this.doTest);
  }
}
