/**
 * Class DocApp
 *
 * An application to produce documents from a spreadsheet.
 *
 * Christophe Bisière
 *
 * version 2022-07-24
 *
 */

/**
 * The Document add-on.
 *
 */
class DocApp {
  /**
   * Apply a function on the document table present in the current worksheet
   *
   * TODO: if many tables, select the one intersecting with the cursor, or
   *  take the first one
   * 
   * @param {function} f The function to call.
   * @param {Range?} r The Range for the whole table.
   *
   */
  static apply(f, r) {
    try {
      if (r == undefined) {
        /* no range provided: lookup for a table */
        r = DocTable.locate(SpreadsheetApp.getActiveSheet());
        if (r == null) {
          throw new Error(LF.i18n([
            'Cannot find any document table in this worksheet.',
            'Aucune table de documents dans cette feuille.',
          ]));
        }
      }
      Logger.log('Document table at coordinates %s', r.getA1Notation());

      const dTable = new DocTable(r);

      f(dTable);

    } catch (e) {
      LF.SheetHelper.alertUser(LF.i18n(
        [
          'Error: ' + e,
          'Erreur : ' + e,
        ]));
    }
  }

  /**
   * Check the table is not empty
   *
   * @param {DocTable} dTable The data table to check.
   */
  static checkNonEmpty(dTable) {
    dTable.checkNonEmpty(LF.i18n([
      'The document table is empty.',
      'la table des documents est vide.',
    ]));
  }

  /**
   * Checks several columns exist in the table, raising an exception on the first column that does not exist.
   *
   * @param {DocTable} dTable The data table to check.
   * @param {string[]} labels - The column labels.
   */
  static checkColumnsExist(dTable, labels) {
    dTable.checkColumnsExist(labels, LF.i18n([
      'Missing column:',
      'colonne manquante :',
    ]));
  }

  /**
  * Run all actions
  *
  */
  static run() {
    Logger.log("RUN CALL");
    this.apply(this.doRun);
  }

  static doRun(dTable) {

    DocApp.checkNonEmpty(dTable);
    DocApp.checkColumnsExist(dTable, [COL_ACTION]);
    dTable.ensureColumnsExist([COL_DOCUMENT_ID, COL_DOCUMENT_URL, COL_DOCUMENT_MODEL_ID, COL_STATUS, COL_TIMESTAMP]);

    let [count, errorCount] = dTable.run();

    /* prepare string containing counters to display */
    let ar = [];
    const spacing = LF.i18n(['', ' ']);
    for (const [action, n] of count) {
      if (n > 0) {
        ar.push(`${action}${spacing}: ${n}`);
      }
    }
    let s = ar.join(', ');
    if (s.length == 0) {
      s = LF.i18n(['None', 'aucune']);
    } else {
      s = `[${s}]`;
    }

    LF.SheetHelper.alertUser(LF.i18n(
      [
        'Done: ' +
        'successfuly executed actions: ' + s + '; ' +
        errorCount + ' error(s).',
        'Exécution terminée : ' +
        'actions exécutée(s) avec succès : ' + s + ' ; ' +
        errorCount + ' erreur(s).',
      ]));
  }


  /**
   * Insert a table at a specific location
   *
   * TODO: check there is nothing there!
   * 
   * @param {Range} c The target cell where to insert the first column.
   * @param {boolean} mini True to insert a reduced set of columns only.
   */
  static insertHeaderAt(c, mini=false) {
    c.setValue(COL_ACTION);
    if (mini) {
      this.apply(this.doCompleteHeaderMini, c);
    } else {
      this.apply(this.doCompleteHeader, c);
    }
  }

  /**
   * Insert a table at the current cell
   * 
   * @param {boolean} mini True to insert a reduced set of columns only.
   */
  static insertHeader(mini = false) {
    let c = SpreadsheetApp.getActiveSheet().getSelection().getCurrentCell();
    this.insertHeaderAt(c, mini);
  }

  /**
   * Insert a table with only the most useful columns 
   * 
   */
  static insertHeaderMini() {
    this.insertHeader(true);
  }

  /**
   * Insert a table with all the columns 
   * 
   */
  static insertHeaderMaxi() {
    this.insertHeader(false);
  }

  /**
   * Add missing columns
   *
   */
  static completeHeader() {
    Logger.log("COMPLETE CALL");
    this.apply(this.doCompleteHeader);
  }

  static doCompleteHeader(dTable) {
    dTable.complete(false);
  }
  static doCompleteHeaderMini(dTable) {
    dTable.complete(true);
  }

  /**
   * Append a sample row
   *
   */
  static sample() {
    Logger.log("SAMPLE CALL");
    //    LF.DataTable.apply(this.doAppendSample, null, DocTable);
    this.apply(this.doAppendSample);
  }

  static doAppendSample(dTable) {
    dTable.sample();
  }

}
