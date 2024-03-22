/**
 * Class DocTable
 *
 * A DocTable is a table containing data about Documents to create.
 *
 * Christophe Bisière
 *
 * version 2022-07-25
 *
 */

/*
 * Actions
 */

const ACTION_NONE = 'Skip';
const ACTION_CREATE = 'Create document';
const ACTION_UPDATE = 'Update document';
const ACTION_CONTENT = 'Update content';
const ACTION_REFRESH = 'Refresh data';

const ACTIONS = [ACTION_NONE, ACTION_CREATE, ACTION_UPDATE, ACTION_CONTENT, ACTION_REFRESH];

/*
 * Columns in the document table
 */

/* action to perform */
const COL_ACTION = 'Action';

/* fields used when creating a new Document */
const COL_DOCUMENT_NAME = 'Document Name'; /* name of the target document */
const COL_DOCUMENT_MODEL_ID = 'Document Template Id'; /* id of the model document */
const COL_DOCUMENT_OWNER = 'Document Owner'; /* owner of that document   */
const COL_DOCUMENT_EDITORS = 'Document Editors'; /* list of editors of that document   */
const COL_DOCUMENT_VIEWERS = 'Document Viewers'; /* list of viewers of that document   */
const COL_DOCUMENT_COMMENTERS = 'Document Commenters'; /* list of commenters of that document   */
const COL_DOCUMENT_FORMAT = 'Document Format'; /* output format: 'gdoc', 'pdf', 'docx'  */
const COL_DOCUMENT_FOLDER_ID = 'Document Folder Id'; /* id of the output folder */

/* read-only fields */
const COL_DOCUMENT_ID = 'Document Id'; /* id of the output document (set by the script) */
const COL_DOCUMENT_URL = 'Document URL'; /* URL of the output document (set by the script) */

/* status info */
const COL_STATUS = 'Status'; /* status of the merge (set by the script or set by the user) */
const COL_TIMESTAMP = 'Timestamp'; /* timestamp of the status (set by the script) */

/* list of supported format, to be used by a static method */
const SUPPORTED_FORMATS = ['pdf', 'gdoc', 'docx'];


/**
 * Class representing Documents.
 *
 */
class DocTable {

  /**
   * Create a DocTable.
   * @param {Range} r - The Range for the whole table.
   */
  constructor(r) {
    /* default output type */
    this.DEFAULT_OUTPUT_TYPE = 'gdoc';

    /* supported mime types. */
    this.MIME_TYPES = new Map([
      ['pdf', 'application/pdf'],
      ['gdoc', 'application/vnd.google-apps.document'],
      ['docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document']
    ]);
    /* ditto, reversed (for convenience): index is mime type */
    this.OUTPUT_FORMATS = new Map(Array.from(this.MIME_TYPES, row => row.reverse()));

    this.cols = [
      COL_ACTION,
      COL_DOCUMENT_NAME,
      COL_DOCUMENT_MODEL_ID,
      COL_DOCUMENT_OWNER,
      COL_DOCUMENT_EDITORS,
      COL_DOCUMENT_VIEWERS,
      COL_DOCUMENT_COMMENTERS,
      COL_DOCUMENT_FORMAT,
      COL_DOCUMENT_FOLDER_ID,
      COL_DOCUMENT_ID,
      COL_DOCUMENT_URL,
      COL_STATUS,
      COL_TIMESTAMP,
    ];

    const defaults = new Map([
      [COL_ACTION, ACTION_UPDATE],
      [COL_DOCUMENT_OWNER, 'me'],
      [COL_DOCUMENT_FORMAT, this.DEFAULT_OUTPUT_TYPE],
    ]);

    const formats = new Map([
      [COL_ACTION, DocTable.setRangeFormatAsAction],
      [COL_DOCUMENT_NAME, LF.SheetHelper.setRangeFormatAsText],
      [COL_DOCUMENT_MODEL_ID, LF.SheetHelper.setRangeFormatAsText],
      [COL_DOCUMENT_OWNER, LF.SheetHelper.setRangeFormatAsText],
      [COL_DOCUMENT_EDITORS, LF.SheetHelper.setRangeFormatAsText],
      [COL_DOCUMENT_VIEWERS, LF.SheetHelper.setRangeFormatAsText],
      [COL_DOCUMENT_COMMENTERS, LF.SheetHelper.setRangeFormatAsText],
      [COL_DOCUMENT_FORMAT, DocTable.setRangeFormatAsDocumentFormat],
      [COL_DOCUMENT_FOLDER_ID, LF.SheetHelper.setRangeFormatAsText],
      [COL_DOCUMENT_ID, LF.SheetHelper.setRangeFormatAsText],
      [COL_DOCUMENT_URL, LF.SheetHelper.setRangeFormatAsText],
      [COL_TIMESTAMP, LF.SheetHelper.setRangeFormatAsDatetime],
      [COL_STATUS, LF.SheetHelper.setRangeFormatAsText],
    ]);

    this.dt = new LF.DataTable(r, defaults, formats);
  }


  /* static members */

  /**
   * Locate the target Document table in a Sheet
   *
   * @param {Sheet} sheet The Sheet containing the Document table to search for.
   * @return {?Range} The Range of the Classroom table found.
   */
  static locate(sheet) {
    return LF.DataTable.locateFromLabel(sheet, COL_DOCUMENT_MODEL_ID);
  }

  /**
   * Format a cell as an action if it contains a valiid action.
   * 
   * We do not set the format if the current value is invalid, as
   * it raises an exception that seems to be uncatchable.
   *
   * @param {Range} c - The cell to format as an action.
   */
  static setRangeFormatAsAction(c) {
    if (ACTIONS.includes(String(c.getValue()).trim())) {
      LF.SheetHelper.setRangeFormatAsList(c, ACTIONS);
    }
  }

  /**
   * Format a cell as a document format.
   *
   * We do not set the format if the current value is invalid, as
   * it raises an exception that seems to be uncatchable.
   * 
   * @param {Range} c - The cell to format as document format.
   */
  static setRangeFormatAsDocumentFormat(c) {
    if (SUPPORTED_FORMATS.includes(String(c.getValue()).trim())) {
      LF.SheetHelper.setRangeFormatAsList(c, SUPPORTED_FORMATS);
    }
  }

  /**
   * Check a condition.
   *
   * @param {boolean} condition - The condition to check.
   * @param {string} message - The message to display if the condition is not met.
   */
  assert(condition, message) {
    const a1 = this.dt.getRange() == undefined ? 'undefined' :
      this.dt.getRange().getA1Notation(); /* "==" so also null */
    const prompt = 'Error: DocTable (' + a1 + '): ' +
      (message || ' assertion failed');
    assert(condition, prompt);
  }

  /*
   * Returns a map from a Google file object.
   *
   * @param {File} file - The file object.
   * @param {string[]} labels - The column labels to consider.
   * @return {Map} - The map of document properties.
   */
  getMapFromFileObject(file, labels) {
    const dmap = new Map();
    for (const col of labels) {
      let v;
      switch (col) {
        case COL_DOCUMENT_NAME:
          v = file.getName();
          break;
        case COL_DOCUMENT_ID:
          v = file.getId();
          break;
        case COL_DOCUMENT_URL:
          v = file.getUrl();
          break;
        case COL_DOCUMENT_OWNER:
          v = file.getOwner().getEmail();
          break;
        case COL_DOCUMENT_EDITORS:
          v = file.getEditors().map(user => user.getEmail()).join(LF.DATATABLE_LIST_SEPARATOR);
          break;
        case COL_DOCUMENT_VIEWERS:
          v = file.getViewers().map(user => user.getEmail()).join(LF.DATATABLE_LIST_SEPARATOR);
          break;
        case COL_DOCUMENT_COMMENTERS:
          v = LF.DriveHelper.getCommenters(file).map(user => user.getEmail()).join(LF.DATATABLE_LIST_SEPARATOR);
          break;
        case COL_DOCUMENT_FOLDER_ID:
          v = LF.DriveHelper.getFirstParent(file).getId();
          break;
        case COL_DOCUMENT_FORMAT:
          const mime = file.getMimeType();
          v = this.OUTPUT_FORMATS.has(mime) ? this.OUTPUT_FORMATS.get(mime) : 'unsupported';
          break;
        default:
          /* ignore columns not in the list, as they may be user defined columns 
           or columns that are not file properties */
          break;
      }
      if (v !== undefined) {
        dmap.set(col, v);
      }
    }
    return dmap;
  }

  /*
   * Execute a merge operation using values defined in a a map.
   *
   * TODO: cleanup created files in case of error
   * 
   * @param {Map} dmap - The document map to use to create the merged document.
   * @param {boolean} in_place - Update the existing output document, keeping the same Google id.
   * @param {boolean} set_props - Set the document properties, besides its content.
   * @return {File} file - The file for the merged document. // TODO: return indication it has been created or updated
   */
  mergeFromMap(dmap, in_place, set_props) {
    const activeUser = Session.getActiveUser();

    /* model */
    let modelId = LF.getValue(COL_DOCUMENT_MODEL_ID, dmap); //TODO: accept URL, extract ID
    if (modelId === false) {
      throw new LF.DataTableCellError(COL_DOCUMENT_MODEL_ID, LF.i18n([
        'Missing template. A template id must be provided in column "' + COL_DOCUMENT_MODEL_ID + '".',
        'template manuqante. Un idenfiant de template doit être saisi dans la colonne "' + COL_DOCUMENT_MODEL_ID + '".',
      ]));
    }
    let modelFile;
    try {
      modelFile = DriveApp.getFileById(modelId);
    } catch (e) {
      throw new LF.DataTableCellError(COL_DOCUMENT_MODEL_ID, LF.i18n([
        "Cannot access the template file: " + e,
        "impossible d'accéder au document template : " + e,
      ]));
    }
    Logger.log('model file: "%s" (%s)', modelFile, modelFile.getId());

    /* target file name, defaulting to model name */
    let fileName = LF.getValue(COL_DOCUMENT_NAME, dmap);
    if (fileName === false) {
      fileName = modelFile.getName(); //FIXME: remove extention if any?
    }

    /* file format */
    let fileFormat = LF.getValue(COL_DOCUMENT_FORMAT, dmap);
    if (fileFormat === false) {
      fileFormat = this.DEFAULT_OUTPUT_TYPE;
    }
    Logger.log('document format: "%s"', fileFormat);
    if (!this.MIME_TYPES.has(fileFormat)) {
      throw new LF.DataTableCellError(COL_DOCUMENT_FORMAT, LF.i18n([
        'Invalid or unsupported document format: "' + fileFormat + '".',
        'format de document invalide ou non supporté : "' + fileFormat + '".'
      ]));
    }

    /* owner, defaulting to the user runing the script */
    const ownerVal = LF.getValue(COL_DOCUMENT_OWNER, dmap)
    const owner = ownerVal !== false && ownerVal !== 'me' ? ownerVal : activeUser.getEmail();
    Logger.log('owner: "%s"', owner);

    /* list of editors */
    const editors = LF.DataTable.getValueAsList(COL_DOCUMENT_EDITORS, dmap);
    Logger.log('editors: "%s"', editors);

    /* list of viewers */
    const viewers = LF.DataTable.getValueAsList(COL_DOCUMENT_VIEWERS, dmap)
    Logger.log('viewers: "%s"', viewers);

    /* list of commenters */
    const commenters = LF.DataTable.getValueAsList(COL_DOCUMENT_COMMENTERS, dmap)
    Logger.log('commenters: "%s"', commenters);

    /* target folder, defaulting to the parent folder of this sheet */
    const folderId = LF.getValue(COL_DOCUMENT_FOLDER_ID, dmap);
    let folder;
    if (folderId === false) {
      folder = LF.DriveHelper.getFirstParent(SpreadsheetApp.getActiveSpreadsheet());
    } else {
      try {
        folder = DriveApp.getFolderById(folderId);
      } catch (e) {
        throw new LF.DataTableCellError(COL_DOCUMENT_FOLDER_ID, LF.i18n([
          "Cannot access destination folder: " + e,
          "impossible d'accéder au dossier de destination : " + e,
        ]));
      }
    }
    Logger.log('folder: "%s" (%s)', folder, folder.getId());

    /* is the target folder trashed? */
    if (folder.isTrashed()) {
      throw new LF.DataTableCellError(COL_DOCUMENT_FOLDER_ID, LF.i18n([
        "Destination folder '" + folder.getName() + "' is trashed.",
        "le dossier de destination '" + folder.getName() + "' est dans la corbeille."
      ]));
    }

    /* previous and current output documents */
    let prevFile = null;
    let targetFile = null;
    const prevFileId = LF.getValue(COL_DOCUMENT_ID, dmap);

    if (in_place) {
      if (prevFileId === false) {
        throw new LF.DataTableCellError(COL_DOCUMENT_ID, LF.i18n([
          'The id of the previous document (column "' + COL_DOCUMENT_ID + '") is missing. Consider using the action "' + ACTION_CREATE + '" instead.',
          'l\'id du document précédent (colonne "' + COL_DOCUMENT_ID + '") est manquant. Utilisez plutôt l\'action "' + ACTION_CREATE + '".',
        ]));
      }
      try {
        prevFile = DriveApp.getFileById(prevFileId);
      } catch (e) {
        /* previous file does not exist anymore and thus id cannot be reused */
        throw new LF.DataTableCellError(COL_DOCUMENT_ID, LF.i18n([
          'Previous document (id ' + prevFileId + ') does not exist and thus cannot be updated. You may delete the id and run again.',
          'le document (id ' + prevFileId + ') n\'existe pas et ne peut pas être mis à jour. Vous devez effacer cet id et relancer.'
        ]));
      }
      /* to update a file we make sure it is not trashed - TODO: check whether this can be done after working on its content */
      if (prevFile.isTrashed()) {
        prevFile.setTrashed(false);
      }
      Logger.log('previous document: "%s" (%s)', prevFile, prevFile.getId());
      /* we are going to reuse this output file, keeping the same id */
      targetFile = prevFile;
    } else {
      /* duplicate the model */
      Logger.log("Copying template \"" + modelFile.getName() + "\"");
      targetFile = modelFile.makeCopy();
    }
    Logger.log('target document: "%s" (%s)', targetFile, targetFile.getId());

    /* Open the target document */
    Logger.log("Opening file \"" + targetFile.getName() + "\"");
    let targetDocument = DocumentApp.openById(targetFile.getId()); /* FIXME: cannot open a pdf: manage versions? https://developers.google.com/drive/api/guides/change-overview */

    Logger.log("Merging document...");
    MergeHelper.merge(targetDocument, dmap);

    Logger.log("Saving document...");
    targetDocument.saveAndClose();

    /* Convert to a different mimetype when requested */
    /* TODO: extention? */
    if (fileFormat != 'gdoc') {
      const mime = this.MIME_TYPES.get(fileFormat);
      const blob = LF.DriveHelper.getBlobAs(targetFile.getId(), mime);
      targetFile.setContent(blob);

      const convertedFile = DriveApp.createFile(blob);
      Logger.log("Created file \"" + convertedFile.getName() + "\"");
      Logger.log("Trashing file \"" + targetFile.getName() + "\"");
      targetFile.setTrashed(true);
      targetFile = convertedFile;

    }


    /*
     * Set the various properties of this file, as follows:
     */

    /* 0) cleanup location and access right properties inherited from the model */
    LF.DriveHelper.removeAllParentsFromFile(targetFile);
    LF.DriveHelper.removeAllViewersFrom(targetFile, activeUser);
    LF.DriveHelper.removeAllCommentersFrom(targetFile, activeUser);
    LF.DriveHelper.removeAllEditorsFrom(targetFile, activeUser);

    /* 1) copy properties from the previous file 
    TRICKY. WHICH PROPERTIES SHOULD WE KEEP? ON REQUEST?
    if (prevfile != null && !prevfile.isTrashed()) {
      LF.DriveHelper.copyFilePropertiesTo(prevfile, newFile);
    } */

    /* 2) enforce file name */
    Logger.log("Setting name of file \"" + targetFile + "\" to \"" + fileName + "\"");
    targetFile.setName(fileName);

    /* 3) enforce access rights */
    for (const email of commenters) {
      Logger.log('Adding commenter %s to file "%s"', email, targetFile);
      LF.DriveHelper.addCommenterQuiet(email, targetFile);
    }
    for (const email of viewers) {
      Logger.log('Adding viewer %s to file "%s"', email, targetFile);
      LF.DriveHelper.addViewerQuiet(email, targetFile);
    }
    for (const email of editors) {
      Logger.log('Adding editor %s to file "%s"', email, targetFile);
      LF.DriveHelper.addEditorQuiet(email, targetFile);
    }

    /* move to target folder */
    Logger.log('Adding file "%s" to parent folder "%s"', targetFile.getName(), folder.getName());
    folder.addFile(targetFile);

    /* enforce owner */
    Logger.log('Assigning owner %s to file "%s"', owner, targetFile);
    LF.DriveHelper.setOwnerQuiet(owner, targetFile);
    Logger.log('Done');

    if (!in_place && prevFileId !== false) {
      /* silently trash the old file */
      LF.DriveHelper.trashByIdNoFail(prevFileId);
    }

    return targetFile;
  }

  /**
   * Execute the actions in COL_ACTION
   *
   */
  run() {
    /* start date */
    const now = new Date();

    /* create action counters */
    const count = new Map();
    var errorCount = 0;

    /* columns that needs to be non empty to execute an action */
    const requires = new Map([
      [ACTION_NONE, []],
      [ACTION_CREATE, [COL_DOCUMENT_MODEL_ID]],
      [ACTION_UPDATE, [COL_DOCUMENT_MODEL_ID, COL_DOCUMENT_ID]],
      [ACTION_CONTENT, [COL_DOCUMENT_MODEL_ID, COL_DOCUMENT_ID]],
      [ACTION_REFRESH, [COL_DOCUMENT_ID]],
    ]);

    let rows = this.dt.getDataAsMaps();
    for (var [i, map] of rows) {

      try {
        LF.trimStringsInMap(map);
        this.dt.resetRowColor(i);
        
        //const map = new Map(row);
        this.dt.applyDefaultsToMap(map);

        const action = map.get(COL_ACTION);
        if (!ACTIONS.includes(action)) {
          throw new LF.DataTableCellError(COL_ACTION, LF.i18n([
            `Invalid action "${action}": allowed actions are ${ACTIONS}`,
            `Action invalide "${action}" : les actions possibles sont ${ACTIONS}`,
          ]));
        }

        /* check all the required columns for this action are not empty */
        for (const col of requires.get(action)) {
          if (LF.getValue(col, map) === false) {
            throw new Error(LF.i18n([
              `Missing value in column "${col}"".`,
              `Valeur attendue dans la colonne "${col}"".`,
            ]));
          }
        }

        let file;
        if (action != ACTION_NONE) {
          /* execute the action and get the resulting output document */
          if ([ACTION_CREATE, ACTION_UPDATE, ACTION_CONTENT].includes(action)) {
            const in_place = (action != ACTION_CREATE);
            const set_props = (action != ACTION_CONTENT);
            file = this.mergeFromMap(map, in_place, set_props);
          } else if (action == ACTION_REFRESH) {
            const id = LF.getValue(COL_DOCUMENT_ID, map);
            try {
              file = DriveApp.getFileById(id);
            } catch (e) {
              throw new LF.DataTableCellError(COL_DOCUMENT_ID, LF.i18n([
                `Cannot access document file with id ${id}: ${e}`,
                `impossible d'accéder au document d\'id ${id} : ${e}`,
              ]));
            }
          }

          /* only update the table cells that need to be updated */
          const fileProps = this.getMapFromFileObject(file, this.dt.getHeaderAsVector()); //FIXME: filter out inherited properties? (e.g. editors)
          const changed = LF.mapDiff(map, fileProps);
          changed.set(COL_STATUS, LF.i18n([
            `Action "${action}" executed`,
            `Action "${action}" exécutée`,
          ]));
          changed.set(COL_TIMESTAMP, now);
          changed.set(COL_ACTION, ACTION_NONE);
          this.dt.updateRow(i, changed);

          LF.inc(count, action);
        }
      } catch (e) {
        Logger.log("ERROR: %s", e);
        /* update status */
        const status = new Map([
          [COL_STATUS, e.message],
          [COL_TIMESTAMP, now],
        ]);
        this.dt.updateRow(i, status);
        /* set font color of the error cell */
        if (e.name == 'DataTableCellError') {
          this.dt.setErrorColor(i, e.label);
        }
        errorCount = errorCount + 1;
      }
    }
    return [count, errorCount];
  }

  /**
   * Append a sample row
   *
   */
  sample() {
    /* create the sample model document in the same folder as the spreadsheet */
    const doc = createSampleModel(Session.getActiveUserLocale());
    const folder = LF.DriveHelper.getFirstParent(SpreadsheetApp.getActiveSpreadsheet());
    const file = DriveApp.getFileById(doc.getId());
    folder.addFile(file);

    const userColName = LF.i18n(['Champ 1', 'Champ 1']);
    const m = new Map([
      [COL_ACTION, ACTION_CREATE],
      [COL_DOCUMENT_NAME, LF.i18n(['Merged document', 'Document fusionné'])],
      [COL_DOCUMENT_MODEL_ID, file.getId()],
      [COL_DOCUMENT_FORMAT, this.DEFAULT_OUTPUT_TYPE],
      [userColName, LF.i18n(['Hello', 'Bonjour'])]
    ]);
    this.dt.ensureColumnsExist(m.keys());
    this.dt.addRowFromMap(m);
  }

  /**
   * Complete the table with some missing columns
   *
   * @param {boolean} mini True to insert a reduced set of columns only.
   */
  complete(mini=false) {
    const miniCols = [
      COL_ACTION,
      COL_DOCUMENT_NAME,
      COL_DOCUMENT_MODEL_ID
    ];
    this.dt.ensureColumnsExist(mini ? miniCols : this.cols);
  }


  /*
  * convenient accessors to the DataTable member, as ES6 modules are not yet 
  * supported in V8 engine. 
  */

  /**
   * Checks the table has at least one row, raising an exception if the table is empty.
   *
   * @param {string} message - The error message attached to the exception.
   */
  checkNonEmpty(message) {
    this.dt.checkNonEmpty(message);
  }

  /**
   * Checks whether some columns exist in the table.
   *
   * @param {string[]} labels - The column labels.
   * @param {string} prompt - The prompt for "Missing column:".
   */
  checkColumnsExist(labels, prompt) {
    return this.dt.checkColumnsExist(labels, prompt);
  }

  /**
   * Ensure some columns exist in the table.
   *
   * @param {string[]} labels - The column labels.
   */
  ensureColumnsExist(labels) {
    this.dt.ensureColumnsExist(labels);
  }

}
