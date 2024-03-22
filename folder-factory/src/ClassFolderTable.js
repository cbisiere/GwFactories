/**
 * Class FolderTable
 *
 * A FolderTable is a table containing data about folders.
 *
 * Christophe Bisière
 *
 * version 2022-04-18
 *
 */

/**
 * Headers in the Folder table.
 *
 */

/* fields used when creating a new folder */
const COL_FOLDER_NAME = "Folder Name";               /* name of the folder to create inside the root folder */
const COL_PARENT_ID = "Folder Parent";                   /* id of the parent folder */
const COL_FOLDER_OWNER = "Folder Owner";             /* owner of that folder   */
const COL_FOLDER_EDITORS = "Folder Editors";         /* list of editors of that folder   */
const COL_FOLDER_VIEWERS = "Folder Viewers";         /* list of viewers of that folder   */

/* read-only fields */
const COL_FOLDER_ID = "Folder Id";            /* id of the folder (set by the script) */

/* status info */
const COL_FOLDER_STATUS = "Status";


/**
 * Class representing a table of folder data.
 *
 */
class FolderTable { /* extends DataTable { NOT SUPPORTED IN V8 */

  /**
   * Create a FolderTable.
   * @param {Range} r - The Range for the whole table.
   */
  constructor(r) {

    this.cols = [
      COL_FOLDER_NAME,
      COL_PARENT_ID,
      COL_FOLDER_OWNER,
      COL_FOLDER_EDITORS,
      COL_FOLDER_VIEWERS,
      COL_FOLDER_ID,
      COL_FOLDER_STATUS,
    ];

    const defaults = new Map([
      [COL_FOLDER_OWNER, 'me'],
    ]);

    const formats = new Map([
      [COL_FOLDER_NAME, LF.SheetHelper.setRangeFormatAsText],
      [COL_PARENT_ID, LF.SheetHelper.setRangeFormatAsText],
      [COL_FOLDER_OWNER, LF.SheetHelper.setRangeFormatAsText],
      [COL_FOLDER_EDITORS, LF.SheetHelper.setRangeFormatAsText],
      [COL_FOLDER_VIEWERS, LF.SheetHelper.setRangeFormatAsText],
      [COL_FOLDER_ID, LF.SheetHelper.setRangeFormatAsText],
    ]);

    this.dt = new LF.DataTable(r, defaults, formats);
    // super(r, defaults, formats);
  }


  /* static members */

  /**
   * Locate the target table in a Sheet
   *
   * @param {Sheet} sheet - The Sheet containing the Classroom table to search for.
   * @return {?Range} The Range of the Classroom table found.
   */
  static locate(sheet) {
    return LF.DataTable.locateFromLabel(sheet, COL_FOLDER_NAME);
  }

  /**
   * Check a condition.
   *
   * @param {boolean} condition - The condition to check.
   * @param {string} message - The message to display if the condition is not met.
   */
  assert(condition, message) {
    let a1 = this.dt.getRange() == undefined ? 'undefined' : this.dt.getRange().getA1Notation(); /* "==" so also null */
    let prompt = 'Error: FolderTable (' + a1 + '): ' + (message || ' assertion failed');
    assert(condition, prompt);
  }

  /*
   * Update a folder map using a Google folder object.
   *
   * TODO: return a boolean "something updated"
   *
   * @param {Map} fmap - The folder map to update.
   * @param {Folder} folder - The folder object.
   */
  updateMapFromFolderObject(fmap, folder) {

    fmap.set(COL_FOLDER_NAME, folder.getName());
    fmap.set(COL_PARENT_ID, LF.DriveHelper.getFirstParent(folder).getId());
    fmap.set(COL_FOLDER_OWNER, folder.getOwner().getEmail());
    fmap.set(COL_FOLDER_EDITORS, folder.getEditors().map(user => user.getEmail()).join(' '));
    fmap.set(COL_FOLDER_VIEWERS, folder.getViewers().map(user => user.getEmail()).join(' '));
    fmap.set(COL_FOLDER_ID, folder.getId())
  }

  /*
   * Update a row in the folder table, using a map of current data and a folder.
   *
   * TODO: return a boolean "something updated"
   *
   * @param {number} i - The row number.
   * @param {Map} fmap - The folder map to update.
   * @param {Folder} folder - The folder object.
   */
  updateRow(i, fmap, folder) {
    this.updateMapFromFolderObject(fmap, folder);
    this.dt.setRowFromMap(i, fmap);

    let parent = DriveApp.getFolderById(fmap.get(COL_PARENT_ID));
    let v1 = LF.SheetHelper.makeLink(parent.getUrl(), parent.getName());
    this. dt.setRichTextValue(i, COL_PARENT_ID, v1);

    let v2 = LF.SheetHelper.makeLink(folder.getUrl(), folder.getName());
    this. dt.setRichTextValue(i, COL_FOLDER_ID, v2);
  }

  /*
   * Create a new Folder from a map.
   *
   * @param {Map} fmap - The folder map to use to create the new classroom.
   * @return {course} - The new Folder object.
   */
  createFolderFromMap(fmap) {

    /* parent folder */
    let parent = null;
    if (fmap.has(COL_PARENT_ID) && fmap.get(COL_PARENT_ID).length > 0) {
      parent = fmap.get(COL_PARENT_ID);
    } else {
      parent = LF.DriveHelper.getFirstParent(SpreadsheetApp.getActiveSpreadsheet());
    }

    /* create the folder */
    var folder = parent.createFolder(fmap.get(COL_FOLDER_NAME))
    Logger.log("New folder \"" + folder.getName() + "\"");


    return folder;
  }

  /* high level functions */

  /*
   * Refresh existing folder data, using folder ids.
   *
   */
  refresh() {
    this.assert(this.dt.has(COL_FOLDER_ID), "missing column:'" + COL_FOLDER_ID + "'");

    let folders = this.dt.getDataAsMaps();
    Logger.log("Number of folder up for refresh: %s", folders.size);

    for (let [i, m] of folders) {
      let v = m.get(COL_FOLDER_ID);
      let id = v === undefined ? '' : v.toString().trim();
      if (id.length > 0) {
        Logger.log("Looking up for folder id: %s", id);

        let folder = DriveApp.getFolderById(id);
        Logger.log("Got folder: %s", folder);

        if (folder !== null) {
          this.updateRow(i, m, folder);
        }
      }
    }
  }

  /*
   * Update folder data, possibly adding new folder rows.
   *
   * TODO: based on a parent folder? a user?
   */
  update() {
  }

  /*
   * Create new folders from folder data rows without folder ids.
   *
   * @return {Array} - number of folders created, number of errors.
   */
  create() {
    this.dt.ensureColumnsExist([COL_FOLDER_ID, COL_FOLDER_STATUS]);

    /* various counters */
    var creation_count = 0;
    var error_count = 0;

    /* create the courses */
    let rows = this.dt.getDataAsMaps(); // FIXME: exclude read only fields?
    for (var [i, row] of rows) {

      this.dt.applyDefaultsToMap(row);
      LF.trimStringsInMap(row);

      /* do nothing if the id is set */
      if (!row.has(COL_FOLDER_ID) || row.get(COL_FOLDER_ID).length > 0) {
        continue;
      }

      /* do nothing if the name is empty */
      if (row.get(COL_FOLDER_NAME).length == 0) {
        continue;
      }

      try {
        let folder = this.createFolderFromMap(row);

        this.updateRow(i, row, folder);

        creation_count = creation_count + 1;

      } catch(e) {
        Logger.log("ERROR: %s", e.message);
        error_count = error_count + 1;
      }
    }
    return [creation_count, error_count];
  }

  /**
   * Append a sample row
   *
   */
  sample() {
    const m = new Map([
      [COL_FOLDER_NAME, "New folder"],
      [COL_FOLDER_OWNER, "me"],
      [COL_FOLDER_EDITORS, "jane.doe@teacher.abc.edu, john.doe@teacher.abc.edu"],
    ]);
    this.dt.addRowFromMap(m);
  }

  /**
   * Complete the table with all missing columns
   *
   */
  complete() {
    this.dt.ensureColumnsExist(this.cols);
  }
}
