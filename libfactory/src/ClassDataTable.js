/**
 * Class DataTable
 *
 * A DataTable is a rectangular group of cells, with a header row on top
 *
 * Christophe Bisière
 *
 * version 2022-04-17
 *
 * Note:
 *  - "var DataTable = class DataTable {...}"" is needed in ES6, as class
 *    declaration in the library is not accessible to library users:
 *    https://stackoverflow.com/questions/60456303/can-i-use-in-google-apps-scripts-a-defined-class-in-a-library-with-es6-v8
 */

/* separator used to generate a list as a string */
const DATATABLE_LIST_SEPARATOR = ' ';

/* default color to mark a cell's status */
const DATATABLE_COLOR_UPDATED = 'green';
const DATATABLE_COLOR_ERROR = 'red';

/* error in a cell of a data table */
var DataTableCellError = class DataTableCellError {
  constructor(label, message) {
    this.label = label;
    this.message = message;
    this.name = 'DataTableCellError';
  }
  toString() {
    return `${this.name}(${this.label}): ${this.message}`;
  }
}

/**
 * Class representing a data table, that is, a rectangulat area of cells,
 *  including the header on top and the data rows below.
 *
 * Note: assigning the class to a variable is required to export the name to
 *  to library users.
 */

var DataTable = class DataTable {
  /**
   * Create a DataTable.
   * @param {Range} rgCells - The range for the whole table.
   * @param {Map} defaults - A map of default values.
   * @param {Map} formats - A map of function to format cells.
   */
  constructor(rgCells, defaults=null, formats=null) {
    this.assert(rgCells != undefined, 'table range not specified');
    this.r = rgCells;
    this.defaults = defaults;
    this.formats = formats;
    /* cached map */
    this.map = null;
  }

  /**
   * Returns a Range containing a rectanglular data table, including the header
   *  row, using one header cell a as "seed" from which the full table is found
   *  by expansion
   *
   * The seed cell (or cells) is first expanded using getDataRegion() and then
   *  rows above the seed are removed. To simulate getDataRegion(), select a
   *  cell within the region and type Ctrl-a (or command-a on macOS). This
   *  implies that tables within the same sheet has to be separated by blank
   *  rows and columns.
   *
   * @see {@link https://developers.google.com/apps-script/reference/spreadsheet/range#getdataregion}
   *
   * @param {Range} rgTopCell - The "seed" cell (or cells) from which the Range
   *                            is built.
   * @return {Range} The Range of the data table found.
   */
  static locateFromHeaderCell(rgTopCell) {
    const rgRegion = rgTopCell.getDataRegion();
    const numDataRows = rgRegion.getLastRow() - rgTopCell.getRow() + 1;
    const rgResult = rgTopCell.getSheet().getRange(rgTopCell.getRow(),
        rgRegion.getColumn(), numDataRows, rgRegion.getNumColumns());

    return rgResult;
  }

  /**
   * Locate the target data table table in a Sheet, based on a column label
   *
   * When the sheet contains more than one table, select the table intersecting
   *  with the current cell, or, otherwise, the first table.
   *
   * @param {Sheet} sheet - The Sheet containing the table to look for.
   * @param {string} label - The column label to look for.
   * @return {?Range} The Range of the table found.
   */
  static locateFromLabel(sheet, label) {
    const arrCells = SheetHelper.findCellsByContent(sheet, label);

    /* collect all candidate regions */
    const rs1 = [];
    for (const r of arrCells) {
      rs1.push(DataTable.locateFromHeaderCell(r));
    }

    /* sort by row-column */
    rs1.sort(function(r1, r2) {
      return r1.getRow() - r2.getRow() +
          (r1.getRow() == r2.getRow() ? r1.getColumn() - r2.getColumn() : 0);
    });

    /* keep regions that do not overlap */
    const rs2 = [];
    while (rs1.length) {
      const q = rs1.pop();
      if (!rs1.some(function(r) {
        return SheetHelper.rangeIntersect(r, q);
      })) {
        rs2.unshift(q);
      }
    }

    /* look for a region that intersects with the current cell */
    const c = SpreadsheetApp.getActiveSheet().getSelection().getCurrentCell();
    const ri = rs2.find(function(r) {
      return SheetHelper.rangeIntersect(r, c);
    });

    /* return that region, or the first in the list otherwwise */
    return rs2.length == 0 ? null : ri != undefined ? ri : rs2[0];
  }

  /* 
   * static method to handle lists of items
   * 
   * A data table cell may accept a list of items as value, e.g. email addresses. 
   * When pasing a list, accepted separators are one or more spaces, commas, 
   * and semicolons. When writing, a single blank space is used. 
   */
  
  /**
   * Return a value in a given map as a list of items.
   *
   * @param {string} label - The column label.
   * @param {Map} m - The label-to-value map.
   * @return {Array} The list of values.
   */
  static getValueAsList(label, m) {
    const v = getValue(label, m);
    return v !== false ? itemsInString(v, "[\\s,;]+") : [];
  }


  /**
   * Check a condition.
   *
   * @param {boolean} condition - Condition to check.
   * @param {string} message - Message to display if the condition is not met.
   */
  assert(condition, message) {
    const a1 = this.getRange() == undefined ? 'undefined' : this.getRange().getA1Notation();
    const prompt = 'Error: DataTable (' + a1 + '):' + (message || ' assertion failed');
    assert(condition, prompt);
  }



  /* getters / setters: range */

  /**
   * Set the Range of cells.
   *
   * @param {Range} r - The Range.
   */
  setRange(r) {
    this.r = r;
  }

  /**
   * Return the table (including header) as a Range object.
   *
   * @return {Range|undefined} The whole data table as a Range object.
   */
  getRange() {
    return this.r;
  }

  /* getters: dimensions, has label... */

  /**
   * Return the number of data rows.
   *
   * @return {number} The number of data rows in the data table.
   */
  getNumRows() {
    return this.getRange().getNumRows() - 1;
  }

  /**
   * Return true if the table has no data rows.
   *
   * @return {boolean} True if the table has no rows, False otherwise.
   */
  isEmpty() {
    return this.getNumRows() === 0;
  }

  /**
   * Return the number of columns.
   *
   * @return {number} The number of columns in the data table.
   */
  getNumColumns() {
    return this.getRange().getNumColumns();
  }

  /**
   * True whether a given column label exists.
   *
   * @param {string} label - The column label.
   * @return {boolean} True if the column exist.
   */
  has(label) {
    const m = this.getMap();
    return m.has(label);
  }

  /* getters: Range */

  /**
   * Return the header row as a Range object.
   *
   * @return {Range} The header of the data table.
   */
  getHeader() {
    return this.getRange().offset(0, 0, 1);
  }

  /**
   * Return consecutive rows as a Range object.
   *
   * @param {number} i - The first row number.
   * @param {number} nb - The number of rows.
   * @return {Range} The rows as a Range.
   */
  getRows(i, nb) {
    const n = this.getNumRows();
    this.assert((1 <= i) && (i + nb - 1 <= n),
        'row range (' + i + ',' + (i + nb - 1) + ') outside table range');
    const r = this.getHeader().offset(i, 0, nb);
    return r;
  }

  /**
   * Return a row as a Range object.
   *
   * @param {number} i - The row number.
   * @return {Range} The row as a Range.
   */
  getRow(i) {
    return this.getRows(i, 1);
  }

  /**
   * Return the last row as a Range object.
   *
   * @return {Range} The row as a Range.
   */
  getLastRow() {
    return this.getRow(this.getNumRows());
  }

  /**
   * Return consecutive columns (including header) as a Range object.
   *
   * @param {number} j - The first column number.
   * @param {number} nb - The number of columns.
   * @return {Range} The columns as a Range.
   */
  getColumns(j, nb) {
    const n = this.getNumColumns();
    this.assert((1 <= j) && (j + nb - 1 <= n),
        'column range (' + j + ',' + (j + nb - 1) + ') outside table range');
    const r = this.getRange().offset(0, j-1, this.getRange().getNumRows(), nb);
    return r;
  }

  /**
   * Return a column (including header) as a Range object.
   *
   * @param {number} j - The column number.
   * @return {Range} The column as a Range.
   */
  getColumn(j) {
    return this.getColumns(j, 1);
  }

  /**
   * Return the last column (including header) as a Range object.
   *
   * @return {Range} The last column as a Range.
   */
  getLastColumn() {
    return this.getColumn(this.getNumColumns());
  }

  /**
   * Return the data part as a Range object, or null if the data table
   *  is empty.
   * @return {Range|null} The data part of the data table.
   */
  getData() {
    const n = this.getNumRows()
    return n == 0 ? null : this.getRows(1, n);
  }

  /* getter/setter: individual cell */

  /**
   * Get a cell.
   *
   * @param {number} i - The row number of the cell.
   * @param {string} label - The label of the column of the cell.
   * @return {Range|null} The cell.
   */
  getCell(i, label) {
    this.assert(this.has(label), 'label "' + label + '" is not in the table');
    const map = this.getMap();
    const c = this.getRow(i).getCell(1, map.get(label));
    return c;
  }

  /**
   * Set a cell value.
   *
   * @param {number} i - The row number.
   * @param {string} label - The label of the column of the cell.
   * @param {Object} value - The value to set.
   */
  setValue(i, label, value) {
    this.getCell(i, label).setValue(value);
  }

  /**
   * Set a cell value as a RichTextValue.
   *
   * @param {number} i - The row number.
   * @param {string} label - The label of the column of the cell.
   * @param {RichTextValue} value - The RichTextValue to set.
   */
  setRichTextValue(i, label, value) {
    this.getCell(i, label).setRichTextValue(value);
  }

  /* getters: Array */

  /**
   * Return the header row as a single-dimensional array.
   * @return {Array} The header of the data table.
   */
  getHeaderAsVector() {
    return this.getHeader().getValues()[0];
  }

  /**
   * Return the data rows as a two dimensional array, or an empty array
   *  if the data table is empty
   * @return {Array|[]} The data part of the data table.
   */
  getDataAsArray() {
    const r = this.getData();
    return r == null ? [] : r.getValues();
  }

  /**
   * Return rows as a two dimensional array.
   *
   * @param {number} i - The first row number.
   * @param {number} nb - The number of rows.
   * @return {Array} The rows as a two dimensional array.
   */
  getRowsAsArray(i, nb) {
    // TODO: assert nb>0
    return this.getRows(i, nb).getValues();
  }

  /**
   * Return one row as a single dimensional array.
   *
   * @param {number} i - The row number.
   * @return {Array} The rows as a single dimensional array.
   */
  getRowAsVector(i) {
    return this.getRows(i, 1).getValues()[0];
  }

  /* setters: Array */

  /**
   * Set rows data using a two dimensional array.
   *
   * @param {number} i - The first row number.
   * @param {number} nb - The number of rows.
   * @param {Array} data - The array of data.
   */
  setRowsFromArray(i, nb, data) {
    const r = this.getRows(i, nb);
    this.assert(data.length == nb,
        'wrong number of rows in data:' + data.length);
    r.setValues(data);
  }

  /**
   * Set row data using a one dimensional array.
   *
   * @param {number} i - The row number.
   * @param {Array} vec - The row data.
   */
  setRowFromVector(i, vec) {
    this.setRowsFromArray(i, 1, [vec]);
  }

  /**
   * Set table data using a two dimensional array.
   * @param {Array} data - The array of data.
   */
  setDataFromArray(data) {
    this.setRowsFromArray(1, this.getNumRows(), data);
  }


  /* getters: Map */

  /**
   * Return a Map of column label to column index
   *
   * Use lazy evaluation.
   *
   * @return {Map} The map of label to index.
   */
  getMap() {
    if (this.map == null) {
      this.map = new Map();
      for (const [j, label] of this.getHeaderAsVector().entries()) {
        this.map.set(label, j+1);
      }
    }
    return this.map;
  }

  /**
   * Return an empty Map of column label to undefined values.
   *
   * @return {Map} The map of column label to undefined values.
   */
  getEmptyRowAsMap() {
    // TODO: store in cache?
    const map = this.getMap();
    const m = new Map();
    for (const [label, j] of map) {
      m.set(label, undefined);
    }
    return m;
  }

  /**
   * Return a Map of column label to cell value for a given row.
   *
   * @param {number} i - The row number.
   * @return {Map} The map of column label to value.
   */
  getRowAsMap(i) {
    const vec = this.getRowAsVector(i);
    const map = this.getMap();
    const res = new Map();
    for (const [label, j] of map) {
      res.set(label, vec[j-1]);
    }
    return res;
  }

  /**
   * Return a Map of maps, each being a map column label to cell value
   *  for a given row.
   *
   * @param {number} i - The first row number.
   * @param {number} nb - The number of rows.
   * @return {Map} The map of maps.
   */
  getRowsAsMaps(i, nb) {
    const m = new Map();
    for (let k=i; k<=i+nb-1; k++) {
      m.set(k, this.getRowAsMap(k));
    }
    return m;
  }

  /**
   * Return a Map of maps, each being a map column label to cell value
   *  for a given row.
   *
   * @return {Map} The map of maps.
   */
  getDataAsMaps() {
    return this.getRowsAsMaps(1, this.getNumRows());
  }

  /* setters: Map */

  /**
   * Set a data row from a Map of column label to cell value.
   *
   * @param {number} i - The row number.
   * @param {Map} m - The label-to-value map.
   */
  setRowFromMap(i, m) {
    for (const [k, v] of m) {
      if (this.has(k)) {
        const c = this.getCell(i, k);
        c.setValue(v); 
      }
    }
  }


  /**
   * Update a data row from a Map of column label to cell value.
   * 
   * The map only contains values that may be updated.
   *
   * @param {number} i - The row number.
   * @param {Map} m - The label-to-value map.
   * @param {string?} updateColor - The CSS font color to set when a 
   *                           value did change, or null.
   */
  updateRowDataFromMap(i, m, updateColor=null) {
    for (const [k, v] of m) {
      if (this.has(k)) {
        const c = this.getCell(i, k);
        if (updateColor == null) {
          c.setValue(v); 
        } else if (c.getValue() !== v) {
          c.setValue(v).setFontColor(updateColor); 
        }
      }
    }
  }

  /* set default values and formats */

  /**
   * Complete a map with defaults values when specified.
   *
   * This may create new labels in the map.
   *
   * @param {Map} row - The row to which apply defaults.
   */
  applyDefaultsToMap(row) {
    if (this.defaults != null) {
      for (const [label, defaultValue] of this.defaults) {
        if (!row.has(label) || row.get(label) == undefined ||
            row.get(label).toString().trim() === '') {
          row.set(label, defaultValue);
        }
      }
    }
  }


  /**
   * Format a cell.
   *
   * TODO: Only automatic format are changed.
   *
   * @param {Range} c - The cell.
   * @param {string} label - The column label.
   */
  setCellFormat(c, label) {
    if (this.formats != null && this.formats.has(label)) {
      //const format = c.getNumberFormat();
      //Logger.log("format of %s is %s", c.getA1Notation(), format);
      this.formats.get(label)(c);
    }
  }

  /**
   * Format a row.
   *
   * @param {number} i - The row number.
   */
  setRowFormat(i) {
    if (this.formats != null) {
      const map = this.getMap();
      const r = this.getRow(i);
      for (const [label, j] of map) {
        const c = r.offset(0, j-1, 1, 1);
        this.setCellFormat(c, label);
      }
    }
  }

  /**
   * Format the data part of a column.
   *
   * @param {string} label - The column label.
   */
  setColumnFormat(label) {
    assert(this.has(label), 'setColumnFormat: unknown label "' + label + '"');
    if (!this.isEmpty() && this.formats != null && this.formats.has(label)) {
      const map = this.getMap();
      const j = map.get(label);
      const rCol = this.getColumn(j);
      const rData = rCol.offset(1, 0, rCol.getNumRows()-1);
      rData.clearFormat(); /* cheating? :-) */
      this.formats.get(label)(rData);
      Logger.log("format set for: %s", rData.getA1Notation());
    }
  }


  /* High-level functions (that do handle formatting). */

  /**
   * Reset the color of a row.
   *
   * @param {number} i - The row number.
   */
  resetRowColor(i) {
    this.getRow(i).setFontColor(null);
  }

  /**
   * Set the color of a cell to the error color.
   * 
   * The map only contains values that may be updated. The
   * function rest the color and the format of the whole row.
   *
   * @param {number} i - The row number.
   * @param {string} label - The label of the column.
   */
  setErrorColor(i, label) {
    if (this.has(label)) {
        const c = this.getCell(i, label);
        c.setFontColor(DATATABLE_COLOR_ERROR);
    }
  }

  /**
   * Update a row from a Map of column label to cell value.
   * 
   * The map only contains values that may be updated. The
   * function rest the color and the format of the whole row.
   *
   * @param {number} i - The row number.
   * @param {Map} m - The label-to-value map.
   */
  updateRow(i, m) {
    this.setRowFormat(i);
    this.resetRowColor(i); //FIXME: should not apply to user data
    this.updateRowDataFromMap(i, m, DATATABLE_COLOR_UPDATED); /* FIXME: buggy */
  }


  /**
   * Add a new, empty row.
   *
   * @return {number} The new row index.
   */
  addRow() {
    const lastInTable = this.getRange().getLastRow();
    const lastInSheet = this.getRange().getSheet().getMaxRows();

    if (lastInTable === lastInSheet) {
      this.getRange().getSheet().insertRowsAfter(lastInSheet, 1);
    } else {
      const rLast = this.isEmpty() ? this.getHeader() : this.getLastRow();
      rLast.offset(1,0).insertCells(SpreadsheetApp.Dimension.ROWS);
    }
    this.setRange(this.getRange().offset(0, 0, this.getRange().getNumRows()+1));

    const index = this.getNumRows();
    this.setRowFormat(index);

    return index;
  }

  /**
   * Append a new rows from a map.
   *
   * @param {Map} m - The map.
   */
  addRowFromMap(m) {
    // TODO: applies defaults here?
    const i = this.addRow();
    this.updateRowDataFromMap(i, m);
  }

  /**
   * Appends new rows from a set of maps.
   *
   * @param {Set} rows - The set of maps.
   */
  addRowsFromSet(rows) {
    for (const row of rows) {
      this.addRowFromMap(row);
    }
  }

  /**
   * Insert new columns in the table.
   *
   * @param {number} j - The column number before which the new column must be
   *                      inserted.
   * @param {string[]} labels - The new column labels.
   */
  insertColumnsBefore(j, labels) {
    const n = this.getNumColumns();
    this.assert((1 <= j) && (j <= n+1),
        'insertColumnsBefore: invalid column index ' + j + ' when inserting');

    const m = labels.length; /* number of columns to insert */
    this.assert(m > 0, 'insertColumnsBefore: no columns specified');

    let r = null; /* first column of the inserted columns */

    if (j == n+1) {
      const lastInTable = this.getRange().getLastColumn();
      const lastInSheet = this.getRange().getSheet().getMaxColumns();
      const available = lastInSheet - lastInTable;

      if (available < m) {
        /* no room for append: insert new columns */
        this.getRange().getSheet().insertColumnsAfter(lastInSheet, m-available);
        Logger.log("%s new sheet columns inserted", m-available);
      }
      r = this.getLastColumn().offset(0, 1);
    } else {
      /* insert m columns before columns j */
      r = this.getColumn(j);
      for (let k=0; k<m; k++) {
        r.insertCells(SpreadsheetApp.Dimension.COLUMNS); // FIXME: could disturb other table data in the sheet
      }
    }
    Logger.log('inserted %s new column at %s', m, r.getA1Notation());

    /* update instance variables */
    this.setRange(this.getRange().offset(0, 0, this.getRange().getNumRows(),
        this.getRange().getNumColumns() + m));
    this.map = null;

    /* set header labels */
    r.getCell(1, 1).offset(0, 0, 1, m).setValues([labels]);

    /* format the new cells if any */
    for (const label of labels) {
      this.setColumnFormat(label);
    }
  }

  /**
   * Append new columns to the table.
   *
   * The list may be empty, before or after filtering.
   *
   * @param {string[]} labels - The new column labels.
   * @param {boolean=} nodup - Only append columns with new labels.
   */
  addColumns(labels, nodup=false) {
    const arr = [];
    for (const label of labels) {
      if (!nodup || !this.has(label)) {
        arr.push(label);
      }
    }
    if (arr.length > 0) {
      this.insertColumnsBefore(this.getNumColumns() + 1, arr);
    }
  }

  /**
   * Append a new column to the table.
   *
   * @param {string} label - The new column label.
   * @param {boolean=} nodup - Only append columns with new labels.
   */
  addColumn(label, nodup=false) {
    this.addColumns([label], nodup);
  }

  /**
   * Ensure some columns exist in the table.
   *
   * @param {string[]} labels - The column labels.
   */
  ensureColumnsExist(labels) {
    this.addColumns(labels, true);
  }

  /**
   * Ensure a column exists in the table.
   *
   * @param {string} label - The column label.
   */
  ensureColumnExists(label) {
    this.ensureColumnsExist([label]);
  }

  /**
   * Checks several columns exist in the table, raising an exception on the first column that does not exist.
   *
   * @param {string[]} labels - The column labels.
   * @param {string} prompt - The prompt for "Missing column:".
   */
  checkColumnsExist(labels, prompt) {
    for (const label of labels) {
      if (!this.has(label)) {
        throw new Error(prompt + ' \'' + label  + '\'.');        
      }
    }
  }

  /**
   * Check a column exists in the table, raising an exception if it does not.
   *
   * @param {string} label - The column label.
   */
  checkColumnExists(label, prompt) {
    this.checkColumnsExist([label], prompt);
  }

  /**
   * Checks the table has at least one row, raising an exception if the table is empty.
   *
   * @param {string} message - The error message attached to the exception.
   */
  checkNonEmpty(message) {
    if (this.isEmpty()) {
      throw new Error(message);        
    }
  }
};
