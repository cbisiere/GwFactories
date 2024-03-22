/**
 * Class SheetHelper
 *
 * A class for Sheet helpers
 *
 * Christophe Bisière
 *
 * version 2021-01-01
 *
 */

var SheetHelper = class SheetHelper {

 /**
   * Display a message if the UI is available.
   *
   * @param {String} message The message to display.
   */
  static alertUser(message) {
    Logger.log(message);
    try {
      SpreadsheetApp.getUi().alert(message);
    } catch(e) {
    }
  }

 /**
   * Format each cell in a range as text.
   *
   * @param {Range} r - The range to format as text.
   */
  static setRangeFormatAsText(r) {
    r.setNumberFormat('@');
  }

 /**
   * Format each cell in a range as datetime.
   *
   * @see {@link https://developers.google.com/sheets/api/guides/formats}
   *
   * @param {Range} r The range to format as datetime.
   * @param {String} format The date-time format pattern.
   */
  static setRangeFormatAsDatetime(r, format='yyyy-MM-dd HH:mm:ss') {
    Logger.log("Formating \"%s\" as \"%s\"", r.getA1Notation(), format);
    r.setNumberFormat(format); //TODO: i18n
  }

  /**
   * Format each cell in a range as a list of values.
   *
   * @param {Range} r The range to format as a list.
   * @param {Array} values The list of allowed values.
   * @param {boolean} allowInvalid True if input that fails data validation is allowed.
   */
  static setRangeFormatAsList(r, values, allowInvalid = false) {
    const validation = SpreadsheetApp.newDataValidation()
      .setAllowInvalid(allowInvalid)
      .requireValueInList(values, true)
      .build();
    r.setDataValidation(validation);
  }

 /**
   * Return a RichTextValue containing a hyperlink.
   *
   * @param {String} url The url of the hyperlink.
   * @param {String} label The label of the hyperlink.
   * @return {RichTextValue} The RichTextValue containing the link.
   */
  static makeLink(url, label) {
    return SpreadsheetApp.newRichTextValue()
      .setText(label)
      .setLinkUrl(url)
      .build();
  }

 /**
  * Return true if two ranges intersect
  *
  * https://stackoverflow.com/questions/36358955/google-app-script-check-if-range1-intersects-range2
  */
  static rangeIntersect(r1, r2) {
    return (r1.getLastRow() >= r2.getRow()) && (r2.getLastRow() >= r1.getRow()) && (r1.getLastColumn() >= r2.getColumn()) && (r2.getLastColumn() >= r1.getColumn());
  }

 /**
  * Returns a range which is the first n row of r
  *
  * Does *not* check if there is enough rows in r. In that case, the range will
  * be larger than r, expanding downward.
  */
  static getFirstRows(r, n) {
    return r.getSheet().getRange(r.getRow(), r.getColumn(), n, r.getNumColumns());
  }

 /**
  * Returns a range which is the last n row of r.
  *
  * Does *not* check if there is enough rows in r. In that case, the range will
  * be larger than r, expanding upward.
  */
  static getLastRows(r, n) {
    return r.getSheet().getRange(r.getLastRow()-n, r.getColumn(), r.getLastRow(), r.getNumColumns());
  }

 /**
  * Returns a range which is made of n rows just below r.
  *
  * Does *not* check if there is enough rows below r.
  */
  static getRowsBelow(r, n) {
    return r.getSheet().getRange(r.getRow()+1, r.getColumn(), n, r.getNumColumns());
  }

 /**
  * Returns a range which is made of column number n in r.
  *
  * Does *not* check if r has enough columns.
  */
  static getColumn(r, n) {
    return r.getSheet().getRange(r.getRow(), r.getColumn()+n-1, r.getNumRows(), 1);
  }

  /**
   * Returns an array of cells (each being a Range) in a sheet, matching an entire
   *  string
   *
   * Note: is case insensitive by default (add matchCase(true) to change that)
   *
   * https://developers.google.com/apps-script/reference/spreadsheet/text-finder
   */
  static findCellsByContent(sheet, findText) {
    var finder = sheet.createTextFinder(findText).matchEntireCell(true)
    return finder.findAll();
  }
}
