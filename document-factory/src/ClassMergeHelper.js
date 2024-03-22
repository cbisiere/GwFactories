/**
 * Library of DocMerge-specific functions
 *
 * Christophe Bisière
 *
 * version 2019-04-17
 *
 */

class MergeHelper {

  /**
   * Find all tags of the form: "<text1<tag?>text2>" in a document part,
   *  and return an array of
   *   [
   *     "<text1<tag>text2>",
   *     RangeElement range,
   *     "text1",
   *     "tag",
   *     "text2",
   *     opt?, (true if the tag name is followed by a "?")
   *     Object part (could be Body, etc.)
   *   ]
   *
   * TODO: find a way to allow "<" or ">" in text1 and text2
   *
   */
  static findAllTags(docPart) {
    const pattern = '<[^<>]*<[^<>\\?]+\\??>[^<>]*>';

    const res = [];

    const ranges = findAllRanges(docPart, pattern);

    /* decompose each search result */
    for (let i = 0; i < ranges.length; i++) {
      const item = [];

      const range = ranges[i];
      let match = getStringFromRange(range); /* "<text1<tag>text2>" */

      item.push(match);
      item.push(range);

      match = match.substring(1, match.length - 1); /* "text1<tag>text2" */

      const a = match.split(/[<>]/); /* ["text1", "tag", "text2"] */

      /* extract optional trailing "?" to the tag */
      let opt = false;
      let tag = a[1];
      if (tag.substring(tag.length - 1) == '?') {
        opt = true;
        tag = tag.substring(0, tag.length - 1);
      }

      item.push(a[0]);
      item.push(tag);
      item.push(a[2]);
      item.push(opt);
      item.push(docPart);

      res.push(item);
    }

    return res;
  }

  /**
   * Modify a document by replacing tags with values given in a map.
   *
   * @param {Document} document - The opened document to modify.
   * @param {Map} fmap - The label-to-value map to use to replace tags.
   */

  static merge(document, map) {
    /* look for tag constructs in all parts of the document */
    const parts = getDocumentParts(document);
    const tags = [];

    for (var p = 0; p < parts.length; p++) {

      var arr = MergeHelper.findAllTags(parts[p]);
      for (var t = 0; t < arr.length; t++) {
        tags.push(arr[t]);
      }
    }

    /* Massive search-replace, starting from the end */
    for (var t = tags.length - 1; t >= 0; t--) {

      Logger.log(`Processing tag "${tags[t]}"`);
      /* true if the tag had a trailing "?" */
      var opt = tags[t][5];

      /* column label for the tag, or null if no column exists for this tag */
      var tagName = tags[t][3];
      var tagIsKnown =  map.has(tagName);

      /* value to substitute to the tag */
      let tagValue;
      if (!tagIsKnown) {
        /* no column exists with this tag */
        if (opt) {
          /* in optional mode, we simply substitute a blank string */
          tagValue = '';
        } else {
          /* else we skip the replacement, and the ugly pattern stays as-is in the output document */
          continue;
        }
      } else {
        /* a column exists: get the replacement from the worksheet */
        tagValue = map.get(tagName);
      }

      /* replacement target */
      var range = tags[t][1];
      var element = range.getElement();
      var text1 = tags[t][2];
      var text2 = tags[t][4];

      /* indexes of the substring to replace */
      var first = range.getStartOffset();
      var last = range.getEndOffsetInclusive();
      
      /* does the replacement would fill all the element containing the tag (paragraph, item...) */
      var isFullRep = (tags[t][0] == element.getText()) && (text1 == "") && (text2 == "");
      /* parent of the target text element */
      var parent = element.getParent();
      /* is this container a list item? */
      var isItem = parent.getType() == DocumentApp.ElementType.LIST_ITEM;
      /* lines in the replacement text */
      var lines = tagValue.split("\n");

      /* are we in the special case where extra list items must be created? */
      var isNewItemsCase = (isItem && isFullRep && (lines.length > 0));

      Logger.log(`IsItem ${isItem}`);
      Logger.log(`Full replacement of the container ${isFullRep}`);
      Logger.log(`Replace index is ${first} to ${last}`);
      Logger.log(`Number of lines is ${lines.length}`);
      LF.logObject(lines,"lines");

      /* special case: full replacement of a list item with multiple lines: create one new item per extra line */
      if (isNewItemsCase) {
        tagValue = lines.shift(); /* will replace the existing item with the first line in the tag value  */
      }

      /* replacement text */
      var replaceTxt = "";
      /* conditional replacement */
      if (tagValue.length > 0) { 
        replaceTxt = text1 + tagValue + text2;
      }
    
      Logger.log(`Replacement is ${replaceTxt}`);
      Logger.log(`Length of replacement is ${replaceTxt.length}`);

      /* first replacement */
      if (replaceTxt.length > 0) {

        var op_size = 1;
        var tag_pos = first + op_size + text1.length + 1;
        var tag_len = tagName.length

        element.deleteText(last - op_size + 1, last);                 /* delete final ">" or ">>" */
        element.deleteText(tag_pos, last - op_size - text2.length);   /* delete "tag>" */
        element.insertText(tag_pos, tagValue);                        /* insert tag value after second "<" */
        element.deleteText(tag_pos - op_size, tag_pos - 1);           /* delete second "<" */
        element.deleteText(first, first + op_size - 1);               /* delete first "<" or "<<" */

      } else {
        element.deleteText(first, last);
      }

      /* create the new items with the remaining lines */
      if (isNewItemsCase) {
        var part = tags[t][6];
        var index = part.getChildIndex(parent);
        var glyph = parent.getGlyphType(); /* bullet, etc. */
        for (var i in lines) {
          index += 1;
          part.insertListItem(index, lines[i]).setGlyphType(glyph);
        }
      }

      /* if the replacement ends up creating an empty paragraph, item or table row, it is deleted */
      var pg = element.getParent()
      Logger.log(`Parent type is ${pg.getType()}`);
      Logger.log(`Parent content is "${pg.asText().getText()}"`);
      if (pg.asText().getText().length == 0) {
        /* empty paragraph */
        Logger.log(`Parent paragraph is empty`);
        /* backup the parent, which might be a TableCell */
        var cell = pg.getParent()
        Logger.log(`Removing empty paragraph`);
        pg.removeFromParent();

        /* detect empty table row */
        if (cell != null && cell.getType() == DocumentApp.ElementType.TABLE_CELL) {
          var r = cell.getParent()
          if (r != null && r.getType() == DocumentApp.ElementType.TABLE_ROW) {

            /* are the cells all empty in this row? */
            var allEmpty = true
            for (var j = 0; j < r.getNumCells(); j++) {
              var label = r.getCell(j)
              if (label != null && label.getType() == DocumentApp.ElementType.TABLE_CELL) {
                var isEmpty = label.asText().getText().length == 0
                allEmpty = allEmpty && isEmpty
                Logger.log('cell ' + j + ' out of ' + r.getNumCells() + ': ' + isEmpty);
                if (!allEmpty)
                  break;
              }
            }
            if (allEmpty) {
              var table = r.getParentTable()
              var row_index = table.getChildIndex(r)
              Logger.log('Removing row ' + row_index);
              table.removeRow(row_index)
            }
          }
        }
      }
    }
  }
}