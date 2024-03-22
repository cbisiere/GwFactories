/**
 * Library of functions handling text Documents
 *
 * Christophe Bisière
 *
 * version 2015-10-04
 *
 */


/**
 * Return an array of RangeElement objects matching a pattern in a document
 * part (i.e. header, body, footer, or footnote)
 *
 */
function findAllRanges(docPart, pattern) {
  const res = [];

  let range = docPart.findText(pattern);
  while (range != null) {
    res.push(range);
    range = docPart.findText(pattern, range);
  }

  return res;
}

/**
 * Return the string pointed by a RangeElement
 *
 */
function getStringFromRange(range) {
  let str = range.getElement().asText().getText();

  if (range.isPartial()) {
    str = str.substring(range.getStartOffset(),
      range.getEndOffsetInclusive() + 1);
  }

  return str;
}

/**
 * Return an array of (possibly unique) strings matching a pattern in a
 *  document part
 *
 */
function findAll(docPart, pattern, unique) {
  const res = [];
  const ranges = findAllRanges(docPart, pattern);

  for (let i = 0; i < ranges.length; i++) {
    const str = getStringFromRange(ranges[i]);
    if (!unique || res.indexOf(str) == -1) {
      res.push(str);
    }
  }
  return res;
}

/**
 * Return an array of document parts
 */

function getDocumentParts(document) {
  const parts = [];

  if (document.getHeader() != null) {
    parts.push(document.getHeader());
  }
  if (document.getBody() != null) {
    parts.push(document.getBody());
  }
  if (document.getFooter() != null) {
    parts.push(document.getFooter());
  }
  const footnotes = document.getFootnotes();
  for (let f = 0; f < footnotes.length; f++) {
    parts.push(footnotes[f].getFootnoteContents());
  }

  return parts;
}
