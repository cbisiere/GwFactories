/**
 * Library of general purpose functions
 *
 * Christophe Bisière
 *
 * version 2015-10-04
 * updated 2020-07-01
 *   - add itemsInString, quotedItemsInString
 * updated 2022-07-23
 *   - add inc
 *
 */


/**
 * Return a map of changes between two maps.
 *
 * Retains items from the second map that are new with respect to the first map, 
 * that is, do not exist in the first map or have a different value. 
 * 
 * @param {Map} m1 - The original map.
 * @param {Map} m2 - The changed map.
 * @return {Map} The map of changes between the two map.
 */
function mapDiff(m1, m2) {
  const m = new Map(m2);
  for ([k, v] of m2) {
    if ((m1.has(k)) && (m1.get(k) === v)) {
      m.delete(k);
    }
  }
  return m;
}

/**
 * Return the a non-empty value in a given map of strings or false if it is 
 * not possible, that is, if the key does not exist or the string value has
 * a length of zero. 
 *
 * The function does not trim the value before checking its length.
 * 
 * @param {string} key - The key to look for.
 * @param {Map} map - The key-to-value map.
 * @return {string|false} The non-empty value of false.
 */
function getValue(key, map) {
  if (!map.has(key)) {
    return false;
  }
  const v = map.get(key);
  if (v.length == 0) {
    return false;
  }
  return v;
}

/**
 * Increase a counter stored in a map
 *
 * The map item is created if it does not exist 
 *
 * @param {Map} row - The row containing the counter.
 * @param {string} key - The key of the counter in the map.
 * @param {integer} delta - The value to add to the counter.
 */
function inc(map, key, delta = 1) {
  if (!map.has(key)) {
    map.set(key, 0);
  }
  map.set(key, map.get(key) + delta);
}

/**
 * Return a map in which all values have been trimmed.
 *
 * Undefined values are replaced by ''.
 *
 * @param {Map} row - The row to which apply defaults.
 */
function trimStringsInMap(map) {
  for (const [k, v] of map) {
    const s = v == undefined ? '' : v.toString().trim();
    map.set(k, s);
  }
}


/**
 * Return the key and map of the first map in a map of maps having a specific
 *  value for
 * a specific key, or null.
 *
 * Note: use '==', not '==='.
 */
function findMap(maps, key, value) {
  for (const [k, m] of maps) {
    if (m.get(key) == value) {
      return [k, m];
    }
  }
  return null;
}


/**
 * Log a map.
 *
 */
function logMap(m, s) {
  for (const [k, v] of m) {
    Logger.log("%s[%s]=%s", s, k, v);
  }
}

/**
 * Log an object map.
 *
 */
function logObject(o, s) {
  for (const p in o) {
    Logger.log("%s[%s]=%s", s, p, o[p]);
  }
}

/**
 * Log a list of ranges.
 *
 */
function logRanges(ar, s) {
  for (const i in ar) {
    Logger.log("%s[%s]=%s", s, i, ar[i].getA1Notation());
  }
}

/**
 * Raise an exception if a condition is not met.
 *
 * https://stackoverflow.com/questions/15313418/what-is-assert-in-javascript
 */
function assert(condition, message) {
  if (!condition) {
    message = message || "Assertion failed";
    if (typeof Error !== "undefined") {
      throw new Error(message);
    }
    throw message; // Fallback
  }
}

/**
 * Escape all regexp special characters in a string
 *
 */
function stringToRegex(s) {
  return s.replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&");
}

/**
 * return an array of non-empty items in a string using a regexp as separator,
 * or null if no items can be found
 */
function itemsInString(s, regexp) {
  const re = new RegExp(regexp);
  const a = s.split(re);
  return a.filter(function(e) {
    return e != '';
  });
}

/*
 * return an array of non-empty double-quoted items in a string,
 * or null if no items can be found
 */
function quotedItemsInString(s, quote) {
  const re = new RegExp('"');
  const a = s.split(re);
  return a.filter(function(e, i) {
    return i % 2 && e != '';
  });
}

/*
 * return a string trimmed from a set of chars
 */
function trimChars(s, chars) {
  const re = new RegExp('^[' + chars + ']*|[' + chars + ']*$', 'g');
  return s.replace(re, '');
}

/**
 * Select a translation in an array of translated strings (0: defaut, 1: fr)
 *
 */
function i18n(messages) {
  const lang = Session.getActiveUserLocale();
  if (lang == 'fr') {
    return messages[1];
  }
  return messages[0];
}
