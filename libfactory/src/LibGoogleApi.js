/**
 * Google API helpers
 *
 * Christophe Bisière
 *
 * version 2020-12-13
 *
 */

/**
 * Retrieves a map of objects using an object list function of the Google API.
 *
 * Objects are retrieved using a function func. If arg is not undefined, arg is
 *  a mandatory first argument of func. Optional arguments are passed in
 *  optionalArgs.
 * The function returns a page of results. Each page has a field with name
 *  listName, the list of results on the page. The map is indexed by object
 *  property indexName.
 *
 * @param {*|undefined} arg First argument, if any, of the list function
 * @param {*|undefined} optionalArgs Optional argument, if any, of the list
 *                       function
 * @param {(function({pageSize:integer,pageToken: string})
 *  |function(*,{pageSize:integer,pageToken: string})} func
 *   The Google API list function
 * @param {string} listName The key to look for in Google response to get the
 *  returned list
 * @param {string} indexName The object property to use as index in the map
 * @param {number=} maxObjects The maximum number of objects to return
 * @param {number=} pageSize The maximum number of objects to retrieve in a call
 * @return {!Map<*,Object>} A map of requested objects indexed by property
 *  indexName.
 */
function getGoogleList(arg, optionalArgs, func, listName, indexName,
    maxObjects, pageSize) {
  const all = new Map();
  let nbObjects = 0;
  let nbPages = 0;
  let pageToken = false;

  while (true) {
    nbPages += 1;

    const a = optionalArgs || {};

    if (pageSize !== undefined) {
      a.pageSize = Math.min(maxObjects - nbObjects, pageSize);
    }
    if (pageToken) {
      a.pageToken = pageToken;
    }

    const response = (arg === undefined ? func(a) : func(arg, a));

    const objects = response[listName] || [];

    const n = objects.length;
    Logger.log('%s objects retrieved in page %s', n, nbPages);
    if (n == 0) {
      break;
    }

    for (let i = 0; i < n; i++) {
      const obj = objects[i];
      all.set(obj[indexName], obj);
    }

    nbObjects += n;
    pageToken = response.nextPageToken;

    if ((!pageToken) || (nbObjects >= maxObjects)) {
      break;
    }
  }
  return all;
}
