/**
 * Class DriveHelper
 *
 * A class for files and folders helpers
 *
 * Christophe Bisière
 *
 * version 2016-12-07
 *
 */

var DriveHelper = class DriveHelper {

  /**
   * Extract a Google document id from a URL
   *
   * Works for the following formats:
   * https://docs.google.com/a/tsm-education.fr/document/d/1CfHFmsCD0Fs3c3Nhg4F2oMUF6ARypJapwiWKCpOVwqQ/edit?usp=drivesdk
   * https://drive.google.com/open?id=0B8mYKkDGgfGyfl83MkUwMmN3UGVESWZhODBHdkhNcjV5MTVRaU0xTmNic2RJQkJLYmtBZWs&authuser=0
   *
   */
  static getIdfromUrl(url) {
    let matches = url.match(/id\=([^&]+)/);
    if (!matches) {
      matches = url.match(/([-\w]{25,})/);
    }

    if (!matches) {
      throw new Error(i18n([
        'Cannot find Google document Id in URL \'' + url + '\'.',
        'Impossible de retrouver l\'Id du document Google dans l\'URL \'' +
        url + '\'.',
      ]));
    }
    return matches[1];
  }

  /**
   * Get a folder by id, or null if it does not exist
   *
   */
  static getFolderByIdNoFail(id)
  {
    try {
      var folder = DriveApp.getFolderById(id);
      return folder;
    } catch(e) {
    }
    return null;
  }

  /**
   * Return the first parent folder of a folder or a file
   *
   */
  static getFirstParent(doc)
  {
    var id = doc.getId();
    var file = DriveApp.getFileById(id);
    var folder = file.getParents().next();
    return folder;
  }

  /**
   * Remove all parent folders from a file
   *
   */

  static removeAllParentsFromFile(file)
  {
    var parents = file.getParents();
    while (parents.hasNext()) {
      var parent = parents.next();
      Logger.log("Removing file \"" + file.getName() + "\" from parent folder \"" + parent.getName() + "\"");
      parent.removeFile(file);
    }
  }

  /**
   * Remove all parent folders from a folder
   *
   */

  static removeAllParentsFromFolder(folder)
  {
    var parents = folder.getParents();
    while (parents.hasNext()) {
      var parent = parents.next();
      Logger.log("Removing folder \"" + folder.getName() + "\" from parent folder \"" + parent.getName() + "\"");
      parent.removeFolder(folder);
    }
  }

  /**
   * Add the parents of file from_file to the list of parents of file
   */

  static addParentsToFile(from_file, file)
  {
    /* Add parent folders of from_file to file */
    var parents = from_file.getParents();
    while (parents.hasNext()) {
      var parent = parents.next();
      Logger.log("Adding file \"" + file.getName() + "\" to parent folder \"" + parent.getName() + "\"");
      parent.addFile(file);
    }
  }

  /**
   * Add the parents of file from_file to the list of parents of to_folder
   */

  static addParentsToFolder(from_file, folder)
  {
    /* Add parent folders of from_file to to_folder */
    var parents = from_file.getParents();
    while (parents.hasNext()) {
      var parent = parents.next();
      Logger.log("Adding folder \"" + folder.getName() + "\" to parent folder \"" + parent.getName() + "\"");
      parent.addFolder(folder);
    }
  }


  /**
   * Remove all viewers but one
   *
   */
  static removeAllViewersFrom(file, userToKeep) {
    for (user of file.getViewers()) {
      if (user && user.getEmail() != userToKeep.getEmail()) {
        Logger.log("Removing viewer " + user.getEmail() + " from file " + file.getName());
        file.removeViewer(user);
      }
    }
  }

  /**
   * Return the list of all commenters of a file.
   *
   */
  static getCommenters(file) {
      return file.getViewers().filter(user => file.getAccess(user) === 'COMMENT')
  } 

  /**
   * Remove all commenters but one
   *
   */
  static removeAllCommentersFrom(file, userToKeep) {
    for (user of DriveHelper.getCommenters(file)) {
      if (user && user.getEmail() != userToKeep.getEmail()) {
        Logger.log("Removing viewer " + user.getEmail() + " from file " + file.getName());
        file.removeCommenter(user);
      }
    }
  }

  /**
   * Remove all editors but one
   *
   */
  static removeAllEditorsFrom(file, userToKeep) {
    for (user of file.getEditors()) {
      if (user && user.getEmail() != userToKeep.getEmail()) {
        Logger.log("Removing viewer " + user.getEmail() + " from file " + file.getName());
        file.removeEditor(user);
      }
    }
  }


  /**
   * Copy file name from a file to another
   */

  static copyNameTo(fromFile, toTile)
  {
    var name = fromFile.getName();
    Logger.log("Setting name of file \"" + toTile.getName() + "\" to \"" + name + "\"");
    toTile.setName(name);
  }


  /**
   * Adda an editor to a file, without notifying the user
   */

  static setOwnerQuiet(email, file)
  {
    Drive.Permissions.insert(
      {
        'role': 'owner',
        'type': 'user',
        'value': email
      },
      file.getId(),
      {
        'sendNotificationEmails': 'false'
      }
    );
  }

  /**
   * Add an editor to a file, without notifying the user
   */

  static addEditorQuiet(email, file)
  {
    Drive.Permissions.insert(
      {
        'role': 'writer',
        'type': 'user',
        'value': email
      },
      file.getId(),
      {
        'sendNotificationEmails': 'false'
      }
    );
  }

  /**
   * Add a viewer to a file, without notifying the user
   */

  static addViewerQuiet(email, file)
  {
    Drive.Permissions.insert(
      {
        'role': 'reader',
        'type': 'user',
        'value': email
      },
      file.getId(),
      {
        'sendNotificationEmails': 'false'
      }
    );
  }

  /**
   * Add a commenter to a file, without notifying the user
   */

  static addCommenterQuiet(email, file)
  {
    Drive.Permissions.insert(
      {
        'role': 'reader',
        'additionalRoles': ['commenter'],
        'type': 'user',
        'value': email
      },
      file.getId(),
      {
        'sendNotificationEmails': 'false'
      }
    );
  }


  /**
   * Copy all properties but ownership from a file to another
   */

  static copyFilePropertiesTo(fromFile, toFile)
  {
    /* Name */
    copyNameTo(fromFile, toFile);

    /* Parent folders */
    addParentsToFile(fromFile, toFile);

    /* Description */
    var description = fromFile.getDescription();
    if (description == null) {
      description = "";
    }
    Logger.log("Setting description of file \"" + toFile.getName() + "\" to \"" + description + "\"");
    toFile.setDescription(description);

    /* Starred ? [for the user running the script only?] */
    var starred = fromFile.isStarred();
    Logger.log("Setting starred of file \"" + toFile.getName() + "\" to " + starred);
    toFile.setStarred(starred);

    /* Permissions: viewers and commenters */
    var viewers = fromFile.getViewers();
    for (var i = 0; i < viewers.length; i++) {
      var viewer = viewers[i];

      if (fromFile.getAccess(viewer) == DriveApp.Permission.COMMENT) {
        Logger.log("Adding commenter to file \"" + toFile.getName() + "\": \"" + viewer.getEmail() + "\"");
        addCommenterQuiet(viewer.getEmail(), toFile);
      } else {
        Logger.log("Adding viewer to file \"" + toFile.getName() + "\": \"" + viewer.getEmail() + "\"");
        addViewerQuiet(viewer.getEmail(), toFile);
      }
    }

    /* Permissions: editors */
    var editors = fromFile.getEditors();
    for (var i = 0; i < editors.length; i++) {
      var editor = editors[i];

      Logger.log("Adding editor to file \"" + toFile.getName() + "\": \"" + editor.getEmail() + "\"");
      addEditorQuiet(editor.getEmail(), toFile);
    }

    /* Shareable by editors ? */
    var shareable = fromFile.isShareableByEditors();
    Logger.log("Setting shareable by editors of file \"" + toFile.getName() + "\" to " + shareable);
    toFile.setShareableByEditors(shareable);

    /* Sharing */
    var access = fromFile.getSharingAccess();
    var permission = fromFile.getSharingPermission();
    Logger.log("Setting sharing access of file \"" + toFile.getName() + "\" to " + access + " / " + permission);
    toFile.setSharing(access, permission);

  //  /* Owner */
  //  var to_owner = to_file.getOwner();
  //  if (preserved_user == null || to_owner.getEmail() != preserved_user.getEmail()) {
  //    var from_owner = from_file.getOwner();
  //    Logger.log("Setting owner to " + owner.getEmail());
  //    to_file.setOwner(from_owner);
  //  }
  }


  /**
   * Silently trash a file
   *
   */
  static trashNoFail(file) {
    if (file != null && !file.isTrashed()) {
      try {
        file.setTrashed(true);
        Logger.log("Trashed \"" + file.getName() + "\"");
      } catch (e) {
      }
    }
  }

  /**
   * Silently trash a file by id
   *
   */
  static trashByIdNoFail(id) {
    try {
      const file = DriveApp.getFileById(id);
      DriverHelper.trashNoFail(file);
    } catch (e) {
    }
  }

  static getBlobAs(id, mime) {
    const url = 'https://www.googleapis.com/drive/v3/files/' + id + '/export?mimeType=' + mime;
    const blob = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true
    }).getBlob();
    return blob;
  }
}
