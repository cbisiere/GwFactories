/**
 * Class CourseTable
 *
 * A CourseTable is a table containing data about Classroom courses.
 *
 * Christophe Bisière
 *
 * version 2020-11-18
 *
 */

/**
 * Headers in the Course table.
 *
 */

/* fields used when creating a new Course */
const COL_CLASS_NAME = "Class Name";               /* name of the class */
const COL_CLASS_HEADING = "Class Heading";         /* description heading (USAGE??) */
const COL_CLASS_DESCRIPTION = "Class Description"; /* description of the class */
const COL_CLASS_SECTION = "Class Section";         /* section of the class */
const COL_CLASS_ROOM = "Class Room";               /* room name */
const COL_CLASS_OWNER = "Class Owner";             /* class owner: address or 'me' */
const COL_CLASS_TEACHERS = "Class Teachers";       /* comma, semi-colon or space separated list of co-teachers */
const COL_CLASS_TOPICS = "Class Topics";           /* comma, semi-colon or space separated list of quoted topics */
const COL_CLASS_STATE = "Class State";             /* class State */

/* read-only fields */
const COL_CLASS_ID = "Class Id";                       /* class ID (set by the script) */
const COL_CLASS_CODE = "Class Code";                   /* class enrollment code (set by the script) */
const COL_CLASS_URL = "Class Url";                     /* class Url (set by the script) */
const COL_CLASS_CREATION_TIME = "Class Creation Time"; /* class creation time (set by the script) */
const COL_CLASS_UPDATE_TIME = "Class Update Time";     /* class creation time (set by the script) */

/* status info */
const COL_STATUS = "Status";


/**
 * Class representing Classroom courses.
 *
 */
class CourseTable {

  /**
   * Create a CourseTable.
   * @param {Range} r - The Range for the whole table.
   */
  constructor(r) {

    this.cols = [
      COL_CLASS_NAME,
      COL_CLASS_SECTION,
      COL_CLASS_HEADING,
      COL_CLASS_DESCRIPTION,
      COL_CLASS_ROOM,
      COL_CLASS_OWNER,
      COL_CLASS_TEACHERS,
      COL_CLASS_TOPICS,
      COL_CLASS_STATE,
      COL_CLASS_CODE,
      COL_CLASS_URL,
      COL_CLASS_CREATION_TIME,
      COL_CLASS_UPDATE_TIME,
      COL_CLASS_ID,
      COL_STATUS
    ];

    const defaults = new Map([
      [COL_CLASS_OWNER, 'me'],
      [COL_CLASS_STATE, 'PROVISIONED'],
    ]);

    const formats = new Map([
      [COL_CLASS_ID, LF.SheetHelper.setRangeFormatAsText],
      [COL_CLASS_NAME, LF.SheetHelper.setRangeFormatAsText],
      [COL_CLASS_HEADING, LF.SheetHelper.setRangeFormatAsText],
      [COL_CLASS_DESCRIPTION, LF.SheetHelper.setRangeFormatAsText],
      [COL_CLASS_SECTION, LF.SheetHelper.setRangeFormatAsText],
      [COL_CLASS_ROOM, LF.SheetHelper.setRangeFormatAsText],
      [COL_CLASS_OWNER, LF.SheetHelper.setRangeFormatAsText],
      [COL_CLASS_TEACHERS, LF.SheetHelper.setRangeFormatAsText],
      [COL_CLASS_TOPICS, LF.SheetHelper.setRangeFormatAsText],
      [COL_CLASS_ID, LF.SheetHelper.setRangeFormatAsText],
      [COL_CLASS_URL, LF.SheetHelper.setRangeFormatAsText],
      [COL_CLASS_STATE, CourseTable.setRangeFormatAsCourseState],
      [COL_CLASS_CREATION_TIME, LF.SheetHelper.setRangeFormatAsDatetime],
      [COL_CLASS_UPDATE_TIME, LF.SheetHelper.setRangeFormatAsDatetime],
      [COL_CLASS_CODE, LF.SheetHelper.setRangeFormatAsText],
    ]);

    this.dt = new LF.DataTable(r, defaults, formats);

  }


  /* static members */

  /**
   * Locate the target Classroom table in a Sheet
   *
   * @param {Sheet} sheet - The Sheet containing the Classroom table to search for.
   * @return {?Range} The Range of the Classroom table found.
   */
  static locate(sheet) {
    return LF.DataTable.locateFromLabel(sheet, COL_CLASS_NAME);
  }

  /**
   * Format a cell as course status.
   *
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/courses#CourseState}
   *
   * @param {Range} c - The cell to format as course status.
   */
  static setRangeFormatAsCourseState(c) {
    let validation = SpreadsheetApp.newDataValidation()
      .setAllowInvalid(false)
      .requireValueInList(['ACTIVE', 'ARCHIVED', 'PROVISIONED', 'DECLINED', 'SUSPENDED'], true)
      .build();
    c.setDataValidation(validation);
  }

  /**
   * Check a condition.
   *
   * @param {boolean} condition - The condition to check.
   * @param {string} message - The message to display if the condition is not met.
   */
  assert(condition, message) {
    let a1 = this.getRange() == undefined ? 'undefined' : this.getRange().getA1Notation(); /* "==" so also null */
    let prompt = 'Error: CourseTable (' + a1 + '): ' + (message || ' assertion failed');
    assert(condition, prompt);
  }

  /* get: dimensions, has label... */

  /**
   * Return the course table (including header) as a Range object.
   *
   * @return {Range|undefined} The whole course table as a Range object.
   */
  getRange() {
    return this.dt.getRange();
  }

  /**
   * Return the number of courses.
   *
   * @return {number} The number of courses in the course table.
   */
  getNumCourses() {
    return this.dt.getNumRows();
  }

  /**
   * Return the number of columns.
   *
   * @return {number} The number of columns in the data table.
   */
  getNumColumns() {
    return this.dt.getNumColumns();
  }

  /**
   * True whether a given column label exists.
   *
   * @param {string} label - The column label.
   * @return {boolean} True if the column exist.
   */
  has(label) {
    return this.dt.has(label);
  }


  /* get: maps */

  /**
   * Return a Map of courses, each being a map column label to cell value
   *  for a given course.
   *
   * @return {Map} The map of course maps.
   */
  getCoursesAsMaps() {
    return this.dt.getDataAsMaps();
  }

  /**
   * Return an empty Map of column label to undefined values.
   *
   * @return {Map} The map of column label to undefined values.
   */
  getEmptyCourseAsMap() {
    return this.dt.getEmptyRowAsMap();
  }

  /* set: Map */

  /**
   * Set a course row from a Map of column label to cell value.
   *
   * @param {number} i - The course number.
   * @param {Map} cmap - The label-to-value map.
   */
  setCourseFromMap(i, cmap) {
    this.dt.setRowFromMap(i, cmap);
  }

  /*
  * Appends a new course from a map.
  *
  * @param {Map} m - The map.
  */
  addCourseFromMap(m) {
    this.dt.addRowFromMap(m);
  }

  /*
  * Appends new courses from a set of maps.
  *
  * @param {Set} cmaps - The set of maps.
  */
  addCoursesFromSet(cmaps) {
    this.dt.addRowsFromSet(cmaps);
  }

  /**
   * Complete a map with defaults values when specified.
   *
   * This may create new labels in the map.
   *
   * @param {Map} cmap - The course to which apply defaults.
   */
  applyDefaultsToMap(cmap) {
    this.dt.applyDefaultsToMap(cmap);
  }

  /*
   * Update a course map using a Google course object.
   *
   * TODO: return a boolean "something updated"
   *
   * @param {Map} cmap - The course map to update.
   * @param {course} course - The course object.
   */
  updateMapFromCourseObject(cmap, oCourse, invitations = false) {

    Logger.log("updateMapFromCourseObject (IN):");
    LF.logMap(cmap, "cmap");

    /* 1) update fields that are directly mapped with Google Course properties */

    const fields = new Map([
      [COL_CLASS_ID, 'id'],
      [COL_CLASS_NAME, 'name'],
      [COL_CLASS_SECTION, 'section'],
      [COL_CLASS_HEADING, 'descriptionHeading'],
      [COL_CLASS_DESCRIPTION, 'description'],
      [COL_CLASS_ROOM, 'room'],
      [COL_CLASS_STATE, 'courseState'],
      [COL_CLASS_CODE, 'enrollmentCode'],
      [COL_CLASS_URL, 'alternateLink']
    ]);

    for (let [label, property] of fields) {
      if (this.has(label)) {
        let newValue = oCourse[property];
        Logger.log("Updating field \"%s\" with value \"%s\"", label, newValue);
        cmap.set(label, newValue);
      }
    }

    /* 2) update other fields */

    if (this.has(COL_CLASS_CREATION_TIME)) {
      let newValue = new Date(oCourse.creationTime);
      Logger.log("Updating field \"%s\" with value \"%s\"", COL_CLASS_CREATION_TIME, newValue);
      cmap.set(COL_CLASS_CREATION_TIME, newValue);
    }

    if (this.has(COL_CLASS_UPDATE_TIME)) {
      let newValue = new Date(oCourse.updateTime);
      Logger.log("Updating field \"%s\" with value \"%s\"", COL_CLASS_UPDATE_TIME, newValue);
      cmap.set(COL_CLASS_UPDATE_TIME, newValue);
    }

    if (this.has(COL_CLASS_OWNER)) {
      let newValue = Classroom.UserProfiles.get(oCourse.ownerId).emailAddress;
      Logger.log("Updating field \"%s\" with value \"%s\"", COL_CLASS_OWNER, newValue);
      cmap.set(COL_CLASS_OWNER, newValue);
    }

    // TODO: also use invitations (or have a new column for that)
    if (!invitations && this.has(COL_CLASS_TEACHERS)) {
      let teachers = ClassroomHelper.getTeachers(oCourse.id);
      let emails = []
      for (let teacher of teachers.values()) {
        emails.push(teacher.profile.emailAddress);
      }
      let newValue = emails.join(', ');
      Logger.log("Updating field \"%s\" with value \"%s\"", COL_CLASS_TEACHERS, newValue);
      cmap.set(COL_CLASS_TEACHERS, newValue);
    }

    if (this.has(COL_CLASS_TOPICS)) {
      let topics = ClassroomHelper.getTopics(oCourse.id);
      let names = []
      for (let topic of topics.values()) {
        names.push("\"" + topic.name + "\"");
      }
      let newValue = names.join(', ');
      Logger.log("Updating field \"%s\" with value \"%s\"", COL_CLASS_TOPICS, newValue);
      cmap.set(COL_CLASS_TOPICS, newValue);
    }

    Logger.log("updateMapFromCourseObject (OUT):");
    LF.logMap(cmap, "cmap");

  }

  /*
   * Create a new Google Classroom from a map.
   *
   * @param {Map} cmap - The course map to use to create the new classroom.
   * @return {course} - The new Google course object.
   */
  createClassroomCourseFromMap(cmap) {

    /* create the course object to send to Google */
    var oCourse = new Object();

    /* 1) create the Google Classroom Course */

    const fields = new Map([
      [COL_CLASS_ID, 'id'],
      [COL_CLASS_NAME, 'name'],
      [COL_CLASS_SECTION, 'section'],
      [COL_CLASS_HEADING, 'descriptionHeading'],
      [COL_CLASS_DESCRIPTION, 'description'],
      [COL_CLASS_ROOM, 'room'],
      [COL_CLASS_OWNER, 'ownerId'],
      [COL_CLASS_STATE, 'courseState'],
    ]);

    for (const [label, cFieldName] of fields) {
      if (cmap.has(label)) {
        oCourse[cFieldName] = cmap.get(label);
      }
    }
    LF.logObject(oCourse, "course");

    var oCourse = Classroom.Courses.create(oCourse);
    Logger.log('Course created: \"%s\" (%s)', oCourse.name, oCourse.id)

    /* 2) invite the teachers */
    if (cmap.has(COL_CLASS_TEACHERS) && cmap.get(COL_CLASS_TEACHERS).length > 0) {
      let teachers = LF.itemsInString(cmap.get(COL_CLASS_TEACHERS), "[\\s,;]+")
      if (teachers !== null) {
        Logger.log("Number of teachers to invite: %s", teachers.length);
        Logger.log("Teachers to invite: %s", teachers);
        for (const teacher of teachers) {
          /* send the invitation to all the teachers but the owner (as it triggers an error) */
          if (teacher == Classroom.UserProfiles.get(oCourse.ownerId).emailAddress) {
            Logger.log("Owner \"%s\" will not be invited", teacher);
          } else {
            Logger.log("About to invite \"%s\"", teacher);

            var oInvitation = {
              'userId': teacher,
              'courseId': oCourse.id,
              'role': 'TEACHER'
            }
            var oInvitation = Classroom.Invitations.create(oInvitation);
            Logger.log("Invitation created: \"%s\" (%s)", oInvitation.userId, oInvitation.id);
          }
        }
      }
    }

    /* 3) create the topics (in reverse order to have them in the order of the array)*/
    if (cmap.has(COL_CLASS_TOPICS) && cmap.get(COL_CLASS_TOPICS).length > 0) {
      let topics = LF.quotedItemsInString(cmap.get(COL_CLASS_TOPICS))
      if (topics !== null) {
        Logger.log("Number of topic to create: %s", topics.length);

        topics.reverse();
        for (const s of topics) {

          var oTopic = {
            'name': s.trim()
          }
          var oTopic = Classroom.Courses.Topics.create(oTopic, oCourse.id);
          Logger.log("Topic created: \"%s\" (%s)", oTopic.name, oTopic.topicId);
        }
      }
    }

    return oCourse;
  }

  /* high level functions */

  /*
   * Refresh existing course data.
   *
   */
  refresh() {
    this.assert(this.has(COL_CLASS_ID), "missing column:'" + COL_CLASS_ID + "'");

    let courses = this.getCoursesAsMaps();
    Logger.log("Number of cmaps in the map: %s", courses.size);

    for (let [i, c] of courses) {
      let v = c.get(COL_CLASS_ID);
      let courseId = v === undefined ? '' : v.toString().trim();
      if (courseId.length > 0) {
        Logger.log("Looking up for course id: %s", courseId);

        let oCourse = ClassroomHelper.getCourse(courseId);
        Logger.log("Got Course: %s", oCourse);

        if (oCourse !== null) {
          this.updateMapFromCourseObject(c, oCourse);
          this.setCourseFromMap(i, c);
        }
      }
    }
  }

  /*
   * Update course data, possibly adding new course rows.
   *
   */
  update() {

    /* is lookup even possible? */
    let lookup = this.has(COL_CLASS_ID) && (this.getNumCourses() > 0);

    let oCourses = ClassroomHelper.getCourses(['ACTIVE']); //TODO: param & selector?
    let courses = this.getCoursesAsMaps();
    let newCourses = new Set();

    for (let [courseId, oCourse] of oCourses) {

      Logger.log("Table lookup for course \"%s\" (%s)", oCourse.name, courseId);

      /* find the [index, row] of the course id in the data table */
      let res = (lookup ? LF.findMap(courses, COL_CLASS_ID, courseId) : null);

      if (res !== null) {
        let [i, cmap] = res;
        this.updateMapFromCourseObject(cmap, oCourse);
        this.setCourseFromMap(i, cmap);
      } else {
        let cmap = this.getEmptyCourseAsMap();
        this.updateMapFromCourseObject(cmap, oCourse);
        newCourses.add(cmap);
      }
    }
    this.addCoursesFromSet(newCourses);
  }

  /*
   * Create new courses from course data rows without course ids.
   *
   * @return {Array} - number of courses created, number of errors.
   */
  create() {
    this.dt.ensureColumnsExist([COL_CLASS_ID, COL_STATUS]);

    /* various counters */
    var creation_count = 0;
    var error_count = 0;

    /* create the courses */
    let rows = this.getCoursesAsMaps(); // FIXME: exclude read only fields?
    for (var [i, row] of rows) {

      this.applyDefaultsToMap(row);
      LF.trimStringsInMap(row);

      /* do nothing if the class id is set */
      if (!row.has(COL_CLASS_ID) || row.get(COL_CLASS_ID).length > 0) {
        continue;
      }

      /* do nothing if class name is empty */
      if (row.get(COL_CLASS_NAME).length == 0) {
        continue;
      }

      try {
        let course = this.createClassroomCourseFromMap(row);

        // Do not owerwrite teachers and students,
        // since they had no chance to accept the invitations!
        this.updateMapFromCourseObject(row, course, true);
        this.setCourseFromMap(i, row);

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
      [COL_CLASS_NAME, "Business Economics"],
      [COL_CLASS_SECTION, "Spring 2021"],
      [COL_CLASS_HEADING, "Welcome to Business Economics."],
      [COL_CLASS_DESCRIPTION, "We will be learning to analyze firm’s decisions by applying principles of economic analysis."],
      [COL_CLASS_ROOM, "301"],
      [COL_CLASS_OWNER, "norman.bates@teacher.abc.edu"],
      [COL_CLASS_TEACHERS, "jane.doe@teacher.abc.edu, john.doe@teacher.abc.edu"],
      [COL_CLASS_TOPICS, '"Syllabus", "Week 1", "Week 2", "Exams"']
    ]);
    this.addCourseFromMap(m);
  }

  /**
   * Complete the table with all missing columns
   *
   */
  complete() {
    this.dt.ensureColumnsExist(this.cols);
  }
}
