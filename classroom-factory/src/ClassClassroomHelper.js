/**
 * Class ClassroomHelper
 *
 * A class for Classroom helpers
 *
 * Christophe Bisière
 *
 * version 2021-05-13
 *
 */

class ClassroomHelper {
  /**
   * Change the name of a course.
   *
   * The caller must be an admin or the current owner.
   *
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/courses/patch}
   *
   * @param {string} courseId The id of the course
   * @param {string} newName New name of the course
   * @return {?Course} The updated course.
   */
  static setCourseName(courseId, newName) {
    const course = {'name': newName};
    return Classroom.Courses.patch(course, courseId,  {'updateMask': 'name'});
  }

  /**
   * Change the owner of a course.
   *
   * The caller must be an admin.
   *
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/courses/patch}
   *
   * @param {string} courseId The id of the course
   * @param {string} newOwnerId Id of the new owner
   * @return {?Course} The updated course.
   */
  static setCourseOwner(courseId, newOwnerId) {
    const course = {'ownerId': newOwnerId};
    return Classroom.Courses.patch(course, courseId, {'updateMask': 'ownerId'});
  }

  /**
   * Retrieves a course by id.
   *
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/courses/get}
   *
   * @param {string} courseId The id of the course
   * @return {?Course} The course or null if it cannot be retrieved.
   */
  static getCourse(courseId) {
    let course = null;
    try {
      course = Classroom.Courses.get(courseId);
      Logger.log('Course "%s" (%s) found', course.name, courseId);
    } catch (err) {
      Logger.log('Course with id "%s" not found', courseId);
      course = null;
    }
    return course;
  }

  /**
   * Returns a map of courses, indexed by course id.
   *
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/courses#Course}
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/courses/list}
   *
   * @param {?Array.<CourseState>} courseStates Restricts returned courses to
   *  those in one of these states
   * @return {!Map<string,Course>} A map of Course objects indexed by class id.
   */
  static getCourses(courseStates) {
    const optionalArgs = {};
    if (courseStates !== undefined) {
      optionalArgs.courseStates = courseStates;
    }
    return LF.getGoogleList(undefined, optionalArgs,
        Classroom.Courses.list, 'courses', 'id');
  }

  /**
   * Returns a map of topics, indexed by topic id.
   *
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/courses.topics#Topic}
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/courses.topics/list}
   *
   * @param {string} courseId The id of the course for which topics are returned
   * @return {!Map<string,Topic>} A map of Topic objects indexed by topic id.
   */
  static getTopics(courseId) {
    return LF.getGoogleList(courseId, undefined, Classroom.Courses.Topics.list,
        'topic', 'topicId');
  }

  /**
   * Returns a map of teachers, indexed by user id.
   *
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/courses.teachers#Teacher}
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/courses.teachers/list}
   *
   * @param {string} courseId The id of the course for which teachers are returned
   * @return {!Map<string,Teacher>} A map of Teacher objects indexed by user id.
   */
  static getTeachers(courseId) {
    const teachers = LF.getGoogleList(courseId, undefined,
        Classroom.Courses.Teachers.list, 'teachers', 'userId');
    /* return the map in reverse order, as the Google list is itself reversed */
    return new Map(Array.from(teachers).reverse());
  }

  /**
   * Returns a map of students, indexed by user id.
   *
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/courses.students#Student}
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/courses.students/list}
   *
   * @param {string} courseId The id of the course for which students are returned
   * @return {!Map<string,Student>} A map of Student objects indexed by user id.
   */
  static getStudents(courseId) {
    return LF.getGoogleList(courseId, undefined, Classroom.Courses.Students.list, 'students', 'userId');
  }

  /**
   * Returns a map of invitations sent to a user for a given course, indexed by invitation id.
   *
   * At least the user or the course must be specified (or both).
   *
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/invitations#Invitation}
   * @see {@link https://developers.google.com/classroom/reference/rest/v1/invitations/list}
   *
   * @param {string} courseId The id of the course for which invitations are returned
   * @param {string} userId The id of the user for which invitations are returned
   * @return {!Map<string,Student>} A map of Invitation objects indexed by invitation id.
   */
  static getInvitations(courseId, userId) {
    const optionalArgs = {};
    if (courseId !== undefined) {
      optionalArgs.courseId = courseId;
    }
    if (userId !== undefined) {
      optionalArgs.userId = userId;
    }
    return LF.getGoogleList(undefined, optionalArgs, Classroom.Invitations.list,
        'invitations', 'id');
  }
}
