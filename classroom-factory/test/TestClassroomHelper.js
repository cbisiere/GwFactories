/**
 * Classroom helpers (tests)
 *
 * Christophe Bisière
 *
 * version 2020-12-08
 *
 */

function testSetCourseName() {
  let course = ClassroomHelper.setCourseName('171690719648', 'Test class NEW name');

  Logger.log('Got: "%s"', course);
}

function testSetCourseOwner() {
  let course = ClassroomHelper.setCourseOwner('171690719648', 'laurence.mahul@tsm-education.fr');

  Logger.log('Got: "%s"', course);
}

function testGetCourse() {
  let course = ClassroomHelper.getCourse('137570615802');

  Logger.log('Got: "%s"', course);
}

function testGetCourses() {
  let courses = ClassroomHelper.getCourses(['ACTIVE']);

  Logger.log('%s courses retrieved', courses.size);

  for (let [id, course] of courses) {
    Logger.log('%s (%s)', course.name, course.id);
  }
}

function testGetTopics() {
  let topics = ClassroomHelper.getTopics('137570615802');

  Logger.log('%s topics retrieved', topics.size);

  for (let [id, topic] of topics) {
    Logger.log('%s (%s)', topic.name, topic.topicId);
  }
}

function testGetTeachers() {
  let teachers = ClassroomHelper.getTeachers('137570615802');

  Logger.log('%s teachers retrieved', teachers.size);

  for (let [id, teacher] of teachers) {
    Logger.log('%s (%s)', teacher.profile.emailAddress, id);
  }
}

function testGetStudents() {
  let students = ClassroomHelper.getStudents('137570615802');

  Logger.log('%s students retrieved', students.size);

  for (let [id, student] of students) {
    Logger.log('%s (%s)', student.profile.emailAddress, student.userId);
  }
}

function testGetInvitations() {
  let invitations = ClassroomHelper.getInvitations('337398865288');

  Logger.log('%s invitations retrieved', invitations.size);

  for (let [id, invitation] of invitations) {
    Logger.log('%s as %s (%s)', Classroom.UserProfiles.get(invitation.userId).emailAddress, invitation.role, invitation.id);
  }
}
