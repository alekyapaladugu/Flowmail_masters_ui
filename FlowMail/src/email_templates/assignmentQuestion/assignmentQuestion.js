(function () {
  'use strict';

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    registerEvents();
  };

  function registerEvents() {
    document.getElementById("template-form").onsubmit = validateForm;
  }

  function sendMessage(message) {
    Office.context.ui.messageParent(message);
  }

  function validateForm() {
    console.log('hi')
    let data = {}
    data.professorName = document.getElementById('professor-name').value
    data.courseNumber = document.getElementById('course-no').value
    data.assignmentNo = document.getElementById('assignment-question-content').value
    data.studentId = document.getElementById('student-id').value
    data.studentName = document.getElementById('student-name').value

    sendMessage(JSON.stringify(data))
  }

})();