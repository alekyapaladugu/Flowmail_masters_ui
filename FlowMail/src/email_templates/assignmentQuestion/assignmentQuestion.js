(function(){
    'use strict';
  
    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function(reason){
        let data ={}
        data.professorName = document.getElementById('professor-name').value
        data.courseNumber = document.getElementById('course-no').value
        data.assignmentNo = document.getElementById('assignment-question-content').value
        data.studentId = document.getElementById('student-id').value
        data.studentName = document.getElementById('student-name').value
        if(data.length > 0) {
          document.getElementById("settings-icon").removeAttribute('disabled');
        }
        document.getElementById("template-form").onsubmit = sendMessage(JSON.stringify(data))
      
    };

  
    function sendMessage(message) {
      Office.context.ui.messageParent(message);
    }

  })();