function sendReminderEmailSecondWeek() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Form Responses 1'); // Replace 'Sheet1' with the appropriate sheet name

  // Data from quiz sheet
  var quizData = quizSheet.getDataRange().getValues();

  // Data from employee sheet
  var employeeData = halodocEmployeeData();

  // Data on employees who have not completed the quiz
  var employeeNotCompleteQuiz = checkEmployeeQuiz(employeeData, quizData)["not complete"];

  // Data on employees whose quiz scores are insufficient
  var employeeInsufficientValue = checkEmployeeQuiz(employeeData, quizData)["insufficient"];

  // Chief data
  var chiefEmailData = chiefData(employeeData);

  // Send email
  var sendEmail = sendAutomaticEmailChief(employeeNotCompleteQuiz, employeeInsufficientValue, chiefEmailData);
  Logger.log(sendEmail); // Email info sent or not
}

function sendAutomaticEmailChief(employeeNotCompleteQuiz, employeeInsufficientValue, chiefEmailData) {

  // Send email to Chief
  if (employeeNotCompleteQuiz.length > 0 || employeeInsufficientValue.length > 0) {

    for (var chief in chiefEmailData) {
      // Body Email
      var emailBody = `
        <html>
          <head>
            <style>
            body {
              font-family: Arial, sans-serif;
              margin: 0;
              padding: 0;
              line-height: 1.6;
            }
            .container {
              max-width: 600px;
              margin: 20px auto;
              padding: 20px;
              border: solid 1px #9370DB;
              background-color: #fff;
              border-radius: 10px;
              box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
              float: left;
            }
            h1 {
              font-size: 16px;
              margin-bottom: 20px;
            }
            p, td {
              color: #666;
              font-size: 16px;
              margin-bottom: 10px;
              text-align: justify;
            }
            table {
              width: 100%;
              border-collapse: collapse;
              margin-bottom: 20px;
            }
            th, td {
              border: 1px solid #9370DB;
              padding: 8px;
              text-align: left;
            }
            th {
              color: #9370DB;
              font-size: 16px;
              margin-bottom: 10px;
              background-color: #E6E6FA;
            }
            .persentase {
              width: 50%;
            }
            </style>
          </head>
          <body>
            <div class="container">
              <h1>Dear Bapak/Ibu ${chief},</h1>
              <p>As the follow up of our previous update on information security and data privacy awareness program on Date. Herewith, we would like to provide you with the update of the completion status on the Information Security and Data Privacy awareness program with the topic of “Social Engineering Attack” and “8 Principles of Data Protection” in your department, as follow:</p>
      `;
      
      if (employeeNotCompleteQuiz.length > 0) {
        emailBody += `
          <h1>Halosquads who have not participated:</h1>
          <table>
          <tr><th>Name</th><th>Email</th><th>Departement</th></tr>
        `;

        var y = 0;
        for (var k = 0; k < employeeNotCompleteQuiz.length; k++) {
          if (employeeNotCompleteQuiz[k]["chief"] == chief) {

            // Write in the body of the email
            y += 1;
            emailBody += <tr><td>${y}. ${employeeNotCompleteQuiz[k]["name"]}</td><td>${employeeNotCompleteQuiz[k]["email"]}</td><td>${employeeNotCompleteQuiz[k]["departement"]}</td></tr>;
          }
        }
        emailBody += </table>;
      }

      if (employeeInsufficientValue.length > 0) {
        emailBody += `
          <h1>Participated Halosquads but quiz score below 70 (Failed):</h1>
          <table>
          <tr><th>Name</th><th>Email</th><th>Score</th><th>Departement</th></tr>
        `;

        var y = 0;
        for (var k = 0; k < employeeInsufficientValue.length; k++) {
          if (employeeInsufficientValue[k]["chief"] == chief) {
            
            // Write in the body of the email
            y += 1;
            emailBody += <tr><td>${y}. ${employeeInsufficientValue[k]["name"]}</td><td>${employeeInsufficientValue[k]["email"]}</td><td>${employeeInsufficientValue[k]["score"]}</td><td>${employeeInsufficientValue[k]["departement"]}</td></tr>;
          }
        }

        emailBody += </table>;
      }

      emailBody += `
              <p>We would like to ask for your support to take a moment to follow up with the above Halosquads to ensure participation, quiz completion and understanding of the material. Your support in this matter is greatly appreciated.</p>
              <p>Thank you for your attention to this important initiative.</p>
              <p>Best regards,<br>ISDP - Security GRC and Data Privacy Team</p>
            </div>
          </body>
        </html>
      `;
      
      var ccEmail = ["email1@gmail.com","email2@gmail.com"];
      MailApp.sendEmail({
        to: chiefEmailData[chief],
        cc: ccEmail.join(','),
        subject: '2nd Week Update - Q2 2024',
        htmlBody: emailBody
      });
    }    
    return "Chief email sent successfully";
  } else {
    return "Chief email was not sent successfully";
  }
}
