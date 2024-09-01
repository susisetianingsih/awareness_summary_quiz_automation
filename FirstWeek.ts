function sendReminderEmailFirstWeek() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName('Form Responses 1'); // Replace 'Sheet1' with the appropriate sheet name

  // Data from quiz sheet
  var quizData = quizSheet.getDataRange().getValues();

  // Data from employee sheet
  var employeeData = halodocEmployeeData();

  // Data on employees who have not completed the quiz
  var employeeNotCompleteQuiz = checkEmployeeQuiz(employeeData, quizData)["not complete"];
  
  // Data on employees who have completed the quiz
  var employeeCompleteQuiz = checkEmployeeQuiz(employeeData, quizData)["complete"];
  
  // Data on employees whose quiz scores are insufficient
  var employeeInsufficientValue = checkEmployeeQuiz(employeeData, quizData)["insufficient"];
  
  // Data on employees whose quiz scores are sufficient
  var employeeSufficientValue = checkEmployeeQuiz(employeeData, quizData)["sufficient"];

  // HoD data
  var hodEmailData = hodData(employeeData);

  // Chief data
  var chiefEmailData = chiefData(employeeData);

  // Departement data
  var halodocDepartementData = departementData(employeeData);

  // Send email HoD
  var sendEmailHoD = sendAutomaticEmailHoD(employeeNotCompleteQuiz, employeeInsufficientValue, hodEmailData);
  Logger.log(sendEmailHoD); // Email info sent or not

  // Send email Chief
  var sendEmailChief = sendAutomaticEmailChiefResume(employeeNotCompleteQuiz, employeeCompleteQuiz, employeeInsufficientValue, employeeSufficientValue, halodocDepartementData, chiefEmailData);
  Logger.log(sendEmailChief); // Email info sent or not
}

function sendAutomaticEmailHoD(employeeNotCompleteQuiz, employeeInsufficientValue, hodEmailData) {

  // Send email to HoD
  if (employeeNotCompleteQuiz.length > 0 || employeeInsufficientValue.length > 0) {

    for (var hod in hodEmailData) {
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
            </style>
          </head>
          <body>
            <div class="container">
              <h1>Dear Bapak/Ibu ${hod},</h1>
              <p>As part of our ongoing efforts to enhance cybersecurity awareness within our organization, and compliance to ISO 27001 and ISO 27701, we have socialized and launched the information security and data privacy awareness program on Date. Herewith, we would like to provide you with an update as of today on the completion status of the Information Security and Data Privacy awareness program with the topic of “Social Engineering Attack” and “8 Principles of Data Protection” in your department, as follow:</p>
      `;

      if (employeeNotCompleteQuiz.length > 0) {
        emailBody += `
          <h1>Halosquads who have not participated:</h1>
          <table>
          <tr><th>Name</th><th>Email</th></tr>
        `;

        var y = 0;
        for (var k = 0; k < employeeNotCompleteQuiz.length; k++) {
          if (employeeNotCompleteQuiz[k]["hod"] == hod) {
            
            // Write in the body of the email
            y += 1;
            emailBody += <tr><td>${y}. ${employeeNotCompleteQuiz[k]["name"]}</td><td>${employeeNotCompleteQuiz[k]["email"]}</td></tr>;
          }
        }

        emailBody += </table>;
      }

      if (employeeInsufficientValue.length > 0) {
        emailBody += `
          <h1>Participated Halosquads but quiz score below 70 (Failed):</h1>
          <table>
          <tr><th>Name</th><th>Email</th><th>Score</th></tr>
        `;

        var y = 0;
        for (var k = 0; k < employeeInsufficientValue.length; k++) {
          if (employeeInsufficientValue[k]["hod"] == hod) {
            
            // Write in the body of the email
            y += 1;
            emailBody += <tr><td>${y}. ${employeeInsufficientValue[k]["name"]}</td><td>${employeeInsufficientValue[k]["email"]}</td><td>${employeeInsufficientValue[k]["score"]}</td></tr>;
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

      MailApp.sendEmail({
        to: hodEmailData[hod],
        subject: '1st Week Update on ISDP Awareness - Q2 2024',
        htmlBody: emailBody
      });
    }
    return "HoD email sent successfully";
  } else {
    return "HoD email was not sent successfully";
  }
}

function sendAutomaticEmailChiefResume(employeeNotCompleteQuiz, employeeCompleteQuiz, employeeInsufficientValue, employeeSufficientValue, halodocDepartementData, chiefEmailData) {

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
            margin-top: 20px;
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
        </style>
      </head>
      <body>
        <div class="container">
          <h1>Dear All,</h1>
          <p>As part of our ongoing efforts to enhance cybersecurity awareness within our organization, and compliance to ISO 27001 and ISO 27701, we have socialized and launched the information security and data privacy awareness program on Date. Herewith, we would like to provide you with an update as of today on the completion status of the Information Security and Data Privacy awareness program with the topic of “Social Engineering Attack” and “8 Principles of Data Protection” in each department, as follow:</p>
          <table>
            <tr><th>Department Name</th><th>Participated</th><th>Not Participated</th><th>Pass</th><th>Failed</th></tr>
  `;

  var y = 0;
  for (var data in halodocDepartementData) {
    var departementNotCompleteData = [];
    var departementCompleteData = [];
    var departementInsufficientData = [];
    var departementSufficientData = [];
    for (var k = 0; k < employeeNotCompleteQuiz.length; k++) {
      if (employeeNotCompleteQuiz[k]["departement"] == halodocDepartementData[data]) {
        departementNotCompleteData.push(employeeNotCompleteQuiz[k]);
      }
    }
    for (var k = 0; k < employeeCompleteQuiz.length; k++) {
      if (employeeCompleteQuiz[k]["departement"] == halodocDepartementData[data]) {
        departementCompleteData.push(employeeCompleteQuiz[k]);
      }
    }
    for (var k = 0; k < employeeInsufficientValue.length; k++) {
      if (employeeInsufficientValue[k]["departement"] == halodocDepartementData[data]) {
        departementInsufficientData.push(employeeInsufficientValue[k]);
      }
    }
    for (var k = 0; k < employeeSufficientValue.length; k++) {
      if (employeeSufficientValue[k]["departement"] == halodocDepartementData[data]) {
        departementSufficientData.push(employeeSufficientValue[k]);
      }
    }
    // Calculate the percentage of employees for each department
    var employeeNotComplete = (departementNotCompleteData.length / (departementNotCompleteData.length + departementCompleteData.length)) * 100;
    var employeeComplete = (departementCompleteData.length / (departementNotCompleteData.length + departementCompleteData.length)) * 100;
    var employeeInsufficient = (departementInsufficientData.length / (departementInsufficientData.length + departementSufficientData.length)) * 100;
    var employeeSufficient = (departementSufficientData.length / (departementInsufficientData.length + departementSufficientData.length)) * 100;

    // Write in the body of the email
    y += 1;
    emailBody += <tr><td>${y}. ${halodocDepartementData[data]}</td><td>${employeeComplete.toFixed(2)}%</td><td>${employeeNotComplete.toFixed(2)}%</td><td>${employeeSufficient.toFixed(2)}%</td><td>${employeeInsufficient.toFixed(2)}%</td></tr>;
  }

  emailBody += `
          </table>
          <p>We would like to ask for your support to encourage Halosquads in your team to participate, complete the quiz and understand the material until next week.</p>
          <p>Your support in this matter is greatly appreciated.</p>
          <p>Best regards,<br>ISDP - Security GRC and Data Privacy Team</p>
        </div>
      </body>
    </html>`;

  var chiefEmailList = Object.values(chiefEmailData);
  var ccEmail = ["email1@gmail.com","email2@gmail.com"];
  MailApp.sendEmail({
    to: chiefEmailList.join(','),
    cc: ccEmail.join(','),
    subject: '1st Week Update on ISDP Awareness - Q2 2024',
    htmlBody: emailBody
  });
  return "Chief Email sent successfully";
}
