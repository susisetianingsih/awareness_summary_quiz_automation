function sendReminderEmailChief() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName("response"); // Replace 'Sheet1' with the appropriate sheet name

  // Data from quiz sheet
  var quizData = quizSheet.getDataRange().getValues();

  // Data from employee sheet
  var employeeData = halodocEmployeeData();

  // Data on employees who have not completed the quiz
  var employeeNotCompleteQuiz = checkEmployeeQuiz(employeeData, quizData)[
    "not complete"
  ];

  // Data on employees whose quiz scores are insufficient
  var employeeInsufficientValue = checkEmployeeQuiz(employeeData, quizData)[
    "insufficient"
  ];

  // Chief data
  var chiefEmailData = chiefData(employeeData);

  // Send email
  var sendEmail = sendAutomaticEmailChief(
    employeeNotCompleteQuiz,
    employeeInsufficientValue,
    chiefEmailData
  );
  Logger.log(sendEmail); // Email info sent or not
}

function chiefData(employeeData) {
  // Save Chief dictionary
  var dataChief = {};

  for (var i = 1; i < employeeData.length; i++) {
    const nameChief = employeeData[i][5];
    const emailChief = employeeData[i][6];

    // If the name is not already in the Employee data object, add the Chief's name and email to the object
    if (!dataChief[nameChief]) {
      dataChief[nameChief] = emailChief;
    }
  }
  return dataChief;
}

function sendAutomaticEmailChief(
  employeeNotCompleteQuiz,
  employeeInsufficientValue,
  chiefEmailData
) {
  // Send email to Chief
  if (
    employeeNotCompleteQuiz.length > 0 ||
    employeeInsufficientValue.length > 0
  ) {
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
                background-color: #fff;
                border-radius: 5px;
                border: solid 1px #E0004D;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
              }
              h1 {
                font-size: 24px;
                margin-bottom: 20px;
              }
              h2 {
                font-size: 20px;
                margin-bottom: 20px;
              }
              p {
                color: #666;
                font-size: 16px;
                margin-bottom: 10px;
              }
              table {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 20px;
              }
              th, td {
                border: 1px solid #E0004D;
                padding: 8px;
                text-align: left;
              }
              th {
                background-color: rgba(128, 0, 128, 0.1);
              }
            </style>
          </head>
          <body>
            <div class="container">
              <h1>Dear Bapak/Ibu ${chief},</h1>
              <p>I hope this message finds you well. As part of our ongoing efforts to enhance cybersecurity awareness within our organization, I would like to provide you with an update on the completion status of the 'Edisi I: Cyber Security' quiz.</p>
      `;

      if (employeeNotCompleteQuiz.length > 0) {
        emailBody += `
          <h2>✍️ Employees who have not completed the quiz:</h2>
          <table>
          <tr><th>Name</th><th>Email</th><th>Departement</th></tr>
        `;

        var y = 0;
        for (var k = 0; k < employeeNotCompleteQuiz.length; k++) {
          if (employeeNotCompleteQuiz[k]["chief"].toLowerCase() == chief) {
            // Write in the body of the email
            y += 1;
            emailBody += `<tr><td>${y}. ${employeeNotCompleteQuiz[k]["name"]}</td><td>${employeeNotCompleteQuiz[k]["email"]}</td><td>${employeeNotCompleteQuiz[k]["departement"]}</td></tr>`;
          }
        }
        emailBody += `</table>`;
      }

      if (employeeInsufficientValue.length > 0) {
        emailBody += `
          <h2>🌟 Employees with scores below 70:</h2>
          <table>
          <tr><th>Name</th><th>Email</th><th>Score</th><th>Departement</th></tr>
        `;

        var y = 0;
        for (var k = 0; k < employeeInsufficientValue.length; k++) {
          if (employeeInsufficientValue[k]["chief"].toLowerCase() == chief) {
            // Write in the body of the email
            y += 1;
            emailBody += `<tr><td>${y}. ${employeeInsufficientValue[k]["name"]}</td><td>${employeeInsufficientValue[k]["email"]}</td><td>${employeeInsufficientValue[k]["score"]}</td><td>${employeeInsufficientValue[k]["departement"]}</td></tr>`;
          }
        }

        emailBody += `</table>`;
      }

      emailBody += `
              <p>Please take a moment to follow up with these employees to ensure completion and understanding of the material. Your support in this matter is greatly appreciated.</p>
              <p>Thank you for your attention to this important initiative.</p>
              <p>Best regards,<br>Tim ISDP GRC</p>
            </div>
          </body>
        </html>
      `;

      MailApp.sendEmail({
        to: chiefEmailData[chief],
        subject:
          "Reminder: Cybersecurity Quiz Completion and Scores - [Edisi I]",
        htmlBody: emailBody,
      });
    }
    return "Email sent successfully";
  } else {
    return "Email was not sent successfully";
  }
}
