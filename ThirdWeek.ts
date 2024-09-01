function sendReminderEmailThirdWeek() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quizSheet = ss.getSheetByName("Form Responses 1"); // Replace 'Sheet1' with the appropriate sheet name

  // Data from quiz sheet
  var quizData = quizSheet.getDataRange().getValues();

  // Data from employee sheet
  var employeeData = halodocEmployeeData();

  // Data on employees who have not completed the quiz
  var employeeNotCompleteQuiz = checkEmployeeQuiz(employeeData, quizData)[
    "not complete"
  ];

  // Data on employees who have completed the quiz
  var employeeCompleteQuiz = checkEmployeeQuiz(employeeData, quizData)[
    "complete"
  ];

  // Data on employees whose quiz scores are insufficient
  var employeeInsufficientValue = checkEmployeeQuiz(employeeData, quizData)[
    "insufficient"
  ];

  // Data on employees whose quiz scores are sufficient
  var employeeSufficientValue = checkEmployeeQuiz(employeeData, quizData)[
    "sufficient"
  ];

  // Departement data
  var halodocDepartementData = departementData(employeeData);

  // Chief data
  var chiefEmailData = chiefData(employeeData);

  // Send email
  var sendEmail = sendAutomaticEmailGRC(
    employeeNotCompleteQuiz,
    employeeCompleteQuiz,
    employeeInsufficientValue,
    employeeSufficientValue,
    halodocDepartementData,
    chiefEmailData
  );
  Logger.log(sendEmail); // Email info sent or not
}

function sendAutomaticEmailGRC(
  employeeNotCompleteQuiz,
  employeeCompleteQuiz,
  employeeInsufficientValue,
  employeeSufficientValue,
  halodocDepartementData,
  chiefEmailData
) {
  // Calculate the percentage of employees who have not completed and completed the quiz
  var notComplete =
    (employeeNotCompleteQuiz.length /
      (employeeNotCompleteQuiz.length + employeeCompleteQuiz.length)) *
    100;
  var complete =
    (employeeCompleteQuiz.length /
      (employeeNotCompleteQuiz.length + employeeCompleteQuiz.length)) *
    100;

  // Calculate the percentage of employees whose quiz scores were sufficient and insufficient
  var insufficient =
    (employeeInsufficientValue.length /
      (employeeInsufficientValue.length + employeeSufficientValue.length)) *
    100;
  var sufficient =
    (employeeSufficientValue.length /
      (employeeInsufficientValue.length + employeeSufficientValue.length)) *
    100;

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
          <h1>Dear All,</h1>
          <p>With regards to the information security and data privacy awareness program Q2 2024 edition. Herewith, we would like to provide you with the final result of the completion status on the awareness program with the topic of “Social Engineering Attack” and “8 Principles of Data Protection” in each department, as follow:</p>
          <h1>Halosquads participation rate:</h1>
          <table>
            <tr><th class="persentase">Participated</th><td class="persentase">${complete.toFixed(
              2
            )}%</td></tr>
            <tr><th class="persentase">Not Participated</th><td class="persentase">${notComplete.toFixed(
              2
            )}%</td></tr>
          </table>
          <h1>Halosquads quiz status:</h1>
          <table>
            <tr><th class="persentase">Pass</th><td class="persentase">${sufficient.toFixed(
              2
            )}%</td></tr>
            <tr><th class="persentase">Failed</th><td class="persentase">${insufficient.toFixed(
              2
            )}%</td></tr>
          </table>
          <p>Here is the breakdown by department:</p>
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
      if (
        employeeNotCompleteQuiz[k]["departement"] ==
        halodocDepartementData[data]
      ) {
        departementNotCompleteData.push(employeeNotCompleteQuiz[k]);
      }
    }
    for (var k = 0; k < employeeCompleteQuiz.length; k++) {
      if (
        employeeCompleteQuiz[k]["departement"] == halodocDepartementData[data]
      ) {
        departementCompleteData.push(employeeCompleteQuiz[k]);
      }
    }
    for (var k = 0; k < employeeInsufficientValue.length; k++) {
      if (
        employeeInsufficientValue[k]["departement"] ==
        halodocDepartementData[data]
      ) {
        departementInsufficientData.push(employeeInsufficientValue[k]);
      }
    }
    for (var k = 0; k < employeeSufficientValue.length; k++) {
      if (
        employeeSufficientValue[k]["departement"] ==
        halodocDepartementData[data]
      ) {
        departementSufficientData.push(employeeSufficientValue[k]);
      }
    }
    // Calculate the percentage of employees for each department
    var employeeNotComplete =
      (departementNotCompleteData.length /
        (departementNotCompleteData.length + departementCompleteData.length)) *
      100;
    var employeeComplete =
      (departementCompleteData.length /
        (departementNotCompleteData.length + departementCompleteData.length)) *
      100;
    var employeeInsufficient =
      (departementInsufficientData.length /
        (departementInsufficientData.length +
          departementSufficientData.length)) *
      100;
    var employeeSufficient =
      (departementSufficientData.length /
        (departementInsufficientData.length +
          departementSufficientData.length)) *
      100;

    // Write in the body of the email
    y += 1;
    emailBody += (
      <tr>
        <td>
          ${y}. ${halodocDepartementData[data]}
        </td>
        <td>${employeeComplete.toFixed(2)}%</td>
        <td>${employeeNotComplete.toFixed(2)}%</td>
        <td>${employeeSufficient.toFixed(2)}%</td>
        <td>${employeeInsufficient.toFixed(2)}%</td>
      </tr>
    );
  }

  emailBody += `
          </table>
          <p>May the results could serve as a reference for improving the employees' knowledge about Halodoc's cybersecurity awareness.</p>
          <p>Thank you for your support in this program and see you again in the next awareness session!</p>
          <p>Best regards,<br>ISDP - Security GRC and Data Privacy Team</p>
        </div>
      </body>
    </html>`;

  var chiefEmailList = Object.values(chiefEmailData);
  var ccEmail = ["email1@gmail.com", "email2@gmail.com"];
  MailApp.sendEmail({
    to: chiefEmailList.join(","),
    cc: ccEmail.join(","),
    subject: "3rd Week Result - Q2 2024",
    htmlBody: emailBody,
  });
  return "Email sent successfully";
}
