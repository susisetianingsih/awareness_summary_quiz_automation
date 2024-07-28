function sendReminderEmailGRC() {
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

  // Send email
  var sendEmail = sendAutomaticEmailGRC(
    employeeNotCompleteQuiz,
    employeeCompleteQuiz,
    employeeInsufficientValue,
    employeeSufficientValue,
    halodocDepartementData
  );
  Logger.log(sendEmail); // Email info sent or not
}

function departementData(employeeData) {
  // Save departement dictionary
  var departementData = new Set();

  for (var i = 1; i < employeeData.length; i++) {
    const nameDepartemen = employeeData[i][2];

    // Add the department name to the Set
    departementData.add(nameDepartemen);
  }

  // Convert Set to Array
  return [...departementData];
}

function sendAutomaticEmailGRC(
  employeeNotCompleteQuiz,
  employeeCompleteQuiz,
  employeeInsufficientValue,
  employeeSufficientValue,
  halodocDepartementData
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
          .persentase {
            width: 50%;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h1>Dear Bapak Mohammad Nasir,</h1>
          <p>Here is a summary of the Cyber Security Knowledge Quiz results:</p>
          <h2>✍️ 
          Employe quiz complention status</h2>
          <table>
            <tr><th class="persentase">Completed</th><td class="persentase">${complete.toFixed(
              2
            )}%</td></tr>
            <tr><th class="persentase">Not Completed</th><td class="persentase">${notComplete.toFixed(
              2
            )}%</td></tr>
          </table>
          <h2>🌟 Employe quiz score status</h2>
          <table>
            <tr><th class="persentase">Sufficient</th><td class="persentase">${sufficient.toFixed(
              2
            )}%</td></tr>
            <tr><th class="persentase">Insufficient</th><td class="persentase">${insufficient.toFixed(
              2
            )}%</td></tr>
          </table>
          <p>Here is the breakdown by department:</p>
          <table>
            <tr><th>Department Name</th><th>Not Completed</th><th>Completed</th><th>Fail</th><th>Pass</th></tr>
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
    var employeSufficient =
      (departementSufficientData.length /
        (departementInsufficientData.length +
          departementSufficientData.length)) *
      100;

    // Write in the body of the email
    y += 1;
    emailBody += `<tr><td>${y}. ${
      halodocDepartementData[data]
    }</td><td>${employeeNotComplete.toFixed(
      2
    )}%</td><td>${employeeComplete.toFixed(
      2
    )}%</td><td>${employeeInsufficient.toFixed(
      2
    )}%</td><td>${employeSufficient.toFixed(2)}%</td></tr>`;
  }

  emailBody += `
          </table>
          <p>May the results of this quiz serve as a reference for improving the quality of employees' knowledge about Halodoc's cybersecurity.</p>
        </div>
      </body>
    </html>`;

  MailApp.sendEmail({
    to: "susisetia542@gmail.com",
    subject: "Cyber Security Knowledge Quiz Results [Edisi I]",
    htmlBody: emailBody,
  });
  return "Email sent successfully";
}
