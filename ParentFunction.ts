function halodocEmployeeData() {
  // Take a spreadsheet containing data on all employees
  var employeeSS = SpreadsheetApp.openByUrl(
    "https://docs.google.com/spreadsheets/employe"
  ); // link spreadsheet
  var employeeSheet = employeeSS.getSheetByName("test"); // Replace 'Sheet1' with the appropriate sheet name

  // Retrieve data from employee sheets
  var employeeData = employeeSheet.getDataRange().getValues();
  return employeeData;
}

function checkEmployeeQuiz(employeeData, quizData) {
  // Save employees who have not completed the Quiz
  var employeeNotCompleteQuiz = [];

  // Save employees who have completed the Quiz
  var employeeCompleteQuiz = [];

  // Quiz Value
  var employeeValue = {};

  // Email
  var employeeEmail = {};

  // Departement
  var employeeDepartement = {};

  // HoD
  var employeeHoD = {};

  // Chief
  var employeeChief = {};

  // Looping to match employee names with employee data and quiz result data
  for (var i = 0; i < employeeData.length; i++) {
    var found = false;
    var highestScore = 0;
    for (var j = 0; j < quizData.length; j++) {
      // If the data matches
      if (employeeData[i][1] == quizData[j][3]) {
        found = true;
        var currentScore = quizData[j][1];
        if (currentScore > highestScore) {
          highestScore = currentScore;
          employeeValue[employeeData[i][0]] = currentScore;
          employeeEmail[employeeData[i][0]] = employeeData[i][1];
          employeeDepartement[employeeData[i][0]] = employeeData[i][2];
          employeeHoD[employeeData[i][0]] = employeeData[i][3];
          employeeChief[employeeData[i][0]] = employeeData[i][5];
        }
      }
    }

    // If no data matches
    if (!found) {
      employeeNotCompleteQuiz.push({
        name: employeeData[i][0],
        email: employeeData[i][1],
        departement: employeeData[i][2],
        hod: employeeData[i][3],
        chief: employeeData[i][5],
      });
    }
  }

  // Save employees of sufficient value
  var employeeSufficientValue = [];

  // Save employees of insufficient value
  var employeeInsufficientValue = [];

  // Looping to find employees whose quiz scores are sufficient and insufficient (under 70)
  for (var karyawan in employeeValue) {
    employeeCompleteQuiz.push({
      name: karyawan,
      email: employeeEmail[karyawan],
      score: employeeValue[karyawan],
      departement: employeeDepartement[karyawan],
      hod: employeeHoD[karyawan],
      chief: employeeChief[karyawan],
    });

    if (employeeValue[karyawan] < 70) {
      employeeInsufficientValue.push({
        name: karyawan,
        email: employeeEmail[karyawan],
        score: employeeValue[karyawan],
        departement: employeeDepartement[karyawan],
        hod: employeeHoD[karyawan],
        chief: employeeChief[karyawan],
      });
    } else {
      employeeSufficientValue.push({
        name: karyawan,
        email: employeeEmail[karyawan],
        score: employeeValue[karyawan],
        departement: employeeDepartement[karyawan],
        hod: employeeHoD[karyawan],
        chief: employeeChief[karyawan],
      });
    }
  }

  result = {
    complete: employeeCompleteQuiz,
    "not complete": employeeNotCompleteQuiz,
    insufficient: employeeInsufficientValue,
    sufficient: employeeSufficientValue,
  };

  return result;
}
