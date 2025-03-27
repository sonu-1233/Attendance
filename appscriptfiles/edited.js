var FormalRecordSheetId = "1hJ4Y-iOUejY5UWIYIU_d56oirQkK07LfONlMAdq0hN4"; // sheet used for company records
var sheet = SpreadsheetApp.openById(FormalRecordSheetId);

var KynetAttendanceSheetId = "1h5fuBwMqz8a09ESHaX5fH5bfuCzowx3kTOyqgoN7rR4";
var KynetAttendanceSheet = SpreadsheetApp.openById(KynetAttendanceSheetId);
var register = KynetAttendanceSheet.getSheetByName("register");
var daysThisMonth = KynetAttendanceSheet.getSheetByName("daysThisMonth");
var expectedHoursThisMonth = KynetAttendanceSheet.getSheetByName("ExpectedHoursThisMonth");


//------------------formal shee Id ------
var salarySheetId = "1t1YGjQ8hFreAzO6YzpFBOXXNSfoSE-bJjmcyFgiFgj8";
var getSalarySheet = SpreadsheetApp.openById(salarySheetId);


//-------------------salary  Suggestion sheet---------------
var suggestSalarySheetId = "1b9udamcYeQMp6dzyqwa9WJ6ciylnLobP2m2Nw_EqXqQ";
var suggestsalarysheet = SpreadsheetApp.openById(suggestSalarySheetId);

//---------------------------------------------------

function setTrigger() {  // UNCOMMENT THE TRIGGER IF U WANNA TO USE IT
  ScriptApp.newTrigger("onEditSalarySuggestionSheet")
    .forSpreadsheet('1b9udamcYeQMp6dzyqwa9WJ6ciylnLobP2m2Nw_EqXqQ')
    .onEdit()
    .create();
}

function toTitleCase(str) {
  return str.split(' ').map(word =>
    word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()
  ).join(' ');
}

var kynetHolidayList = {
  "Makar Sankranti"  :     '01',
  "Republic Day"     :	   '01',
  "Shivaratri"       :	   '01',
  "Holi"         	   :	   '01',
  "Independence Day" :	   '01',
  "Raksha Bandhan"   :	   '01',
  "Janmashtami"      :	   '01',
  "Dussehra"         :	   '01',
  "Deepavali"        :	   '02',
};

var holdiayForMarriedWomen = {
  "Karva Chauth" : '01',
};

function createSalarySuggestionSheet() {
  var sheets = getSalarySheet.getSheets();  
  var salarySheetLast = sheets[sheets.length - 1];
  var date = new Date();
  var year = date.getFullYear();
  var month = date.getMonth() + 1; 
  var totalDays = new Date(year, month, 0).getDate();
  date.setMonth(date.getMonth());
 
  var options = { year: 'numeric', month: 'long' };
  var getMonthName = date.toLocaleDateString('en-US', options).replace(" ", " ");
  var sheetName = "SalarySuggestion-" + getMonthName;
  var newSuggestSalarySheet = suggestsalarysheet.insertSheet(sheetName, sheet.getSheets().length);
 
  newSuggestSalarySheet.getRange('A50').setValue('last');
  newSuggestSalarySheet.getRange('J1').setValue('last');
  newSuggestSalarySheet.getRange('A1000').setValue('rows to be deleted');
  var lastRow = newSuggestSalarySheet.getLastRow();
  
  if (lastRow > 50) {
    newSuggestSalarySheet.deleteRows(51, lastRow - 50);
  }
  var totalColumns = newSuggestSalarySheet.getLastColumn();
  newSuggestSalarySheet.getRange('A50').clear();
  newSuggestSalarySheet.getRange('J1').clear();

  if (totalColumns > 0) {
    var firstRowRange = newSuggestSalarySheet.getRange(1, 1, 1, totalColumns);
    firstRowRange.merge();
    firstRowRange.setValue('Kynet Web Solutions Pvt. Ltd.');
    firstRowRange.setFontSize(24);
    newSuggestSalarySheet.setRowHeight(1, 50);
    firstRowRange.setBackground('#B7E1CD');
    firstRowRange.setHorizontalAlignment('center');

    var secondRowRange = newSuggestSalarySheet.getRange(2, 1, 1, totalColumns);
    secondRowRange.merge();
    secondRowRange.setValue('Salary Suggestion Sheet For The Month Of ' + getMonthName);
    secondRowRange.setFontSize(16);
    newSuggestSalarySheet.setRowHeight(2, 45);
    secondRowRange.setBackground('#B7E1CD');
    secondRowRange.setHorizontalAlignment('left');

    newSuggestSalarySheet.setColumnWidth(1, 60);
    newSuggestSalarySheet.setColumnWidth(2, 320);
    newSuggestSalarySheet.setColumnWidth(3, 200);
    newSuggestSalarySheet.setColumnWidth(4, 230);
    newSuggestSalarySheet.setColumnWidth(5, 280);
    newSuggestSalarySheet.setColumnWidth(6, 200);
    newSuggestSalarySheet.setColumnWidth(7, 250);
    newSuggestSalarySheet.setColumnWidth(8, 330);
    newSuggestSalarySheet.setColumnWidth(9, 200);
    newSuggestSalarySheet.setColumnWidth(10, 330);

    var lastRow = newSuggestSalarySheet.getMaxRows();
    for (var i = 3; i <= lastRow; i++) {
      newSuggestSalarySheet.setRowHeight(i, 40);
    }

    var thirdRow = 3;
    newSuggestSalarySheet.getRange(thirdRow, 1).setValue('S. No.');
    newSuggestSalarySheet.getRange(thirdRow, 1).setFontSize(13);
    newSuggestSalarySheet.getRange(thirdRow, 1).setHorizontalAlignment('right');

    newSuggestSalarySheet.getRange(thirdRow, 2).setValue('Name');
    newSuggestSalarySheet.getRange(thirdRow, 2).setFontSize(14);
    newSuggestSalarySheet.getRange(thirdRow, 2).setHorizontalAlignment('right');
    
    newSuggestSalarySheet.getRange(thirdRow, 3).setValue('Total Working Days');
    newSuggestSalarySheet.getRange(thirdRow, 3).setFontSize(14);
    newSuggestSalarySheet.getRange(thirdRow, 3).setHorizontalAlignment('right');

    newSuggestSalarySheet.getRange(thirdRow, 4).setValue('Total Salary Last Month');
    newSuggestSalarySheet.getRange(thirdRow, 4).setFontSize(14);
    newSuggestSalarySheet.getRange(thirdRow, 4).setHorizontalAlignment('right');

    newSuggestSalarySheet.getRange(thirdRow, 5).setValue('OverTime This Month (days)');
    newSuggestSalarySheet.getRange(thirdRow, 5).setFontSize(14);
    newSuggestSalarySheet.getRange(thirdRow, 5).setHorizontalAlignment('right');

    newSuggestSalarySheet.getRange(thirdRow, 6).setValue('OverTime Salary');
    newSuggestSalarySheet.getRange(thirdRow, 6).setFontSize(14);
    newSuggestSalarySheet.getRange(thirdRow, 6).setHorizontalAlignment('right');

    newSuggestSalarySheet.getRange(thirdRow, 7).setValue('Deduction Based on LWP');
    newSuggestSalarySheet.getRange(thirdRow, 7).setFontSize(14);
    newSuggestSalarySheet.getRange(thirdRow, 7).setHorizontalAlignment('right');

    newSuggestSalarySheet.getRange(thirdRow, 8).setValue('Leave WithOut Pay This Month');
    newSuggestSalarySheet.getRange(thirdRow, 8).setFontSize(14);
    newSuggestSalarySheet.getRange(thirdRow, 8).setHorizontalAlignment('right');

    newSuggestSalarySheet.getRange(thirdRow, 9).setValue('Suggest Total Salary');
    newSuggestSalarySheet.getRange(thirdRow, 9).setFontSize(14);
    newSuggestSalarySheet.getRange(thirdRow, 9).setHorizontalAlignment('right');

    newSuggestSalarySheet.getRange(thirdRow, 10).setValue('Salary To Be Credited This Month');
    newSuggestSalarySheet.getRange(thirdRow, 10).setFontSize(14).setFontWeight("bold");
    newSuggestSalarySheet.getRange(thirdRow, 10).setHorizontalAlignment('right');


    var getNamesFromDaysThisMonth = daysThisMonth.getRange(1, 1, 1, daysThisMonth.getLastColumn()).getValues();
    var namesArray = getNamesFromDaysThisMonth[0];
    var allSheetsFormal = sheet.getSheets();
    var secondLastSheetFormal = allSheetsFormal[allSheetsFormal.length - 2];

    var lastRow = secondLastSheetFormal.getLastRow();
    var lastColumn = secondLastSheetFormal.getLastColumn();
    var dataRange = secondLastSheetFormal.getRange(3, 1, lastRow - 2, lastColumn);
    var data = dataRange.getValues();

    var headers = data[0];
    var nameIndex = headers.indexOf("Name");
    var attendanceIndex = headers.indexOf("Total Attendances");
    var workingDaysIndex = headers.indexOf("Total Working Days");
    var leavesIndex = headers.indexOf("Leaves Without Pay This Month");
    var overtimeIndex = headers.indexOf("Overtime (in days)");

    var getFormalSheetData = [];

    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var obj = {};
        obj["Name"] = row[nameIndex];  
        obj["Total Attendances"] = row[attendanceIndex];  
        obj["Total Working Days"] = row[workingDaysIndex];  
        obj["Leaves Without Pay This Month"] = row[leavesIndex];  
        obj["Overtime (in days)"] = row[overtimeIndex];
        getFormalSheetData.push(obj);
    }
    console.log("---------------get the salary sheet data-------------");

    var lastRowSalarySheet = salarySheetLast.getLastRow();
    var lastColumnSalarySheet = salarySheetLast.getLastColumn();
    var dataRangeSalarySheet = salarySheetLast.getRange(3, 1, lastRowSalarySheet - 2, lastColumnSalarySheet);
    var salarySheetData = dataRangeSalarySheet.getValues(); 

    var headersSalarySheet = salarySheetData[0];
    var nameIndexSalary = headersSalarySheet.indexOf("Name");
    var deductionIndex = headersSalarySheet.indexOf("Deduction");
    var netCreditToBankIndex = headersSalarySheet.indexOf("Net Credit To Bank");
    var getSalarySheetData = [];

    for (var i = 1; i < salarySheetData.length; i++) {
        var row = salarySheetData[i];
        var obj = {};
        obj["Name"] = row[nameIndexSalary];  
        obj["Deduction"] = row[deductionIndex];  
        obj["Net Credit To Bank"] = row[netCreditToBankIndex];  
        getSalarySheetData.push(obj);
    }
    let employeeMap = {};
      getSalarySheetData.forEach(emp => {
          let normalizedName = emp.Name.trim().toLowerCase(); 
          employeeMap[normalizedName] = { ...emp };
      });

      getFormalSheetData.forEach(attendance => {
          let normalizedName = attendance.Name.trim().toLowerCase();
          if (employeeMap[normalizedName]) {
              Object.assign(employeeMap[normalizedName], attendance);
          } else {
              employeeMap[normalizedName] = { ...attendance };
          }
      });

      let employeeAllData = Object.values(employeeMap);
      var startRow = 4;
      var startCol = 1; 
      for (var i = 0; i < employeeAllData.length; i++) {
      var row = startRow + i;
      var netCreditLastMonth = employeeAllData[i]['Net Credit To Bank'];
      netCreditLastMonth = Number(netCreditLastMonth);
      var deduction = employeeAllData[i]['Deduction'];
      deduction = Number(deduction);

      var totalSalaryLastMonth = netCreditLastMonth + deduction;
      var overtimeIndays = employeeAllData[i]['Overtime (in days)'];
      overtimeIndays = Number(overtimeIndays);
      overtimeIndays = Math.floor(overtimeIndays);

      var oneDaySalaryCount = totalSalaryLastMonth / totalDays;

      var overtimeSalary = 0;
      if(overtimeIndays != '') {
        var overtimeSalary = oneDaySalaryCount * overtimeIndays;
        overtimeSalary = Math.floor(overtimeSalary)
      }
      var leavewithoutpay = employeeAllData[i]['Leaves Without Pay This Month'];
      leavewithoutpay = Number(leavewithoutpay);

      var suggestDeduction = 0;
      if(leavewithoutpay != "") {
      var suggestDeduction = oneDaySalaryCount * leavewithoutpay;
      suggestDeduction = Math.floor(suggestDeduction);
      }

      var suggestTotalSalary = (totalSalaryLastMonth + overtimeSalary) - suggestDeduction;
      newSuggestSalarySheet.getRange(row, startCol).setValue(i + 1);
      newSuggestSalarySheet.getRange(row, startCol + 1).setValue(employeeAllData[i].Name).setFontSize(13).setHorizontalAlignment("right");
      newSuggestSalarySheet.getRange(row, startCol + 2).setValue(employeeAllData[i]['Total Working Days']);
      newSuggestSalarySheet.getRange(row, startCol + 3).setValue(totalSalaryLastMonth);
      newSuggestSalarySheet.getRange(row, startCol + 4).setValue(overtimeIndays);
      newSuggestSalarySheet.getRange(row, startCol + 5).setValue(overtimeSalary);
      newSuggestSalarySheet.getRange(row, startCol + 6).setValue(suggestDeduction);
      newSuggestSalarySheet.getRange(row, startCol + 7).setValue(leavewithoutpay);
      newSuggestSalarySheet.getRange(row, startCol + 8).setValue(suggestTotalSalary);
    }
  }
}

function onEditSalarySuggestionSheet(e) {
  console.log("on edit sheet Triggered ");
  var getSuggestionsheet = e.source.getActiveSheet();
  var sheetName = getSuggestionsheet.getName();

  var thisRow = e.range.getRow();
  var thisCol = e.range.getColumn();
  var headingValue = getSuggestionsheet.getRange(3, thisCol).getValue();
 
  var monthYear = sheetName.split("-")[1];
  var monthName = monthYear.replace(/[0-9]/g, "");
  var year = parseInt(monthYear.replace(/\D/g, ""), 10);
  var date = new Date(Date.parse(monthName + " 1, " + year));
  var monthNumber = date.getMonth();
  var totalDays = new Date(year, monthNumber + 1, 0).getDate();

  if(headingValue == 'Salary To Be Credited This Month'){
    var totalsalaryToBeCredited = getSuggestionsheet.getRange(thisRow, thisCol).getValue();
    var getOverTimeDays = getSuggestionsheet.getRange(thisRow, 5).getValue();
    getOverTimeDays = Number(getOverTimeDays);
    getOverTimeDays = Math.floor(getOverTimeDays);

    var oneDaySalaryCount = totalsalaryToBeCredited / totalDays;
    oneDaySalaryCount = Math.floor(oneDaySalaryCount);
    var leaveWithoutPayThisMonth = getSuggestionsheet.getRange(thisRow, 8).getValue();
    var deductionBasedOnLWP = leaveWithoutPayThisMonth * oneDaySalaryCount;
    deductionBasedOnLWP = Number(deductionBasedOnLWP);
    var overtimeSalary = 0;

    if(getOverTimeDays != '') {
      var overtimeSalary = oneDaySalaryCount * getOverTimeDays;
      overtimeSalary = Math.floor(overtimeSalary)
    }
    var suggestTotalSalary = (totalsalaryToBeCredited - deductionBasedOnLWP) + overtimeSalary;
    getSuggestionsheet.getRange(thisRow, 6).setValue(overtimeSalary);
    getSuggestionsheet.getRange(thisRow, 7).setValue(deductionBasedOnLWP);
    getSuggestionsheet.getRange(thisRow, 9).setValue(suggestTotalSalary);

    console.log("SALARY SUGGESTION COMPLETED");
  }
}

function createNewSheetEveryMonth() {
  var date = new Date();
  var currentMonthName = date.toLocaleString('default', { month: 'long' });
  const holidayThisMonth = getHolidayThisMonth(date);
  console.log(holidayThisMonth);
  date.setMonth(date.getMonth());
  var options = { year: 'numeric', month: 'long' };
  var getMonthName = date.toLocaleDateString('en-US', options).replace(" ", " ");
  var newSheet = sheet.insertSheet(getMonthName, sheet.getSheets().length);

  const allSheets = sheet.getSheets();
  newSheet.getRange('A50').setValue('last');
  newSheet.getRange('AO1').setValue('last');
  newSheet.getRange('A1000').setValue('rows to be deleted');
  var lastRow = newSheet.getLastRow();
  var lastColumn = newSheet.getLastColumn();

  if (lastRow > 50) {
    newSheet.deleteRows(51, lastRow - 50);
  }

  var totalColumns = newSheet.getLastColumn();
  newSheet.getRange('A50').clear();
  newSheet.getRange('AO1').clear();

  var currentYear = date.getFullYear();
  var currentMonth = date.getMonth();

  if (totalColumns > 0) {
    var firstRowRange = newSheet.getRange(1, 1, 1, totalColumns);
    firstRowRange.merge();

    firstRowRange.setValue('Kynet Web Solutions Pvt. Ltd.');
    firstRowRange.setFontSize(24);
    newSheet.setRowHeight(1, 50);
    firstRowRange.setBackground('#B7E1CD');
    firstRowRange.setHorizontalAlignment('center');

    var secondRowRange = newSheet.getRange(2, 1, 1, totalColumns);
    secondRowRange.merge();
    secondRowRange.setValue('Daily Attendance Register For The Month Of ' + getMonthName);
    secondRowRange.setFontSize(16);
    newSheet.setRowHeight(2, 45);
    secondRowRange.setBackground('#B7E1CD');
    secondRowRange.setHorizontalAlignment('left');

    newSheet.setColumnWidth(1, 60);
    newSheet.setColumnWidth(2, 320);

    var lastRow = newSheet.getMaxRows();
    for (var i = 3; i <= lastRow; i++) {
      newSheet.setRowHeight(i, 40);
    }

    var thirdRow = 3;
    newSheet.getRange(thirdRow, 1).setValue('S. No.');
    newSheet.getRange(thirdRow, 1).setFontSize(13);
    newSheet.getRange(thirdRow, 1).setHorizontalAlignment('right');

    newSheet.getRange(thirdRow, 2).setValue('Name');
    newSheet.getRange(thirdRow, 2).setFontSize(14);
    newSheet.getRange(thirdRow, 2).setHorizontalAlignment('center');

    var daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
    var getNamesFromDaysThisMonth = daysThisMonth.getRange(1, 1, 1, daysThisMonth.getLastColumn()).getValues();
    var namesArray = getNamesFromDaysThisMonth[0];
    var totalAttendancesColumn = 2 + daysInMonth + 1;
    var totalWorkingDaysColumn = 3 + daysInMonth + 1;
    var totalAbsentThisMonthColumn = 4 + daysInMonth + 1;
    var totalAbsetThisYearColumn = 5 + daysInMonth + 1;
    var leavesWithoutPayThisMonth = 6 + daysInMonth + 1;
    var paidLeavesThisYear = 7 + daysInMonth + 1;
    var overtimeThisMonth = 8 + daysInMonth + 1;
    var overtimedoneonDates = 9 + daysInMonth + 1;

    newSheet.getRange(thirdRow, totalAttendancesColumn).setValue('Total Attendances');
    newSheet.getRange(thirdRow, totalAttendancesColumn).setFontSize(15);
    newSheet.setColumnWidth(totalAttendancesColumn, 175);

    newSheet.getRange(thirdRow, totalWorkingDaysColumn).setValue('Total Working Days');
    newSheet.getRange(thirdRow, totalWorkingDaysColumn).setFontSize(15);
    newSheet.setColumnWidth(totalWorkingDaysColumn, 195);

    newSheet.getRange(thirdRow, totalAbsentThisMonthColumn).setValue('Total Absents This Month');
    newSheet.getRange(thirdRow, totalAbsentThisMonthColumn).setFontSize(15);
    newSheet.setColumnWidth(totalAbsentThisMonthColumn, 245);

    newSheet.getRange(thirdRow, totalAbsetThisYearColumn).setValue('Total Absents This Year');
    newSheet.getRange(thirdRow, totalAbsetThisYearColumn).setFontSize(15);
    newSheet.setColumnWidth(totalAbsetThisYearColumn, 220);

    newSheet.getRange(thirdRow, leavesWithoutPayThisMonth).setValue('Leaves Without Pay This Month');
    newSheet.getRange(thirdRow, leavesWithoutPayThisMonth).setFontSize(15);
    newSheet.setColumnWidth(leavesWithoutPayThisMonth, 300);

    newSheet.getRange(thirdRow, paidLeavesThisYear).setValue('Paid Leaves Left This Year');
    newSheet.getRange(thirdRow, paidLeavesThisYear).setFontSize(15);
    newSheet.setColumnWidth(paidLeavesThisYear, 250);

    newSheet.getRange(thirdRow, overtimeThisMonth).setValue('Overtime (in days)');
    newSheet.getRange(thirdRow, overtimeThisMonth).setFontSize(15);
    newSheet.setColumnWidth(overtimeThisMonth, 200);

    newSheet.getRange(thirdRow, overtimedoneonDates).setValue('Overtime done on (dates)');
    newSheet.getRange(thirdRow, overtimedoneonDates).setFontSize(15);
    newSheet.setColumnWidth(overtimedoneonDates, 250);

    if (lastColumn > overtimedoneonDates) {
      var columnsToDelete = lastColumn - overtimedoneonDates;
      newSheet.deleteColumns(overtimedoneonDates + 1, columnsToDelete);
    }

    var names = [];
    for (var i = 0; i < namesArray.length; i++) {
      if (namesArray[i] !== '' && namesArray[i] !== ' ') {
        names.push(namesArray[i]);
      }
    }

    var startRow = 4;
    var startColumn = 2;
    var nextStartColumn = 1;

    names.forEach((name, index) => {
      const formattedName = toTitleCase(name);
      newSheet.getRange(startRow + index, nextStartColumn).setValue(index + 1);

      const serialNumberCell = newSheet.getRange(startRow + index, nextStartColumn);
      serialNumberCell.setValue(index + 1);
      serialNumberCell.setFontSize(10);
      serialNumberCell.setHorizontalAlignment("right");

      newSheet.getRange(startRow + index, startColumn).setValue(formattedName);

      const nameCell = newSheet.getRange(startRow + index, startColumn);
      nameCell.setValue(formattedName);
      nameCell.setFontSize(14);
      nameCell.setHorizontalAlignment("right");
    });

    var countSaturdaySunday = 0;
    var specialCountForDiwaliHolidayCount = 0;

    for (var day = 1; day <= daysInMonth; day++) {
      var currentDate = new Date(currentYear, currentMonth, day);
      var dayOfWeek = currentDate.getDay();

      newSheet.getRange(thirdRow, 2 + day).setValue(day);
      newSheet.getRange(thirdRow, 2 + day).setFontSize(10);
      newSheet.setColumnWidth(2 + day, 40);
      holidayThisMonth.forEach((holiday) => {
        var getHolidayDate = holiday.date.split('-')[2];
        var getHolidayDateInNumber = parseInt(getHolidayDate);

        if (getHolidayDateInNumber == day) {
          var startRow = thirdRow + 1;
          if (holiday.name == "Deepavali") {
            if (dayOfWeek == 0) {
              newSheet.getRange(startRow, 3 + day, names.length, 1).mergeVertically();
              newSheet.getRange(startRow, 3 + day).setValue(holiday.name.toUpperCase());
              newSheet.getRange(startRow, 3 + day).setTextRotation(90);
              newSheet.getRange(startRow, 3 + day).setFontSize(25);
              newSheet.getRange(startRow, 3 + day).setFontColor("#FF0000");
              newSheet.getRange(startRow, 3 + day).setBackground("#F9E4DD");
              newSheet.getRange(startRow, 3 + day).setHorizontalAlignment("center");
              newSheet.getRange(startRow, 3 + day).setVerticalAlignment("middle");
            } else if (dayOfWeek === 6) {
              console.log("SATURDAY");
            } else {
              newSheet.getRange(startRow, 2 + day, names.length, 1).merge();
              newSheet.getRange(startRow, 3 + day, names.length, 1).merge();
              var rangeToMerge = newSheet.getRange(startRow, 2 + day, names.length, 2);
              rangeToMerge.merge();
              newSheet.getRange(startRow, 2 + day).setValue(holiday.name.toUpperCase());
              newSheet.getRange(startRow, 2 + day).setTextRotation(90);
              specialCountForDiwaliHolidayCount = 1;
            }
          } else {
            var startRow = thirdRow + 1;
            newSheet.getRange(startRow, 2 + day, names.length, 1).mergeVertically();
            newSheet.getRange(startRow, 2 + day).setValue(holiday.name.toUpperCase());
            newSheet.getRange(startRow, 2 + day).setTextRotation(90);
          }
          var sundayRange = newSheet.getRange(startRow, 2 + day);
          sundayRange.setFontSize(25);
          sundayRange.setFontColor("#FF0000");
          sundayRange.setBackground("#F9E4DD");
          sundayRange.setHorizontalAlignment("center");
          sundayRange.setVerticalAlignment("middle");
        }
      });

      if (dayOfWeek === 0) {
        var startRow = thirdRow + 1;
        newSheet.getRange(startRow, 2 + day, names.length, 1).mergeVertically();
        newSheet.getRange(startRow, 2 + day).setValue("SUNDAY");
        newSheet.getRange(startRow, 2 + day).setTextRotation(90);
        var sundayRange = newSheet.getRange(startRow, 2 + day);
        sundayRange.setFontSize(25);
        sundayRange.setFontColor("#FF0000");
        sundayRange.setBackground("#F9E4DD");
        sundayRange.setHorizontalAlignment("center");
        sundayRange.setVerticalAlignment("middle");
        countSaturdaySunday++;
      } else if (dayOfWeek === 6) {
        var startRow = thirdRow + 1;
        newSheet.getRange(startRow, 2 + day, names.length, 1).mergeVertically();
        var saturdayRange = newSheet.getRange(startRow, 2 + day);
        saturdayRange.setValue("SATURDAY");
        saturdayRange.setTextRotation(90);
        saturdayRange.setFontSize(25);
        saturdayRange.setFontColor("#FF0000");
        saturdayRange.setBackground("#F9E4DD");
        saturdayRange.setHorizontalAlignment("center");
        saturdayRange.setVerticalAlignment("middle");
        countSaturdaySunday++;
      }
    }

    var getFourthcolumn = newSheet.getRange('4:4');
    var getRangeToCalculateHoliday = getFourthcolumn.getDisplayValues();
    var startColumn = 3;
    var endColumn = daysInMonth + startColumn - 1;

    var getRangeToCalculateHoliday = newSheet.getRange(4, startColumn, 1, endColumn - startColumn + 1);
    var rangeValues = getRangeToCalculateHoliday.getDisplayValues();
    var counTotalHoliday = 0;

    for (y = 0; y < rangeValues[0].length; y++) {
      if (rangeValues[0][y] !== '') {
        counTotalHoliday++;
      }
    }

    var totalHolidayInMonth = counTotalHoliday + specialCountForDiwaliHolidayCount;

    var totalWorkingDays = (daysInMonth - counTotalHoliday) - specialCountForDiwaliHolidayCount;
    if (totalWorkingDays) {
      calculateTotalWorkingHourMonth(totalWorkingDays, totalHolidayInMonth, countSaturdaySunday, holidayThisMonth, daysInMonth, date);
    }

    var numberOfNames = names.length;
    var startRow = thirdRow + 1;
    var startColumn = totalWorkingDaysColumn;
    var valuesToWrite = new Array(numberOfNames).fill([totalWorkingDays]);
    newSheet.getRange(startRow, startColumn, numberOfNames, 1).setValues(valuesToWrite);

    var secondLastSheet = allSheets[allSheets.length - 2];
    var totalWorkingDaysRow = 3;
    var data = secondLastSheet.getRange(totalWorkingDaysRow, 1, 1, secondLastSheet.getLastColumn()).getValues()[0];
    var totalWorkingDaysColumnIndex = data.indexOf("Total Working Days") + 1;
    var totalAbsentThisYearColumnIndex = totalWorkingDaysColumnIndex + 2;
    var totalPaidLeaveLeftThisYearColumnIndex = totalWorkingDaysColumnIndex + 4;

    var totalAbsetLastMonthRange = secondLastSheet.getRange(totalWorkingDaysRow + 1, totalAbsentThisYearColumnIndex, numberOfNames, 1);
    var getTotalAbsentLastMonth = totalAbsetLastMonthRange.getValues();

    var getTotalAbsentThisMonthRange = newSheet.getRange(totalWorkingDaysRow + 1, totalAbsetThisYearColumn, numberOfNames, 1);
    getTotalAbsentThisMonthRange.setValues(getTotalAbsentLastMonth);

    var totalPaidLeaveLeftThisYearRange = secondLastSheet.getRange(totalWorkingDaysRow + 1, totalPaidLeaveLeftThisYearColumnIndex, numberOfNames, 1);
    var getTotalPaidLeaveLeftThisYearValue = totalPaidLeaveLeftThisYearRange.getValues();

    var currentMonthPaidLeaveLeftThisYearColumn = newSheet.getRange(totalWorkingDaysRow + 1, paidLeavesThisYear, numberOfNames, 1);
    currentMonthPaidLeaveLeftThisYearColumn.setValues(getTotalPaidLeaveLeftThisYearValue);

    var paidLeaveForYear = 16;
    if (currentMonthName == 'April') {
      getTotalAbsentThisMonthRange.clear();
      currentMonthPaidLeaveLeftThisYearColumn.setValue(paidLeaveForYear);
    }
  }
}

function markAttendance() {
  var sheets = sheet.getSheets();
  var activeSheet = sheets[sheets.length - 1];

  var names = daysThisMonth.getRange(1, 1, 1, daysThisMonth.getLastColumn()).getValues()[0];
  var dailyWorkingHours = daysThisMonth.getRange(3, 1, 1, daysThisMonth.getLastColumn()).getValues()[0];

  var currentDay = dailyWorkingHours[0];
  console.log(currentDay);
  var getDay = currentDay.getDay();
  var getDateInNumber = currentDay.getDate();

  var namesAndWorkingHours = [];
  for (var i = 1; i < names.length; i++) {
    if (names[i] !== '' && names[i] !== ' ' || dailyWorkingHours[i] !== '' && dailyWorkingHours[i] !== ' ') {
      namesAndWorkingHours.push({
        name: names[i],
        workingHours: dailyWorkingHours[i]
      });
    }
  }
  var dates = activeSheet.getRange(3, 3, 1, activeSheet.getLastColumn() - 8).getValues()[0];


  if (getDay != 0 && getDay != 6) {
    console.log("days from MOnday to Friday");
    for (var i = 0; i < dates.length; i++) {
      if (dates[i] == getDateInNumber) {
        var lastColumn = activeSheet.getLastColumn();
        const holidayName = activeSheet.getRange(4, i + 3).getValue();
        for (var j = 0; j < namesAndWorkingHours.length; j++) {
          var getHolidayValue = activeSheet.getRange(4 + j, i + 3).getValue();
          if (getHolidayValue == '' && holidayName == "") {
            var attendanceMark = '';
            if (namesAndWorkingHours[j].workingHours < 5.5 && namesAndWorkingHours[j].workingHours > 2 && namesAndWorkingHours[j].workingHours !== '') {
              attendanceMark = "H"
            } else if (namesAndWorkingHours[j].workingHours > 5.5) {
              attendanceMark = "P"
              if (namesAndWorkingHours[j].workingHours > 8) {
                var overtime = namesAndWorkingHours[j].workingHours - 8;
                overtime = overtime.toFixed(1);
                var getOvertime = parseFloat(overtime);
                if (getOvertime > 0.5) {
                  var getovertimeInday = parseFloat((getOvertime / 8).toFixed(1));
                  var overtimeonDatesLastValue = activeSheet.getRange(4 + j, lastColumn).getValue();
                  var lastColumnValue = activeSheet.getRange(4 + j, lastColumn - 1).getValue();
                  activeSheet.getRange(4 + j, lastColumn - 1).setValue(lastColumnValue + getovertimeInday);
                  if (overtimeonDatesLastValue == '') {
                    activeSheet.getRange(4 + j, lastColumn).setValue(getDateInNumber).setHorizontalAlignment("right");
                  } else {
                    overtimeonDatesLastValue = overtimeonDatesLastValue + ',' + getDateInNumber;
                    activeSheet.getRange(4 + j, lastColumn).setValue(overtimeonDatesLastValue).setHorizontalAlignment("right");
                  }
                }
              }
            } else {
              attendanceMark = 'A';
            }
            activeSheet.getRange(4 + j, i + 3).setValue(attendanceMark);
            var totalAbsentThisMonthCount = activeSheet.getRange(4 + j, lastColumn - 5).getValue();

            var totalAbsentThisYear = activeSheet.getRange(4 + j, lastColumn - 4).getValue();
            var leaveWithoutPay = activeSheet.getRange(4 + j, lastColumn - 3).getValue();
            var totalPaidLeaveLeft = activeSheet.getRange(4 + j, lastColumn - 2).getValue();
            if (attendanceMark == "A") {
              activeSheet.getRange(4 + j, i + 3).setBackground("#f68973");
              activeSheet.getRange(4 + j, lastColumn - 5).setValue(totalAbsentThisMonthCount + 1);
              if (totalPaidLeaveLeft >= 1) {
                activeSheet.getRange(4 + j, lastColumn - 4).setValue(totalAbsentThisYear + 1);
                activeSheet.getRange(4 + j, lastColumn - 2).setValue(totalPaidLeaveLeft - 1);
              } else if (totalPaidLeaveLeft < 1 && totalPaidLeaveLeft > 0) {
                activeSheet.getRange(4 + j, lastColumn - 4).setValue(totalAbsentThisYear + 0.5);
                activeSheet.getRange(4 + j, lastColumn - 3).setValue(leaveWithoutPay + 0.5);
                activeSheet.getRange(4 + j, lastColumn - 2).setValue(totalPaidLeaveLeft - 0.5);
              } else {
                activeSheet.getRange(4 + j, lastColumn - 3).setValue(leaveWithoutPay + 1);
                activeSheet.getRange(4 + j, i + 3).setValue('LWP');
              }
            }

            var lastColumnCount = activeSheet.getRange(4 + j, lastColumn - 7).getValue();
            if (attendanceMark == "P") {
              activeSheet.getRange(4 + j, lastColumn - 7).setValue(lastColumnCount + 1);
            } else if (attendanceMark == "H") {
              activeSheet.getRange(4 + j, lastColumn - 5).setValue(totalAbsentThisMonthCount + 0.5);
              activeSheet.getRange(4 + j, lastColumn - 7).setValue(lastColumnCount + 0.5);
              activeSheet.getRange(4 + j, i + 3).setBackground("#F4B599");
              if (totalPaidLeaveLeft > 0) {
                activeSheet.getRange(4 + j, lastColumn - 4).setValue(totalAbsentThisYear + 0.5);
                activeSheet.getRange(4 + j, lastColumn - 2).setValue(totalPaidLeaveLeft - 0.5);
              } else {
                activeSheet.getRange(4 + j, lastColumn - 3).setValue(leaveWithoutPay + 0.5);
              }
            }
          } else {
            if (namesAndWorkingHours[j].workingHours > 2) {
              var lastColumnValue = activeSheet.getRange(4 + j, lastColumn - 1).getValue();
              var overtimeonDatesLastValue = activeSheet.getRange(4 + j, lastColumn).getValue();
              var getOvertimeInDay = parseFloat((namesAndWorkingHours[j].workingHours / 8).toFixed(1));
              activeSheet.getRange(4 + j, lastColumn - 1).setValue(lastColumnValue + getOvertimeInDay);

              if (overtimeonDatesLastValue == '') {
                activeSheet.getRange(4 + j, lastColumn).setValue(getDateInNumber).setHorizontalAlignment("right");
              } else {
                overtimeonDatesLastValue = overtimeonDatesLastValue + ',' + getDateInNumber;
                activeSheet.getRange(4 + j, lastColumn).setValue(overtimeonDatesLastValue).setHorizontalAlignment("right");
              }
            }
          }
        }
      }
    }
  } else {
    console.log("Sunday and Saturday");
  }
}

function calculateTotalWorkingHourMonth(totalWorkingDays, totalHolidayInMonth, countSaturdaySunday, holidayThisMonth, daysInMonth, date) {
  var hoursPerDay = 7.75;
  var newdate = new Date(date);

  var currentYear = date.getFullYear();
  var currentMonth = date.getMonth();

  var getHolidayOfThisMonth = [];

  for (var i = 1; i <= daysInMonth; i++) {
    var currentDate = new Date(currentYear, currentMonth, i);
    var date = currentDate.getDate();
    holidayThisMonth.forEach((holiday) => {
      var holdayDate = new Date(holiday.date);
      var getHolidayDate = holiday.date.split('-')[2];
      var getHolidayDateInNumber = parseInt(getHolidayDate);
      var holidayDayOfWeek = holdayDate.getDay();
      if (date == getHolidayDateInNumber) {
        if (holidayDayOfWeek !== 0 && holidayDayOfWeek !== 6) {
          getHolidayOfThisMonth.push(holiday);
        }
      }
    });
  }

  var month = newdate.toLocaleString('default', { month: 'long' });
  var formattedDate = month + "," + currentYear;


  var totalColumns = expectedHoursThisMonth.getMaxColumns();
  var getLastColumn = expectedHoursThisMonth.getRange('1:1').getLastColumn();

  var firstRowValues = expectedHoursThisMonth.getRange(1, 1, 1, totalColumns).getValues()[0];
  var lastFilledColumn = -1;

  for (var col = totalColumns - 1; col >= 0; col--) {
    if (firstRowValues[col] !== "") {
      lastFilledColumn = col + 1;
      break;
    }
  }

  if (totalColumns == lastFilledColumn) {
    expectedHoursThisMonth.insertColumnAfter(totalColumns);
    totalColumns += 1;
    getLastColumn += 1;
  }

  var columnToFill = lastFilledColumn + 1;

  expectedHoursThisMonth.getRange(1, columnToFill).setValue(formattedDate);
  expectedHoursThisMonth.getRange(2, columnToFill).setValue(countSaturdaySunday);

  var holidayRangeOnExpectedHoursSheet = expectedHoursThisMonth.getRange(3, 1, 9, 1).getValues();

  if (getHolidayOfThisMonth.length > 0) {
    holidayRangeOnExpectedHoursSheet.forEach((getHoliday, index) => {
      getHolidayOfThisMonth.forEach((holidayThisMonth) => {
        if (getHoliday[0] == holidayThisMonth.name) {
          expectedHoursThisMonth.getRange(2 + index + 1, columnToFill).setValue(1);
        }
      });
    });
  }
  var totalHours = totalWorkingDays * hoursPerDay;
  expectedHoursThisMonth.getRange(12, columnToFill).setValue(totalHolidayInMonth);
  expectedHoursThisMonth.getRange(13, columnToFill).setValue(daysInMonth);
  expectedHoursThisMonth.getRange(14, columnToFill).setValue(totalWorkingDays);
  expectedHoursThisMonth.getRange(15, columnToFill).setValue(hoursPerDay);
  expectedHoursThisMonth.getRange(16, columnToFill).setValue(totalHours);
  expectedHoursThisMonth.getRange(16, columnToFill).setBackground('#F4CCCC');
}

function getHolidayThisMonth(currentMonthDate) {
  const calendarId = 'en-gb.indian#holiday@group.v.calendar.google.com';
  const startDate = new Date(currentMonthDate);
  startDate.setDate(1);

  const endDate = new Date(startDate);
  endDate.setMonth(startDate.getMonth() + 1);
  endDate.setDate(0);

  try {
    const events = Calendar.Events.list(calendarId, {
      timeMin: startDate.toISOString(),
      timeMax: endDate.toISOString(),
      singleEvents: true,
      orderBy: 'startTime'
    });

    if (events.items && events.items.length > 0) {
      const holidays = events.items.map(event => {
        const holidayName = event.summary;
        const holidayDate = event.start.date || event.start.dateTime;

        let trimmedHolidayName;

        if (holidayName !== 'Janmashtami (Smarta)') {
          trimmedHolidayName = holidayName.includes('/') ? holidayName.split('/')[1]?.trim() : holidayName.replace(/\s*\(.*\)/, '').trim();
        } else {
          trimmedHolidayName = holidayName;
        }
        if (kynetHolidayList[trimmedHolidayName] && holidayName !== 'Makar Sankranti') {
          return {
            name: trimmedHolidayName,
            date: holidayDate,
            numberOfDays: kynetHolidayList[trimmedHolidayName]
          };
        } else if (holidayName == 'Makar Sankranti') {
          const dateObj = new Date(holidayDate);
          dateObj.setDate(dateObj.getDate());
          const formattedDate = dateObj.toISOString().split('T')[0];

          var lohriHoliday = "Lohri";
          return {
            name: lohriHoliday,
            date: formattedDate,
            numberOfDays: kynetHolidayList[trimmedHolidayName]
          };
        }
      }).filter(holiday => holiday !== undefined);

      return holidays;
    } else {
      console.log('No holidays found for the current month.');
    }
  } catch (error) {
    console.log('Error fetching holidays: ' + error.message);
  }
}