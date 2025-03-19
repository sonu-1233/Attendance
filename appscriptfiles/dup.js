var RecordSheetId = "1d6w3DEI6H5TjtNWQ2bbjfOCl7c3keaXhwcYl5UA6yCw"; // sheet used for company records
var sheet = SpreadsheetApp.openById(RecordSheetId);

//-------------------salary  Suggestion sheet---------------
var SuggestSalarySheetId = "1izBDt1h1SlvX60hpXC1iZbRepT7YaFGQqN97dzQI_qI";
var suggestsalarysheet = SpreadsheetApp.openById(SuggestSalarySheetId);
// ---------------------------------------------------------

var KynetAttendanceSheetId = "1h5fuBwMqz8a09ESHaX5fH5bfuCzowx3kTOyqgoN7rR4";
var KynetAttendanceSheet = SpreadsheetApp.openById(KynetAttendanceSheetId);
var register = KynetAttendanceSheet.getSheetByName("register");
var daysThisMonth = KynetAttendanceSheet.getSheetByName("daysThisMonth");
var expectedHoursThisMonth = KynetAttendanceSheet.getSheetByName("ExpectedHoursThisMonth");

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
  console.log("in the create Suggestion sheet");
  var date = new Date();

  var currentMonthName = date.toLocaleString('default', { month: 'long' });
  console.log("current month name");
  console.log(currentMonthName);

  date.setMonth(date.getMonth());
  var options = { year: 'numeric', month: 'long' };
  console.log("date" + date);

  // var getMonthName = date.toLocaleDateString('en-US', options).replace(" ", " ");
  // var newSuggestSalarySheet = suggestsalarysheet.insertSheet(getMonthName, sheet.getSheets().length);
  var newSuggestSalarySheet = suggestsalarysheet.insertSheet( sheet.getSheets().length);
  console.log("newSuggestSalarySheet");
  // console.log(newSuggestSalarySheet);
 
  newSuggestSalarySheet.getRange('A50').setValue('last');
  newSuggestSalarySheet.getRange('M1').setValue('last');
  newSuggestSalarySheet.getRange('A1000').setValue('rows to be deleted');
  var lastRow = newSuggestSalarySheet.getLastRow();
  
  if (lastRow > 50) {
    newSuggestSalarySheet.deleteRows(51, lastRow - 50);
  }
  

  var totalColumns = newSuggestSalarySheet.getLastColumn();
  console.log("total column");
  console.log(totalColumns);

  newSuggestSalarySheet.getRange('A50').clear();
  newSuggestSalarySheet.getRange('M1').clear();
  console.log("total columns");
  console.log(totalColumns);

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
   secondRowRange.setValue('Salary Suggestion Sheet For The Month Of ');
   secondRowRange.setFontSize(16);
   newSuggestSalarySheet.setRowHeight(2, 45);
   secondRowRange.setBackground('#B7E1CD');
   secondRowRange.setHorizontalAlignment('left');

   newSuggestSalarySheet.setColumnWidth(1, 60);
   newSuggestSalarySheet.setColumnWidth(2, 320);
   newSuggestSalarySheet.setColumnWidth(3, 200);

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
    newSuggestSalarySheet.getRange(thirdRow, 2).setHorizontalAlignment('center');
    
    newSuggestSalarySheet.getRange(thirdRow, 3).setValue('Overtime In Days');
    newSuggestSalarySheet.getRange(thirdRow, 3).setFontSize(14);
    newSuggestSalarySheet.getRange(thirdRow, 3).setHorizontalAlignment('center');

    var getNamesFromDaysThisMonth = daysThisMonth.getRange(1, 1, 1, daysThisMonth.getLastColumn()).getValues();
    var namesArray = getNamesFromDaysThisMonth[0];
    console.log("namesArray");
    console.log(namesArray);

    var allSheetsFormal = sheet.getSheets();
    var lastSheetFormal = allSheetsFormal[allSheetsFormal.length - 1]; // Last sheet
    var lastSheetNameFormal = lastSheetFormal.getName();
    console.log(lastSheetNameFormal);


  }


}

function createNewSheetEveryMonth() {
  createSalarySuggestionSheet();
  console.log("function to create new sheet everymonth ");

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
  newSheet.getRange('AO1').setValue('last'); //modifiedline AM1 to AN1
  newSheet.getRange('A1000').setValue('rows to be deleted');
  var lastRow = newSheet.getLastRow();
  var lastColumn = newSheet.getLastColumn();
  console.log("lastColumn");
  console.log(lastColumn);

  if (lastRow > 50) {
    newSheet.deleteRows(51, lastRow - 50);
  }

  var totalColumns = newSheet.getLastColumn();
  console.log("total columns");
  console.log(totalColumns);
  newSheet.getRange('A50').clear();
  newSheet.getRange('AO1').clear();//modifiedline AM1 to AN1

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

   

    //newline
    var overtimeThisMonth = 8 + daysInMonth + 1;
    var overtimedoneonDates = 9 + daysInMonth + 1;

    console.log("paidLeavesThisYear");
    console.log(paidLeavesThisYear);
    console.log(overtimeThisMonth);

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

    //newline
    newSheet.getRange(thirdRow, overtimeThisMonth).setValue('Overtime (in days)');
    newSheet.getRange(thirdRow, overtimeThisMonth).setFontSize(15);
    newSheet.setColumnWidth(overtimeThisMonth, 200);

    newSheet.getRange(thirdRow, overtimedoneonDates).setValue('Overtime done on (dates)');
    newSheet.getRange(thirdRow, overtimedoneonDates).setFontSize(15);
    newSheet.setColumnWidth(overtimedoneonDates, 250);

   //modifiedline  paidLeavesThisYear with overtimedoneonDates

    if (lastColumn > overtimedoneonDates) {
      console.log("last column is greater than the paidleaves left this month");
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

        if(getHolidayDateInNumber == day) {
          var startRow = thirdRow + 1;
          if(holiday.name == "Deepavali") {
            if (dayOfWeek == 0) {
              newSheet.getRange(startRow, 3 + day, names.length, 1).mergeVertically();
              newSheet.getRange(startRow, 3 + day).setValue(holiday.name.toUpperCase());
              newSheet.getRange(startRow, 3 + day).setTextRotation(90);
              newSheet.getRange(startRow, 3 + day).setFontSize(25);
              newSheet.getRange(startRow, 3 + day).setFontColor("#FF0000");
              newSheet.getRange(startRow, 3 + day).setBackground("#F9E4DD");
              newSheet.getRange(startRow, 3 + day).setHorizontalAlignment("center");
              newSheet.getRange(startRow, 3 + day).setVerticalAlignment("middle");
          } else if(dayOfWeek === 6) {
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
    var counTotalHoliday  = 0;

    for(y = 0; y < rangeValues[0].length; y++) {
      if(rangeValues[0][y] !== '') {
      counTotalHoliday++;
      }
    }

    var totalHolidayInMonth = counTotalHoliday + specialCountForDiwaliHolidayCount;
   
    var totalWorkingDays = (daysInMonth - counTotalHoliday) - specialCountForDiwaliHolidayCount;
    if(totalWorkingDays) {
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
    if(currentMonthName == 'April') {
      getTotalAbsentThisMonthRange.clear();
      currentMonthPaidLeaveLeftThisYearColumn.setValue(paidLeaveForYear);
    }
  }
}

function markAttendance() {
  console.log("mark attendance fn");

  var sheets = sheet.getSheets();
  var activeSheet = sheets[sheets.length - 1];
  
  console.log(activeSheet);

  var names = daysThisMonth.getRange(1, 1, 1, daysThisMonth.getLastColumn()).getValues()[0];

  var dailyWorkingHours = daysThisMonth.getRange(3, 1, 1, daysThisMonth.getLastColumn()).getValues()[0]; //CHANGE THIS 6 TO 3

  var currentDay = dailyWorkingHours[0];

  console.log("current day");
  console.log(currentDay);

  var getDay = currentDay.getDay();
  var getDateInNumber = currentDay.getDate();
  console.log("get date in no");
  console.log(typeof(getDateInNumber));
  console.log(getDateInNumber);

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
  console.log("dates");
  console.log(dates);


  if (getDay != 0 && getDay != 6) {
    console.log("days from MOnday to Friday");
    for (var i = 0; i < dates.length; i++) {
      if (dates[i] == getDateInNumber) {
        var lastColumn = activeSheet.getLastColumn();

        console.log("lastColumn");
        console.log(lastColumn);
        console.log("check holiday");
        console.log(activeSheet.getRange(4, i + 3).getValue());

        const holidayName = activeSheet.getRange(4, i + 3).getValue();

        for (var j = 0; j < namesAndWorkingHours.length; j++) {
          console.log(j + 'and ' + i);
          var getHolidayValue = activeSheet.getRange(4 + j, i + 3).getValue();
          console.log("getHolidayValue");
          console.log(typeof(getHolidayValue));
          console.log(getHolidayValue);

          console.log("holiday name");
          console.log(typeof(holidayName));
          console.log(holidayName);

          if(getHolidayValue == '' && holidayName == "") {
            console.log("no holiday");
          var attendanceMark = '';
          if (namesAndWorkingHours[j].workingHours < 5.5 && namesAndWorkingHours[j].workingHours > 2 && namesAndWorkingHours[j].workingHours !== '') {
            attendanceMark = "H"
          } else if (namesAndWorkingHours[j].workingHours > 5.5) {
            attendanceMark = "P"

            console.log("hours greter than 5.5");
            console.log(namesAndWorkingHours[j].workingHours);

            //newline
            if(namesAndWorkingHours[j].workingHours > 8) {
              var overtime = namesAndWorkingHours[j].workingHours - 8;
              overtime = overtime.toFixed(1);
              var getOvertime = parseFloat(overtime);
              console.log("getOvertime");
               console.log(typeof(getOvertime));
              console.log(getOvertime);
              if(getOvertime > 0.5){
                console.log("overtime is greater than half an hour");
                var getovertimeInday = parseFloat((getOvertime / 8).toFixed(1));
                var overtimeonDatesLastValue = activeSheet.getRange(4 + j, lastColumn).getValue();
                console.log("getovertimeInday");
                console.log(getovertimeInday);
                var lastColumnValue = activeSheet.getRange(4 + j, lastColumn - 1).getValue();
                activeSheet.getRange(4 + j, lastColumn - 1).setValue(lastColumnValue + getovertimeInday);

                 if(overtimeonDatesLastValue == ''){
                    console.log("overtime value is none");
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

          var totalAbsentThisYear = activeSheet.getRange(4 + j, lastColumn - 4).getValue();   //modifiedline lastColumn - 2 to lastColumn - 3
          var leaveWithoutPay = activeSheet.getRange(4 + j, lastColumn - 3).getValue();   //modifiedline lastColumn - 1 to lastColumn - 2
          var totalPaidLeaveLeft = activeSheet.getRange(4 + j, lastColumn - 2).getValue();    //modifiedline lastColumn to lastColumn - 1

          
          console.log("totalAbsentThisMonthCount");
          console.log(totalAbsentThisMonthCount);
          console.log("totalAbsentThisYear" + totalAbsentThisYear);
          console.log("leaveWithoutPay" + leaveWithoutPay);
          console.log("totalPaidLeaveLeft" + totalPaidLeaveLeft);
          console.log(typeof(totalPaidLeaveLeft));

          // modified for mark A
          if (attendanceMark == "A") {
            console.log("absent marked");
            activeSheet.getRange(4 + j, i + 3).setBackground("#f68973");
            activeSheet.getRange(4 + j, lastColumn - 5).setValue(totalAbsentThisMonthCount + 1);
            if(totalPaidLeaveLeft >= 1) {
              console.log("paid leaves left is more than 1");
              activeSheet.getRange(4 + j, lastColumn - 4).setValue(totalAbsentThisYear + 1);
              activeSheet.getRange(4 + j, lastColumn - 2).setValue(totalPaidLeaveLeft - 1);
            } else if(totalPaidLeaveLeft < 1 && totalPaidLeaveLeft > 0) {
              console.log("paid leaves left is less than 1 and greater than 0");
              activeSheet.getRange(4 + j, lastColumn - 4).setValue(totalAbsentThisYear + 0.5);
              activeSheet.getRange(4 + j, lastColumn - 3).setValue(leaveWithoutPay + 0.5);
              activeSheet.getRange(4 + j, lastColumn - 2).setValue(totalPaidLeaveLeft - 0.5);
            } else {
              console.log("paid leaves left are 0");
              activeSheet.getRange(4 + j, lastColumn - 3).setValue(leaveWithoutPay + 1);
              activeSheet.getRange(4 + j, i + 3).setValue('LWP');
            }
          }

          var lastColumnCount = activeSheet.getRange(4 + j, lastColumn - 7).getValue();   //modifiedline lastColumn - 5 to lastColumn - 6
          if (attendanceMark == "P") {
            activeSheet.getRange(4 + j, lastColumn - 7).setValue(lastColumnCount + 1);    //modifiedline lastColumn - 5 to lastColumn - 6
          }
          else if (attendanceMark == "H") {
            console.log("half day");
            activeSheet.getRange(4 + j, lastColumn - 5).setValue(totalAbsentThisMonthCount + 0.5);
            activeSheet.getRange(4 + j, lastColumn - 7).setValue(lastColumnCount + 0.5);
            activeSheet.getRange(4 + j, i + 3).setBackground("#F4B599");
            if(totalPaidLeaveLeft > 0) {
              console.log("paid leave left is greater than 0");
              activeSheet.getRange(4 + j, lastColumn - 4).setValue(totalAbsentThisYear + 0.5);
              activeSheet.getRange(4 + j, lastColumn - 2).setValue(totalPaidLeaveLeft - 0.5);
            }else {
             activeSheet.getRange(4 + j, lastColumn - 3).setValue(leaveWithoutPay + 0.5);
            }
          }

        } else {
          //newline
          console.log('HOLIDAY');
          console.log(getHolidayValue);
          console.log(namesAndWorkingHours[j].workingHours);

            if(namesAndWorkingHours[j].workingHours > 2){
               console.log(namesAndWorkingHours[j].workingHours);
               console.log("ovetime");
               var lastColumnValue = activeSheet.getRange(4 + j, lastColumn - 1).getValue();
               var overtimeonDatesLastValue = activeSheet.getRange(4 + j, lastColumn).getValue();
               console.log("overtimeonDatesLastValue");
                 console.log(typeof(overtimeonDatesLastValue));
               console.log(overtimeonDatesLastValue);

               var getOvertimeInDay = parseFloat((namesAndWorkingHours[j].workingHours / 8).toFixed(1));
              //  getOvertimeInDay = Math.round(getOvertimeInDay);
              //  console.log("last column value");
               console.log(lastColumnValue);
               console.log("get overtime in day");
               console.log(getOvertimeInDay);
               activeSheet.getRange(4 + j, lastColumn - 1).setValue(lastColumnValue + getOvertimeInDay);

               if(overtimeonDatesLastValue == ''){
                console.log("overtime value is none");
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

  for(var i = 1; i <= daysInMonth; i++) {
    var currentDate = new Date(currentYear, currentMonth, i);
    var date = currentDate.getDate();
    holidayThisMonth.forEach((holiday) => {
      var holdayDate = new Date(holiday.date);
      var getHolidayDate = holiday.date.split('-')[2];
      var getHolidayDateInNumber = parseInt(getHolidayDate);
      var holidayDayOfWeek = holdayDate.getDay();
      if(date == getHolidayDateInNumber) {
        if(holidayDayOfWeek !== 0 && holidayDayOfWeek !== 6) {
          getHolidayOfThisMonth.push(holiday);
        }
      }
    });
  }

  var month = newdate.toLocaleString('default', { month: 'long' });
  var formattedDate = month + "," + currentYear;// December,2024

    // commented lines below are important
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

      if(totalColumns == lastFilledColumn) {
        // expectedHoursThisMonth.insertColumnAfter(totalColumns);
        totalColumns += 1;
        getLastColumn += 1;
      }

      var columnToFill = lastFilledColumn + 1;

      // expectedHoursThisMonth.getRange(1, columnToFill).setValue(formattedDate);
      // expectedHoursThisMonth.getRange(2, columnToFill).setValue(countSaturdaySunday);

      var holidayRangeOnExpectedHoursSheet = expectedHoursThisMonth.getRange(3, 1, 9, 1).getValues();

     if(getHolidayOfThisMonth.length > 0) {
      holidayRangeOnExpectedHoursSheet.forEach((getHoliday, index) => {
        getHolidayOfThisMonth.forEach((holidayThisMonth) => {
           if(getHoliday[0] == holidayThisMonth.name) {
            // expectedHoursThisMonth.getRange(2 + index + 1, columnToFill).setValue(1);
           }
        });
        });
     }
     var totalHours = totalWorkingDays * hoursPerDay;
    //  expectedHoursThisMonth.getRange(12, columnToFill).setValue(totalHolidayInMonth);
    //  expectedHoursThisMonth.getRange(13, columnToFill).setValue(daysInMonth);
    //  expectedHoursThisMonth.getRange(14, columnToFill).setValue(totalWorkingDays);
    //  expectedHoursThisMonth.getRange(15, columnToFill).setValue(hoursPerDay);
    //  expectedHoursThisMonth.getRange(16, columnToFill).setValue(totalHours);
    //  expectedHoursThisMonth.getRange(16, columnToFill).setBackground('#F4CCCC');
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

        if(holidayName !== 'Janmashtami (Smarta)') {
         trimmedHolidayName  = holidayName.includes('/') ? holidayName.split('/')[1]?.trim() : holidayName.replace(/\s*\(.*\)/, '').trim();
        } else {
         trimmedHolidayName  = holidayName;
        }
        if (kynetHolidayList[trimmedHolidayName] && holidayName !== 'Makar Sankranti') {
          return {
            name: trimmedHolidayName,
            date: holidayDate,
            numberOfDays: kynetHolidayList[trimmedHolidayName]
          };
        } else if(holidayName == 'Makar Sankranti') {
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