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
