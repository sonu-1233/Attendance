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
      newSuggestSalarySheet.setColumnWidth(3, 250);
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
      
      newSuggestSalarySheet.getRange(thirdRow, 3).setValue('Total Working Days In Month');
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
  
      let overtimeValue = parseFloat(row[overtimeIndex]);
      if (isNaN(overtimeValue)) {
          overtimeValue = 0;
      }
  
      var getFormalSheetData = [];
  
      for (var i = 1; i < data.length; i++) {
          var row = data[i];
          var obj = {};
          obj["Name"] = row[nameIndex];  
          obj["Total Attendances"] = row[attendanceIndex];  
          obj["Total Working Days"] = row[workingDaysIndex];  
          obj["Leaves Without Pay This Month"] = row[leavesIndex];  
          obj["Overtime (in days)"] = overtimeValue;
          getFormalSheetData.push(obj);
      }
  
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
        if(employeeAllData[i]['Overtime (in days)']) {
        var overtimeIndays = employeeAllData[i]['Overtime (in days)'];
        } else {
          var overtimeIndays = 0;
        }
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