function createSalarySheet() {
    var allSuggestSalarySheets = suggestsalarysheet.getSheets();
    var lastSuggestionSheet = allSuggestSalarySheets[allSuggestSalarySheets.length - 1];
    var suggestionSheetLastRow = lastSuggestionSheet.getLastRow();
    var suggestionSheetLastColumn = lastSuggestionSheet.getLastColumn();
  
    var names = lastSuggestionSheet.getRange(4, 2, suggestionSheetLastRow - 3, 1).getValues(); // col 2 = B
    var salaries = lastSuggestionSheet.getRange(4, suggestionSheetLastColumn, suggestionSheetLastRow - 3, 1).getValues();
    var deductionBasedOnLWP = lastSuggestionSheet.getRange(4, 7, suggestionSheetLastRow - 3, 1).getValues();
  
    var newSalary = lastSuggestionSheet.getRange(4, 4, suggestionSheetLastRow - 3, 1).getValues();
    var overtime = lastSuggestionSheet.getRange(4, 6, suggestionSheetLastRow - 3, 1).getValues();
  
    var sheets = getSalarySheet.getSheets();  
    var salarySheetLast = sheets[sheets.length - 1];
    var getSalarySheetLastRow = salarySheetLast.getLastRow();
    var getSalarySheetLastColumn = salarySheetLast.getLastColumn();
  
  
    var ee = salarySheetLast.getRange(4, 10, getSalarySheetLastRow - 3, 1).getValues();
    var namesOnSalarySheet = salarySheetLast.getRange(4, 2, getSalarySheetLastRow - 3, 1).getValues(); // col 2 = B
    var getDataFromSalarySheet = [];
  
    for (var i = 0; i < namesOnSalarySheet.length; i++) {
      var name = namesOnSalarySheet[i][0];
      var eedata = ee[i];
        getDataFromSalarySheet.push({
          name: name,
          ee: eedata
        });
    }
  
    var getDataFromSuggestionSheet = [];
  
    for (var i = 0; i < names.length; i++) {
      var name = names[i][0];
      var salary = salaries[i][0];
      var deduction = deductionBasedOnLWP[i][0];
      var newSalaryThisMonth = newSalary[i][0];
      var getOvertime = overtime[i][0];
        getDataFromSuggestionSheet.push({
          name: name,
          salaryToBeCredited: salary,
          deduction: deduction,
          newSalary: newSalaryThisMonth,
          overtime: getOvertime
        });
    }
    const mergedData = getDataFromSalarySheet.map(person => {
    const match = getDataFromSuggestionSheet.find(s => 
        s.name && person.name &&
        s.name.trim().toLowerCase() === person.name.trim().toLowerCase()
      ); 
      return {
        ...person,
        salaryToBeCredited: match ? match.salaryToBeCredited : null,
        deduction : match.deduction,
        newSalary: match.newSalary,
        overtime: match.overtime
      };
    });
  
    var date = new Date();
    var year = date.getFullYear();
    var month = date.getMonth() + 1; 
    date.setMonth(date.getMonth() - 1);
  
    var options = { year: 'numeric', month: 'long' };
    var getMonthName = date.toLocaleDateString('en-US', options).replace(" ", " ");
    var newSalarySheet = getSalarySheet.insertSheet(getMonthName, getSalarySheet.getSheets().length);
   
    newSalarySheet.getRange('A50').setValue('last');
    newSalarySheet.getRange('L1').setValue('last');
    newSalarySheet.getRange('A1000').setValue('rows to be deleted');
    var lastRow = newSalarySheet.getLastRow();
    
    if (lastRow > 50) {
      newSalarySheet.deleteRows(51, lastRow - 50);
    }
    var totalColumns = newSalarySheet.getLastColumn();
    newSalarySheet.getRange('A50').clear();
    newSalarySheet.getRange('L1').clear();
  
    if (totalColumns > 0) {
      var firstRowRange = newSalarySheet.getRange(1, 1, 1, totalColumns);
      firstRowRange.merge();
      firstRowRange.setValue('Kynet Web Solutions Pvt. Ltd.');
      firstRowRange.setFontSize(24);
      newSalarySheet.setRowHeight(1, 50);
      firstRowRange.setBackground('#B7E1CD');
      firstRowRange.setHorizontalAlignment('center');
  
      var secondRowRange = newSalarySheet.getRange(2, 1, 1, totalColumns);
      secondRowRange.merge();
      secondRowRange.setValue('Salary Sheet For The Month Of ' + getMonthName);
      secondRowRange.setFontSize(16);
      newSalarySheet.setRowHeight(2, 45);
      secondRowRange.setBackground('#B7E1CD');
      secondRowRange.setHorizontalAlignment('left');
  
      newSalarySheet.setColumnWidth(1, 60);
      newSalarySheet.setColumnWidth(2, 250);
      newSalarySheet.setColumnWidth(3, 100);
      newSalarySheet.setColumnWidth(4, 95);
      newSalarySheet.setColumnWidth(5, 100);
      newSalarySheet.setColumnWidth(6, 325);
      newSalarySheet.setColumnWidth(7, 205);
      newSalarySheet.setColumnWidth(8, 155);
      newSalarySheet.setColumnWidth(9, 165);
      newSalarySheet.setColumnWidth(10, 80);
      newSalarySheet.setColumnWidth(11, 80);
      newSalarySheet.setColumnWidth(12, 115);
  
      var lastRow = newSalarySheet.getMaxRows();
      for (var i = 3; i <= lastRow; i++) {
        newSalarySheet.setRowHeight(i, 40);
      }
  
      var thirdRow = 3;
      newSalarySheet.getRange(thirdRow, 1).setValue('S. No.').setFontSize(13).setHorizontalAlignment('right').setFontWeight('bold');
  
      newSalarySheet.getRange(thirdRow, 2).setValue('Name').setFontSize(13).setHorizontalAlignment('right').setFontWeight('bold');
  
      newSalarySheet.getRange(thirdRow, 3).setValue('Basic').setFontSize(13).setHorizontalAlignment('right').setFontWeight('bold');
  
      newSalarySheet.getRange(thirdRow, 4).setValue('DA').setFontSize(13).setHorizontalAlignment('right').setFontWeight('bold');
  
      newSalarySheet.getRange(thirdRow, 5).setValue('Basic + DA').setFontSize(13).setHorizontalAlignment('right').setFontWeight('bold');
  
      newSalarySheet.getRange(thirdRow, 6).setValue('Night Shift Charges + Overtime (O.A.)').setFontSize(13).setHorizontalAlignment('right').setFontWeight('bold');
  
      newSalarySheet.getRange(thirdRow, 7).setValue('Special Allowance (SP)').setFontSize(13).setHorizontalAlignment('right').setFontWeight('bold');
  
      newSalarySheet.getRange(thirdRow, 8).setValue('Deduction').setFontSize(13).setHorizontalAlignment('right').setFontWeight('bold');
  
      newSalarySheet.getRange(thirdRow, 9).setValue('Net Credit To Bank').setFontSize(13).setHorizontalAlignment('right').setFontWeight('bold');
  
      newSalarySheet.getRange(thirdRow, 10).setValue('EE').setFontSize(13).setHorizontalAlignment('right').setFontWeight('bold');
  
      newSalarySheet.getRange(thirdRow, 11).setValue('ESI').setFontSize(13).setHorizontalAlignment('right').setFontWeight('bold');
  
      newSalarySheet.getRange(thirdRow, 12).setValue('Gross Salary').setFontSize(13).setHorizontalAlignment('right').setFontWeight('bold');
  
      var startRow = 4;
      var startCol = 1; 
  
      for(let i = 0; i < mergedData.length; i++) {
        var overtime = Number(mergedData[i].overtime);
        var newSalary = Number(mergedData[i].newSalary);
        var deduction = Number(mergedData[i].deduction);
        var totalEE = Number(mergedData[i].ee);
        var netSalary = Number(mergedData[i].salaryToBeCredited);
        var grossSalary = netSalary + totalEE;
  
        var row = startRow + i;
        newSalarySheet.getRange(row, startCol).setValue(i + 1);
        var specialAllowance = 0;
  
        newSalarySheet.getRange(row, startCol + 1).setValue(mergedData[i].name).setFontSize(14).setHorizontalAlignment("right");
        newSalarySheet.getRange(row, startCol + 4).setValue(newSalary).setFontSize(12).setHorizontalAlignment("right");
  
        newSalarySheet.getRange(row, startCol + 5).setValue(overtime).setFontSize(12).setHorizontalAlignment("right");
        newSalarySheet.getRange(row, startCol + 6).setValue(specialAllowance).setFontSize(12).setHorizontalAlignment("right");
        newSalarySheet.getRange(row, startCol + 7).setValue(deduction).setFontSize(12).setHorizontalAlignment("right");
        newSalarySheet.getRange(row, startCol + 8).setValue(netSalary).setFontSize(12).setHorizontalAlignment("right");
        newSalarySheet.getRange(row, startCol + 9).setValue(totalEE).setFontSize(12).setHorizontalAlignment("right");
        newSalarySheet.getRange(row, startCol + 11).setValue(grossSalary).setFontSize(12).setHorizontalAlignment("right");
  
        newSalarySheet.getRange(row, startCol + 2).setFormula("=ROUND((E" + row + "*100/130))");
        newSalarySheet.getRange(row, startCol + 3).setFormula("=E" + row + "-C" + row);
        newSalarySheet.getRange(row, startCol + 11).setFormula("=SUM(I" + row + ",J" + row + ",K" + row + ")");
      }
  
    }
    console.log("execution ended");
  
  }