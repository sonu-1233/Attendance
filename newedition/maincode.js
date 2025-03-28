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
      console.log('in the right column');
      var totalsalaryToBeCredited = getSuggestionsheet.getRange(thisRow, thisCol).getValue();
      var getOverTimeDays = getSuggestionsheet.getRange(thisRow, 5).getValue();
      getOverTimeDays = Number(getOverTimeDays);
      getOverTimeDays = Math.floor(getOverTimeDays);
  
      var oneDaySalaryCount = totalsalaryToBeCredited / totalDays;
      oneDaySalaryCount = Math.floor(oneDaySalaryCount);
      var leaveWithoutPayThisMonth = getSuggestionsheet.getRange(thisRow, 8).getValue();
      var deductionBasedOnLWP = leaveWithoutPayThisMonth * oneDaySalaryCount;
      deductionBasedOnLWP = Math.floor(Number(deductionBasedOnLWP));
      var overtimeSalary = 0;
  
      if(getOverTimeDays != '') {
        var overtimeSalary = oneDaySalaryCount * getOverTimeDays;
        overtimeSalary = Math.floor(overtimeSalary)
      }
      var suggestTotalSalary = (totalsalaryToBeCredited - deductionBasedOnLWP) + overtimeSalary;
      suggestTotalSalary = Math.floor(suggestTotalSalary);
      getSuggestionsheet.getRange(thisRow, 6).setValue(overtimeSalary);
      getSuggestionsheet.getRange(thisRow, 7).setValue(deductionBasedOnLWP);
      getSuggestionsheet.getRange(thisRow, 9).setValue(suggestTotalSalary);
  
      console.log("SALARY SUGGESTION COMPLETED");
    }
  }