function moveAndFillEfficiencyTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Efficiency Tracker");
  const today = new Date();
  const timeZone = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(today, timeZone, "MM/dd/yy");
  
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const yesterdayStr = Utilities.formatDate(yesterday, timeZone, "MM/dd/yy");

  const lastRow = sheet.getLastRow();
  const colA = sheet.getRange("A1:A" + lastRow).getValues().flat();

  let dateRow = -1;
  for (let i = 0; i < colA.length; i++) {
    let cell = colA[i];
    if (cell instanceof Date) {
      let formatted = Utilities.formatDate(cell, timeZone, "MM/dd/yy");
      if (formatted === todayStr) {
        dateRow = i + 1;
        break;
      }
    } else if (typeof cell === "string") {
      if (cell.trim() === todayStr) {
        dateRow = i + 1;
        break;
      }
    }
  }

  if (dateRow === -1) {
    // Today NOT found, find yesterday's date row
    let yesterdayRow = -1;
    for (let i = 0; i < colA.length; i++) {
      let cell = colA[i];
      if (cell instanceof Date) {
        let formatted = Utilities.formatDate(cell, timeZone, "MM/dd/yy");
        if (formatted === yesterdayStr) {
          yesterdayRow = i + 1;
          break;
        }
      } else if (typeof cell === "string") {
        if (cell.trim() === yesterdayStr) {
          yesterdayRow = i + 1;
          break;
        }
      }
    }
    if (yesterdayRow === -1) {
      yesterdayRow = lastRow;
    }
    dateRow = yesterdayRow + 25;
    sheet.insertRowsBefore(dateRow, 25);
    sheet.getRange(dateRow, 1).setValue(today);
  } else {
    sheet.insertRowsAfter(dateRow, 25);
  }

  const row1Values = sheet.getRange(1, 2, 1, 13).getValues();
  sheet.getRange(dateRow - 1, 2, 1, 13).setValues(row1Values);

  const row2ValuesPart1 = sheet.getRange(2, 2, 1, 9).getValues()[0];
  const row2FormulasPart2 = sheet.getRange(2, 11, 1, 4).getFormulasR1C1()[0];
  let row2CombinedValues = row2ValuesPart1;
  let row2CombinedFormulas = [];

  for (let i = 0; i < 4; i++) {
    let formula = row2FormulasPart2[i];
    if (formula) formula = formula.replace(/R\d+/g, "R" + dateRow);
    row2CombinedFormulas.push(formula || "");
  }

  sheet.getRange(dateRow, 2, 1, 9).setValues([row2CombinedValues]);
  sheet.getRange(dateRow, 11, 1, 4).setFormulasR1C1([row2CombinedFormulas]);

  const specialCols = [6, 9, 10];
  const specialFormulas = specialCols.map(col =>
    sheet.getRange(3, col, 1, 1).getFormulaR1C1()
  );

  for (let i = 1; i < 25; i++) {
    const targetRow = dateRow + i;
    for (let j = 0; j < specialCols.length; j++) {
      if (specialFormulas[j]) {
        let f = specialFormulas[j].replace(/R\d+/g, "R" + targetRow);
        sheet.getRange(targetRow, specialCols[j]).setFormulaR1C1(f);
      }
    }
  }
}


