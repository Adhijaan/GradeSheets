function jhNewEntry() {
  var spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(JH_SHEET_NAME);

  sheet.insertRowsBefore(7, 5);
  // Shapes JH Entry
  sheet.getRange("A7:A8").merge();
  sheet.getRange("A9:A10").merge();
  sheet.getRange("B7:B10").merge();
  sheet.getRange("B11:D11").merge();
  sheet.getRange("P7:P10").merge();
  sheet.getRange("Q7:Q11").merge();

  // Adds all text values in JH entry
  sheet.getRange("B11").setValue("صفحة:");
  sheet.getRange("C7").setValue("صفحة\nتسميع");
  sheet.getRange("C9").setValue("صفحة\nسوالات");
  sheet.getRange("D7").setValue("تنبيه\nType");
  sheet.getRange("D8").setValue("تنبيه\nCount");
  sheet.getRange("D9").setValue("تلقين\nType");
  sheet.getRange("D10").setValue("تلقين\nCount");
  sheet.getRange("O11").setValue("امضاء:");
  sheet.getRange("O7").setValue("Total\nTanbih:");
  sheet.getRange("O9").setValue("Total\nTalqeen:");

  // Sets JH base formulas
  // Page Tambhi count
  sheet
    .getRange("E8")
    .setFormula("=LEN(E7)")
    .autoFill(sheet.getRange("E8:N8"), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  // Page Talqeen count
  sheet
    .getRange("E10")
    .setFormula("=LEN(E9)")
    .autoFill(sheet.getRange("E10:N10"), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  // Total Tambhi/Talqeen count
  sheet.getRange("O8").setFormula("=SUM(E8:N8)");
  sheet.getRange("O10").setFormula("=SUM(E10:N10)");
  // Page numbers
  sheet
    .getRange("E11")
    .setFormula('=IF(ISBLANK($B7), "",$B7+COLUMN()-5)')
    .autoFill(sheet.getRange("E11:N11"), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

function jhDate() {
  var spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(JH_SHEET_NAME);

  /**
   * Format:
   * Hijri -  dd MMMM, yyyy
   * English - MM/dd/yy
   */
  //Hijri Date
  var myRes = HijriCalander(); // Get the array that is returned by HijriCalendar function
  var formattedDate = myRes[1] + " " + myRes[5] + ", " + myRes[3]; //formatted Day Month, YYYY
  sheet.getRange("A7:A8").setValue(formattedDate);

  //Enlgish Date
  var date = Utilities.formatDate(new Date(), "EST", "MM/dd/yyyy");
  sheet.getRange("A9:A10").setValue(date);
}

function jhMarhalahFormulas(marhala) {
  var spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(JH_SHEET_NAME);

  let tambhi_weight;
  let talqeen_weight;
  let jh_expected_pages;

  if (marhala === 1) {
    tambhi_weight = 0.1667;
    talqeen_weight = 0.5;
    jh_expected_pages = 1.5;
  } else if (marhala === 2) {
    tambhi_weight = 0.1667;
    talqeen_weight = 0.5;
    jh_expected_pages = 3;
  } else if (marhala === 3) {
    tambhi_weight = 0.1667;
    talqeen_weight = 0.33;
    jh_expected_pages = 4.5;
  } else {
    tambhi_weight = 0.1667;
    talqeen_weight = 0.33;
    jh_expected_pages = 6;
  }

  sheet
    .getRange("P7:P10")
    .setFormula(
      `=IF(ISBLANK(P11), "", IF(OR(NOT(ISNUMBER(O8)), NOT(ISNUMBER(O10)), NOT(ISNUMBER(C8)), NOT(ISNUMBER(C10))), "►", ROUND(10 - ((O8 * ((${tambhi_weight}*${jh_expected_pages})/(C8+C10))) + (O10 * ((${talqeen_weight}*${jh_expected_pages})/(C8+C10)))), 3)))`
    );
}

function jhFormat() {
  var spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(JH_SHEET_NAME);

  // Formats the entry's size, borders, fonts, and color.
  // Row height
  sheet.setRowHeightsForced(7, 4, 33);
  sheet.setRowHeightsForced(11, 1, 25);

  // Borders
  sheet
    .getRange("A7:Q11")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet
    .getRangeList(["A7:D10", "O7:P11"])
    .setBorder(null, null, null, null, null, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("E7:N11").setBorder(null, null, null, null, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  //Fonts
  sheet
    .getRange("A7:Q11")
    .setBackground(null)
    .setFontFamily("Amiri")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRangeList(["A7:A10", "C7", "C9", "D7:D10", "O7", "O9"]).setFontSize(9);
  sheet.getRangeList(["C8", "C10", "E7:N11", "O8", "O10", "P11", "Q7:Q11"]).setFontSize(10);
  sheet.getRangeList(["B7:B10", "B11:D11", "O11", "P7:P10"]).setFontSize(10);

  // Color
  sheet.getRange("A7:A11").setBackground(MonthColor());
}
