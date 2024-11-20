function mjjdNewEntry() {
  //Select the murajaat sheet as active spreadsheet
  var spreadsheet = SpreadsheetApp.getActive();
  sheet = spreadsheet.getSheetByName(MJ_SHEET_NAME);
  // Shapes a blank MJ-JD entry
  sheet.insertRowsBefore(8, 3);
  sheet.getRange("A8:A10").mergeVertically();
  sheet.getRange("Q8:W10").mergeVertically();
  sheet.getRange("Z8:Z10").mergeVertically();
  // Adds all text values in MJ-JD entry
  sheet.getRange("X8").setValue("مراجعة:");
  sheet.getRange("X9").setValue("جزءحالي:");
  sheet.getRange("X10").setValue("جديد:");
}

function mjjdBaseFormulas() {
  //Select the murajaat sheet as active spreadsheet
  var spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(MJ_SHEET_NAME);

  //Set Murajaat Thambi/Talqeen count formulas

  // Setting Tambhi 1
  let tambi1 =
    '= IFS(and(isblank($C8), isblank($D8), isblank($E8), isblank($G8)), "", and(isnumber($D8), isnumber($C8), $C8 > 0, $C8 <= 30, $D8 > 0), len(E8), true, "------")';
  sheet.getRange("F8").setFormula(tambi1);
  sheet.getRange("F8").autoFill(sheet.getRange("F8:F10"), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // Setting Talqeen 1
  sheet.getRange("F8:F10").copyTo(sheet.getRange("H8"), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // Setting Tambhi 2
  let tambi2 =
    '= IFS(and(isblank($C8), isblank($J8), isblank($K8), isblank($M8)), "", and(isnumber($C8), isnumber($J8), $C8 > 0, $C8 <= 30, $J8 > 0), len(K8), true, "------")';
  sheet.getRange("L8").setFormula(tambi2);
  sheet.getRange("L8").autoFill(sheet.getRange("L8:L10"), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // Setting Talqeen 2
  sheet.getRange("L8:L10").copyTo(sheet.getRange("N8"), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  /**
   *The formatting is thrown off - have the formatting function run after formula setting or use only copyTo()
   */

  // Setting Average MJ Marks
  sheet.getRange("P8").setFormula('=IFError(ROUND(AVERAGE(I8,O8)), "")');
  sheet.getRange("P8").autoFill(sheet.getRange("P8:P10"), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  //Set Jadeed Thambi/Talqeen count formulas
  sheet
    .getRange("U8")
    .setFormula('=IfError(IFS(AND(ISTEXT(R8), ISNUMBER(S8)), Sum(LEN(T8)),Istext(T8),Sum(LEN(T8))),"")');

  //Set Line Count Formula
  sheet
    .getRange("V8")
    .setFormula(
      '=IF(AND(ISTEXT(R8), ISNUMBER(S8), ISTEXT(S5), ISNUMBER(V5)), Q_LINE2(TRIM(R8), S8, false) - Q_LINE2(S5, V5, false) ,"")'
    );

  //Set JH score reference
  sheet.getRange("Q8:Q10").setFormula(`='${JH_SHEET_NAME}'!P7`);

  //Set JH signiture reference
  sheet.getRange("Y9").setFormula(`='${JH_SHEET_NAME}'!P11`);
}

function mjjdMarhalahFormulas(marhala) {
  //Select the murajaat sheet as active spreadsheet
  var spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(MJ_SHEET_NAME);

  // Calulcates Tambhi and Talqeen weight according to Marhalah
  let tambhi_weight;
  let talqeen_weight;
  let jd_khata_weight;

  if (marhala === 1) {
    // Juz 27 - 30
    tambhi_weight = 0.1667;
    talqeen_weight = 0.66;
    jd_khata_weight = 3;
  } else if (marhala === 2) {
    // Juz 1-3
    tambhi_weight = 0.25;
    talqeen_weight = 0.66;
    jd_khata_weight = 5;
  } else if (marhala === 3) {
    // Juz 4-5
    tambhi_weight = 0.33;
    talqeen_weight = 1;
    jd_khata_weight = 3;
  } else {
    // Juz 5+
    tambhi_weight = 0.5;
    talqeen_weight = 1;
    jd_khata_weight = 5;
  }
  // Sets MJ score formulas according to marhalah
  let mj_score_1 = sheet.getRange("I8");
  mj_score_1.setFormula(
    `=IFS(and(isblank($C8), isblank($D8), ISBLANK($E8), isblank($G8)), "", AND(ISNumber($D8), ISNumber($C8), $C8 > 0, $C8 <= 30, $D8 > 0), 10-((F8*${tambhi_weight})+(H8*${talqeen_weight})), true,  "----------")`
  );
  mj_score_1.autoFill(sheet.getRange("I8:I10"), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  let mj_score_2 = sheet.getRange("O8");
  mj_score_2.setFormula(
    `=IFS(and(isblank(C8), isblank(J8), isblank(K8), isblank(M8)), "", AND(ISNumber(C8), ISNumber(J8), C8 > 0, C8 <= 30, J8 > 0), 10-((L8*${tambhi_weight})+(N8*${talqeen_weight})), true,  "----------")`
  );
  mj_score_2.autoFill(sheet.getRange("O8:O10"), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // Sets JD score formula
  let jd_score = spreadsheet.getRange("W8:W10");
  jd_score.setFormula(
    `=ifs(and(isnumber(U8), isnumber(V8), V8 <> 0), round(10 - (U8 / CEILING(V8 / ${jd_khata_weight})) * 3, 3), and(isblank(R8), isblank(S8)), "", true, "►")`
  );
}

function mjjdDate() {
  //Select the murajaat sheet as active spreadsheet
  var spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(MJ_SHEET_NAME);

  // Hidden date - for queries and analytics
  const today = new Date();
  sheet.getRange("B8:B10").setValue(today);

  // Visible date (text and background color)
  /** Format:
   * Hijri day of month, Hijri month
   * ---
   * English weekday abv., English month abv., English day of month
   */
  const dateInfo = HijriCalander();
  const weekdays = ["Sun", "Mon", "Tues", "Wed", "Thur", "Fri", "Sat"];
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  let dateString = `${dateInfo[1]} ${dateInfo[5]} \n --- \n ${weekdays[today.getDay()]}, ${
    months[today.getMonth()]
  } ${today.getDate()}`;

  sheet.getRange("A8:A10").setValue(dateString);
}

function mjjdFormat() {
  var spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(MJ_SHEET_NAME);

  // Formats the entry's size, borders, fonts, and color.
  // Row height
  sheet.setRowHeightsForced(8, 3, 27);
  // Borders
  spreadsheet
    .getRange("A8:Z10")
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setBorder(null, null, null, null, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet
    .getRangeList(["A8:D10", "O8:X10", "Y8:Z10"])
    .setBorder(null, null, null, null, true, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // Fonts
  sheet
    .getRange("A8:Z10")
    .setBackground(null)
    .setFontFamily("Amiri")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange("A8:A10").setFontSize(9);
  sheet.getRange("C8:C10").setFontSize(11);
  sheet.getRangeList(["B8:B10", "D8:D10", "F8:F10", "H8:J10", "L8:L10", "N8:O10", "T8:U10", "Y8:Y10"]).setFontSize(10);
  sheet.getRangeList(["E8:E10", "G8:G10", "K8:K10", "M8:M10", "X8:X10", "Z8:Z10"]).setFontSize(8);
  sheet.getRangeList(["R8:S10", "V8:W10"]).setFontSize(14);
  // Color
  sheet.getRange("A8:A10").setBackground(MonthColor());

  /** There should be a better way */
  var conditionalFormatRules = sheet.getConditionalFormatRules();
  conditionalFormatRules.splice(
    64,
    1,
    SpreadsheetApp.newConditionalFormatRule()
      .setRanges([
        sheet.getRange("I8:I268"),
        sheet.getRange("O8:O268"),
        sheet.getRange("Q8:Q268"),
        sheet.getRange("W8:W268"),
      ])
      .setGradientMinpointWithValue("#EA9999", SpreadsheetApp.InterpolationType.NUMBER, "5")
      .setGradientMidpointWithValue("#FFE599", SpreadsheetApp.InterpolationType.NUMBER, "7.5")
      .setGradientMaxpointWithValue("#B6D7A8", SpreadsheetApp.InterpolationType.NUMBER, "10")
      .build()
  );
  sheet.setConditionalFormatRules(conditionalFormatRules);
}
