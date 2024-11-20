/**
 * @OnlyCurrentDoc
 */
// Version 3.9

/**THIS CODE IS PROPERTY OF MUKHAYYAM LOS ANGELES*/
// Funcitons are called by triggers (menu on the left)
// Set the spreadsheet up for a new week
function NewWeek() {
  JuzHaliReset();
  TasmeealJuzReset();
  MurajatJadeedReset();
  // Formatting
  MJformatting();
  JHFormatting();
  TJFormatting();
}
function MurajatJadeedReset() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Murajaat-Jadeed"), true);

  //Making the new weeks sheet
  spreadsheet.getRange("A37:AA7783").activate();
  spreadsheet.getRange("A6:AA7752").moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange("A6").activate();
  spreadsheet
    .getRange("A37:AA66")
    .copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  //Clearing the new sheet
  spreadsheet
    .getRangeList([
      "B8:B33",
      "C8:C33",
      "D8:D33",
      "F8:F33",
      "I8:I33",
      "J8:J33",
      "L8:L33",
      "P10:P33",
      "U10:U33",
      "W10:W33",
      "Y10:Y33",
      "Z10:AA33",
      "Q10:S33",
    ])
    .activate()
    .clear({ contentsOnly: true, skipFilteredRows: true });

  //Date Setter
  for (var i = 0; i < 6; i++) {
    var d = new Date();
    var accCell = "A" + (10 + 4 * i) + ":" + "A" + (11 + 4 * i);
    spreadsheet.getRange(accCell).activate();
    spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
    spreadsheet.getCurrentCell().setValue(DateHelper(d.setDate(d.getDate() + (i + 1 - d.getDay()))));
  }
  for (var i = 0; i < 6; i++) {
    var d = new Date();
    var accCell = "A" + (12 + 4 * i) + ":" + "A" + (13 + 4 * i);
    spreadsheet.getRange(accCell).activate();
    spreadsheet.getCurrentCell().setValue(WriteIslamicDate(d.setDate(d.getDate() + (i + 1 - d.getDay()))));
  }
  spreadsheet.getRange("A9").activate();
  var newWkNum = spreadsheet.getCurrentCell().getValue();
  spreadsheet.getRange("A9").setValue(Number(newWkNum) + 1);

  // JH score grabber
  for (var i = 0; i < 6; i++) {
    var accCell = "P" + (10 + 4 * i) + ":" + "P" + (13 + 4 * i);
    spreadsheet.getRange(accCell).activate();
    var refCell = "='Juz Haali'!R" + (8 + 5 * i);
    spreadsheet.getCurrentCell().setFormula(refCell);
  }

  // JH signature
  for (var i = 0; i < 6; i++) {
    var accCell = "Y" + (11 + 4 * i);
    spreadsheet.getRange(accCell).activate();
    var refCell = "='Juz Haali'!R" + (12 + 5 * i);
    spreadsheet.getCurrentCell().setFormula(refCell);
  }

  // TJ signature
  for (var i = 0; i < 6; i++) {
    var accCell = "Y" + (13 + 4 * i);
    spreadsheet.getRange(accCell).activate();
    var refCell = "='Tasmee al Juz'!R" + (12 + 5 * i);
    spreadsheet.getCurrentCell().setFormula(refCell);
  }

  // Makan al Haali grabber
  spreadsheet.getRange("S7").activate();
  refCell = spreadsheet
    .getRange(LastFill("Murajaat-Jadeed", "Murajaat-Jadeed", "Q", 41, 64, "", false, true)[0])
    .getValue();
  if (refCell != "") {
    spreadsheet.getCurrentCell().setValue(refCell);
  } else {
    //in case nothing was new in the last week
    spreadsheet.getCurrentCell().setValue(spreadsheet.getRange("S38").getValue());
  }
  spreadsheet.getRange("U7").activate();
  refCell = spreadsheet
    .getRange(LastFill("Murajaat-Jadeed", "Murajaat-Jadeed", "R", 41, 64, "", false, true)[0])
    .getValue();
  if (refCell != "") {
    spreadsheet.getCurrentCell().setValue(refCell);
  } else {
    spreadsheet.getCurrentCell().setValue(spreadsheet.getRange("U38").getValue());
  }
  //Hifz Stats
  // CWMO = Current Week Marks Overview
  // YDMO = Year to Date Marks Overview
  spreadsheet.getRange("AB7").activate(); // CWMO Avg. M Marks
  spreadsheet.getCurrentCell().setFormula('=IfError(Round(Average(H8:H33,N8:N33),2),"None")');
  spreadsheet.getRange("AC7").activate(); // CWMO Avg. JH Marks
  spreadsheet.getCurrentCell().setFormula('=IfError(ROUND(AVERAGE(P10:P33), 2), "None")');
  spreadsheet.getRange("AD7").activate(); // CWMO Avg. J Marks
  spreadsheet.getCurrentCell().setFormula('=IfError(ROUND(AVERAGE(U10:U33), 2), "None")');
  spreadsheet.getRange("AE7").activate(); // CWMO Avg. M Pages
  spreadsheet.getCurrentCell().setFormula('=IfError(ROUND(AVERAGE(V10:V33), 2), "None")');
  spreadsheet.getRange("AB9").activate(); // YDMO Avg. M Marks
  spreadsheet.getCurrentCell().setFormula('=IfError(ROUND(AVERAGE(N:N,H:H), 2), "None")');
  spreadsheet.getRange("AC9").activate(); // YDMO Avg. JH Marks
  spreadsheet.getCurrentCell().setFormula('=IfError(ROUND(AVERAGE(P:P), 2), "None")');
  spreadsheet.getRange("AD9").activate(); // YDMO Avg. J Marks
  spreadsheet.getCurrentCell().setFormula('=IfError(ROUND(AVERAGE(U:U), 2), "None")');
  spreadsheet.getRange("AE9").activate(); // YDMO Avg. J Pages
  var jadeedTotals = [];
  for (var i = 0; i < newWkNum; i++) {
    var addRange = "V" + (34 + i * 31);
    jadeedTotals.push(addRange);
  }
  spreadsheet.getCurrentCell().setFormula("=IfError(ROUND(AVERAGE(" + jadeedTotals.toString() + '), 2), "None")');
  spreadsheet.getRange("AG7").activate(); // Weeks to Completion
  var refCell = LastFill("Murajaat-Jadeed", "Juz Haali", "E", 40, 68, "", false, true)[0];
  spreadsheet
    .getCurrentCell()
    .setValue(
      "=IF(ISFORMULA('Juz Haali'!" + refCell + '),"Update JH",ROUND((581 -' + "'Juz Haali'!" + refCell + ")/AE9,2))"
    );
  // Murajaat Tracker
  spreadsheet
    .getRange("AB14")
    .setFormula(
      '=ArrayFormula(if(len(iferror(QUERY($B$37:$P$5253,"SELECT N, I, H, C WHERE B = \'" & To_Text($AB$11) &"\' LIMIT 9",),)),QUERY($B$37:$P$5253,"SELECT N, I, H, C WHERE B = \'" & To_Text($AB$11) &"\' LIMIT 9",),"-"))'
    );
}
function JuzHaliReset() {
  //Making new week's sheet
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Juz Haali"), true);
  spreadsheet.getRange("A39:S5253").activate();
  spreadsheet.getRange("A7:S5253").moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange("A7").activate();
  spreadsheet
    .getRange("A39:S69")
    .copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  // Clearing the new sheet
  spreadsheet
    .getRangeList(["D8:D11", "D13:D16", "D18:D21", "D23:D26", "D28:D31", "D33:D36"])
    .activate()
    .clear({ contentsOnly: true, skipFilteredRows: true });
  for (var i = 0; i < 6; i++) {
    var firstRow = "G" + (8 + 5 * i) + ":" + "P" + (8 + 5 * i);
    var secondRow = "G" + (10 + 5 * i) + ":" + "P" + (10 + 5 * i);
    spreadsheet.getRange(firstRow).activate();
    spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
    spreadsheet.getRange(secondRow).activate();
    spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  }
  spreadsheet
    .getRangeList(["R12", "R17", "R22", "R27", "R32", "R37", "S8:S37"])
    .activate()
    .clear({ contentsOnly: true, skipFilteredRows: true });

  // Date setting
  for (var i = 0; i < 6; i++) {
    var c = new Date();
    var accCell = "B" + (8 + 5 * i) + ":" + "B" + (9 + 5 * i);
    spreadsheet.getRange(accCell).activate();
    spreadsheet.getCurrentCell().setValue(DateHelper(c.setDate(c.getDate() + (i + 1 - c.getDay()))));
  }
  for (var i = 0; i < 6; i++) {
    var c = new Date();
    var accCell = "B" + (10 + 5 * i) + ":" + "B" + (11 + 5 * i);
    spreadsheet.getRange(accCell).activate();
    spreadsheet.getCurrentCell().setValue(WriteIslamicDate(c.setDate(c.getDate() + (i + 1 - c.getDay()))));
  }
  spreadsheet.getRange("A7").activate();
  var newWkNum = spreadsheet.getCurrentCell().getValue().split(" ");
  newWkNum.splice(1, 1, Number(newWkNum[1]) + 1);
  var holder = newWkNum.join(" ");
  spreadsheet.getRange("A7").setValue(holder);
}
function TasmeealJuzReset() {
  //Making new week's sheet
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Tasmee al Juz"), true);
  spreadsheet.getRange("A39:S5253").activate();
  spreadsheet.getRange("A7:S5253").moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange("A7").activate();
  spreadsheet
    .getRange("A39:S69")
    .copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  // Clearing the new sheet
  spreadsheet
    .getRangeList(["D8:D11", "D13:D16", "D18:D21", "D23:D26", "D28:D31", "D33:D36"])
    .activate()
    .clear({ contentsOnly: true, skipFilteredRows: true });
  for (var i = 0; i < 6; i++) {
    var firstRow = "G" + (8 + 5 * i) + ":" + "P" + (8 + 5 * i);
    var secondRow = "G" + (10 + 5 * i) + ":" + "P" + (10 + 5 * i);
    spreadsheet.getRange(firstRow).activate();
    spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
    spreadsheet.getRange(secondRow).activate();
    spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  }
  spreadsheet
    .getRangeList(["R12", "R17", "R22", "R27", "R32", "R37", "S8:S37"])
    .activate()
    .clear({ contentsOnly: true, skipFilteredRows: true });

  // Date setting
  for (var i = 0; i < 6; i++) {
    var c = new Date();
    var accCell = "B" + (8 + 5 * i) + ":" + "B" + (9 + 5 * i);
    spreadsheet.getRange(accCell).activate();
    spreadsheet.getCurrentCell().setValue(DateHelper(c.setDate(c.getDate() + (i + 1 - c.getDay()))));
  }
  for (var i = 0; i < 6; i++) {
    var c = new Date();
    var accCell = "B" + (10 + 5 * i) + ":" + "B" + (11 + 5 * i);
    spreadsheet.getRange(accCell).activate();
    spreadsheet.getCurrentCell().setValue(WriteIslamicDate(c.setDate(c.getDate() + (i + 1 - c.getDay()))));
  }
  spreadsheet.getRange("A7").activate();
  var newWkNum = spreadsheet.getCurrentCell().getValue().split(" ");
  newWkNum.splice(1, 1, Number(newWkNum[1]) + 1);
  var holder = newWkNum.join(" ");
  spreadsheet.getRange("A7").setValue(holder);
}
function KamilJuzTasmeeEntry() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Kamil Juz Tasmee"), true);
  // New entry
  spreadsheet.getRange("A13:AA1017").activate();
  spreadsheet.getRange("A8:AA1012").moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange("A8").activate();
  spreadsheet
    .getRange("A13:AA17")
    .copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  // clearing
  spreadsheet.getRangeList(["A8", "B8:C11", "E8:X8", "E10:X10", "Z12", "AA8:AA12"]).clearContent();
  // Date & week setting
  spreadsheet.getRange("A8:A12").merge();
  var date = new Date();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Murajaat-Jadeed"), true);
  spreadsheet.getRange("A9").activate();
  var wkNum = spreadsheet.getCurrentCell().getValue();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Kamil Juz Tasmee"), true);
  spreadsheet.getRange("A8").setValue("Wk " + wkNum);
  spreadsheet.getRange("B8").setValue(DateHelper(new Date()));
  spreadsheet.getRange("B10").setValue(WriteIslamicDate(new Date()));
  // Format
  spreadsheet.getRange("A8:B12").setBackground(ColorHelper("Kamil Juz Tasmee"));
  if (spreadsheet.getRange("A8").getValue() === spreadsheet.getRange("A13").getValue()) {
    var merge = spreadsheet.getRange("A13").getMergedRanges();
    var start = merge[0].getLastRow();
    spreadsheet.getRange("A8:A" + start).merge();
  }
  // Score setting
  spreadsheet.getRange("Z8:Z11").setValue('=IF(OR(ISBLANK(Z12)),"",(10-((Y9*0.1)+(Y11*0.2))))');
}
function Ikhtebaar(siparaAmt) {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Ikhtebaar al Ajza"));
  // Shifting Sheet
  spreadsheet.getRange("A7:S5253").moveTo(spreadsheet.getRange("A" + (24 + siparaAmt * 2)));
  var oldTitle = LastFill("Ikhtebaar al Ajza", "Ikhtebaar al Ajza", "A", 25, 125, "اختبار", true, false);
  // Creating Title range
  spreadsheet.getRange("A7").activate();
  spreadsheet
    .getRange(oldTitle[0] + ":S" + (oldTitle[1] + 7))
    .copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  // Creating suwaal boxes
  for (var i = 0; i < siparaAmt - 1; i++) {
    spreadsheet.getRange("A" + (15 + i * 2)).activate();
    spreadsheet
      .getRange("A13:S14")
      .copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  }
  // Creating Mulahazaat
  var newMulahazat = 13 + siparaAmt * 2; // Starting row of new Mulahazaat
  var oldMulahazaat = LastFill(
    "Ikhtebaar al Ajza",
    "Ikhtebaar al Ajza",
    "A",
    newMulahazat,
    newMulahazat + 125,
    "ملاحظات",
    true,
    false
  );
  spreadsheet
    .getRange(oldMulahazaat[0] + ":S" + (oldMulahazaat[1] + 8))
    .copyTo(spreadsheet.getRange("A" + newMulahazat), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  // Auto filling the Title
  spreadsheet.getRange("E9:K10").setValue(spreadsheet.getRange("H3:N3").getValue()); // Talib name
  spreadsheet.getRange("O9:P10").setValue(spreadsheet.getRange("Q3").getValue()); // Talib sanah
  spreadsheet.getRange("S9:S10").clearContent(); // Sipara amount
  // Clearing Questions
  spreadsheet.getRange("A13:O" + (siparaAmt * 2 + 12)).clearContent(); // 12 to account for content at top
  spreadsheet.getRange("R13:S" + (siparaAmt * 2 + 12)).clearContent();
  // Clearing Mulahazat + Setting Score Formula
  spreadsheet.getRange("A" + (newMulahazat + 4) + ":N" + (newMulahazat + 8)).clearContent(); // acc Mulahazaat
  spreadsheet.getRange("Q" + (newMulahazat + 4)).clearContent(); // Mukhtabir name
  spreadsheet.getRange("S" + newMulahazat).clearContent(); // Result
  spreadsheet
    .getRange("Q" + newMulahazat)
    .setValue("=IfError(ROUND(AVERAGE(P13:Q" + (newMulahazat - 1) + ')*10,1),"")'); // score
  // Date setting
  spreadsheet.getRange("Q" + (newMulahazat + 6)).setValue(DateHelper(new Date())); // Gregorian
  spreadsheet.getRange("R" + (newMulahazat + 6)).setValue(WriteIslamicDate(new Date())); // Hijri
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Murajaat-Jadeed"), true); // Week # grabber
  var wkNum = spreadsheet.getRange("A9").getValue();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Ikhtebaar al Ajza"), true);
  spreadsheet.getRange("S" + (newMulahazat + 7)).setValue(wkNum);
  spreadsheet
    .getRange("Q" + (newMulahazat + 6) + ":S" + (newMulahazat + 8))
    .setBackground(ColorHelper("Ikhtebaar al Ajza"));
  spreadsheet.getRange("S9").activate();
}
// Sanahs for ikhtebar bc the UI cannot pass parameters
function sanah1() {
  Ikhtebaar(6);
}
function sanah2() {
  Ikhtebaar(13);
}
function sanah3() {
  Ikhtebaar(22);
}
function sanah4() {
  Ikhtebaar(30);
}
function MulahazaatCreator() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Ikhtebaar al Ajza"));
  var mRange = LastFill("Ikhtebaar al Ajza", "Ikhtebaar al Ajza", "A", 7, 150, "ملاحظات", true, false);
  spreadsheet.getRange(mRange[0]).activate();
  // Ahkaam values extractor
  var ahkaam = [
    "U13",
    "U15",
    "U17",
    "U19",
    "U21",
    "U23",
    "W13",
    "W15",
    "W17",
    "W19",
    "W21",
    "W23",
    "Y13",
    "Y15",
    "Y17",
    "Y19",
    "Y21",
    "Y23",
    "AA13",
    "AA15",
    "AA17",
    "AA19",
    "AA21",
    "AA23",
  ];
  var mAhkaam = [];
  for (var i = 0; i < ahkaam.length; i++) {
    if (spreadsheet.getRange(ahkaam[i]).isChecked()) {
      if (ahkaam[i].charAt(0) == "U") mAhkaam.push("V" + (i * 2 + 13));
      if (ahkaam[i].charAt(0) == "W") mAhkaam.push("X" + ((i - 6) * 2 + 13));
      if (ahkaam[i].charAt(0) == "Y") mAhkaam.push("Z" + ((i - 12) * 2 + 13));
      if (ahkaam[i].charAt(0) == "A") mAhkaam.push("AB" + ((i - 18) * 2 + 13));
    }
  }
  var ahkaamV = [];
  for (var i = 0; i < mAhkaam.length; i++) {
    ahkaamV.push(spreadsheet.getRange(mAhkaam[i]).getValue());
  }
  spreadsheet.getRange("E" + (mRange[1] + 4)).setValue(ahkaamV.join("، "));
  // Huruf values extractor
  var huruf = [
    "U27",
    "U29",
    "U31",
    "U33",
    "U35",
    "U37",
    "U39",
    "W27",
    "W29",
    "W31",
    "W33",
    "W35",
    "W37",
    "W39",
    "Y27",
    "Y29",
    "Y31",
    "Y33",
    "Y35",
    "Y37",
    "Y39",
    "AA27",
    "AA29",
    "AA31",
    "AA33",
    "AA35",
    "AA37",
    "AA39",
  ];
  var mHuruf = [];
  for (var i = 0; i < huruf.length; i++) {
    if (spreadsheet.getRange(huruf[i]).isChecked()) {
      if (huruf[i].charAt(0) == "U") mHuruf.push("V" + (i * 2 + 27));
      if (huruf[i].charAt(0) == "W") mHuruf.push("X" + ((i - 7) * 2 + 27));
      if (huruf[i].charAt(0) == "Y") mHuruf.push("Z" + ((i - 14) * 2 + 27));
      if (huruf[i].charAt(0) == "A") mHuruf.push("AB" + ((i - 21) * 2 + 27));
    }
  }
  var hurufV = [];
  for (var i = 0; i < mHuruf.length; i++) {
    hurufV.push(spreadsheet.getRange(mHuruf[i]).getValue());
  }
  spreadsheet.getRange("A" + (mRange[1] + 4)).setValue(hurufV.join("، "));
  spreadsheet.getRange("U13:AB42").uncheck();
}
// All formatting functions
function MJformatting() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Murajaat-Jadeed"), true);
  spreadsheet.getRange("A1:AA35").activate();
  spreadsheet
    .getActiveRangeList()
    .setFontFamily("Amiri")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  spreadsheet.getRange("A6:AA35").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID)
    .setBackground("white");
  spreadsheet.getRange("A6:A33").activate();
  spreadsheet.getActiveRangeList().setBackground(ColorHelper("Murajaat-Jadeed"));
  spreadsheet.getRange("A6:AA7").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange("C7:N7").activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange("N7"));
  spreadsheet
    .getActiveRangeList()
    .setBorder(null, null, null, null, true, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  for (var i = 0; i < 6; i++) {
    var accCell = "A8:" + "O" + (13 + 4 * i);
    spreadsheet.getRange(accCell).activate();
    spreadsheet
      .getActiveRangeList()
      .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  spreadsheet.getRange("B8:B33").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange("P8:AA33").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  for (var i = 0; i < 6; i++) {
    var accCell = "X" + (10 + 4 * i) + ":" + "Y" + (13 + 4 * i);
    spreadsheet.getRange(accCell).activate();
    spreadsheet
      .getActiveRangeList()
      .setBorder(null, null, null, null, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  }
  spreadsheet.getRange("A34:AA35").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // Hifz Stats Format
  spreadsheet.getRange("AB5:AE22").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getRange("AF6:AG9").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getRange("AB5:AE22").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange("AB6:AG9").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange("AB10:AE13").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange("AB13:AE13").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(null, null, null, null, true, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
}
function JHFormatting() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Juz Haali"), true);
  spreadsheet.getRange("A7:B37").activate();
  spreadsheet.getActiveRangeList().setBackground(ColorHelper("Juz Haali"));
  spreadsheet.getRange("C7:S37").activate();
  spreadsheet.getActiveRangeList().setBackground("white");
  spreadsheet.getRange("A1:S37").activate();
  spreadsheet
    .getActiveRangeList()
    .setFontFamily("Amiri")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  spreadsheet.getRange("A7:A37").setTextRotation(-90);
  spreadsheet.getRange("A7:S37").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getRange("B7:S7").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange("A7:A37").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  for (var i = 0; i < 6; i++) {
    var accCell = "B" + (8 + 5 * i) + ":" + "S" + (12 + 5 * i);
    spreadsheet.getRange(accCell).activate();
    spreadsheet
      .getActiveRangeList()
      .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  spreadsheet.getRange("C7:F37").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange("R7:S37").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, true, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}
function TJFormatting() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Tasmee al Juz"), true);
  spreadsheet.getRange("A7:B37").activate();
  spreadsheet.getActiveRangeList().setBackground(ColorHelper("Tasmee al Juz"));
  spreadsheet.getRange("C7:S37").activate();
  spreadsheet.getActiveRangeList().setBackground("white");
  spreadsheet.getRange("A1:S37").activate();
  spreadsheet
    .getActiveRangeList()
    .setFontFamily("Amiri")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  spreadsheet.getRange("A7:A37").setTextRotation(-90);
  spreadsheet.getRange("A7:S37").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getRange("B7:S7").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange("A7:A37").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  for (var i = 0; i < 6; i++) {
    var accCell = "B" + (8 + 5 * i) + ":" + "S" + (12 + 5 * i);
    spreadsheet.getRange(accCell).activate();
    spreadsheet
      .getActiveRangeList()
      .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  spreadsheet.getRange("C7:F37").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange("R7:S37").activate();
  spreadsheet
    .getActiveRangeList()
    .setBorder(true, true, true, true, true, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}
// KJT & Ikhtebaar reformat needs to be created
// Cell Data Extractor(s)
function LastFill(orginalSheet, refSheet, column, startRange, endRange, search, included, toTop) {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(refSheet), true);

  var range = spreadsheet.getRange(column + startRange + ":" + column + endRange).getValues();
  //checker
  if (included) {
    if (toTop) {
      for (var i = range.length - 1; i >= 0; i--) {
        var lastRow = i + 1; // bc Sheets has rows start from 1
        if (
          range[i].every(function (c) {
            return c == search;
          })
        ) {
          break;
        }
      }
    } else {
      for (var i = 0; i < range.length; i++) {
        var lastRow = i + 1; // bc Sheets has rows start from 1
        if (
          range[i].every(function (c) {
            return c == search;
          })
        ) {
          break;
        }
      }
    }
  } else {
    if (toTop) {
      for (var i = range.length - 1; i >= 0; i--) {
        var lastRow = i + 1; // bc Sheets has rows start from 1
        if (
          !range[i].every(function (c) {
            return c == search;
          })
        ) {
          break;
        }
      }
    } else {
      for (var i = 0; i < range.length; i++) {
        var lastRow = i + 1; // bc Sheets has rows start from 1
        if (
          !range[i].every(function (c) {
            return c == search;
          })
        ) {
          break;
        }
      }
    }
  }

  lastRow = lastRow + startRange - 1; // To adjust for range start
  var refCell = column + lastRow;
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(orginalSheet), true);
  var result = [refCell, lastRow];
  return result;
}
function ColorHelper(orginalSheet) {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Murajaat-Jadeed"), true);
  spreadsheet.getRange("A9").activate();
  var wkNum = spreadsheet.getCurrentCell().getValue();
  var colorNum = wkNum % 10;
  var colors = [
    "#e6b8af",
    "#f4cccc",
    "#fce5cd",
    "#fff2cc",
    "#d9ead3",
    "#d0e0e3",
    "#c9daf8",
    "#cfe2f3",
    "#d9d2e9",
    "#ead1dc",
  ];
  var weekColor = colors[colorNum];
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(orginalSheet), true);
  return weekColor;
}
// Date helpers
function DateHelper(aDate) {
  var newDate = new Date(aDate);
  var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  var months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];
  var formattedDate = days[newDate.getDay()] + " " + months[newDate.getMonth()] + " " + newDate.getDate();
  return formattedDate;
}
function WriteIslamicDate(aDate) {
  var wdNames = ["يوم الأحد", "يوم الإثنين", "يوم الثلاثاء", "يوم الأربعاء", "يوم الخميس", "يوم الجمعة", "يوم السبت"];
  var iMonthNames = [
    "محرم الحرام",
    "صفر المظفر",
    "ربيع الاول",
    "ربيع الاخر",
    "جمادي الاولى",
    "جمادي الاخرى",
    "رجب الاصب",
    "شعبان الكريم",
    "رمضان المعظم",
    "شوّال المکرّم",
    "ذو القعدة الحرام",
    "ذو الحِجَّة الحرام",
  ];
  var iDate = Kuwaiticalendar(aDate);
  var outputIslamicDate = wdNames[iDate[4]] + ", " + iDate[5] + " " + iMonthNames[iDate[6]];
  return outputIslamicDate;
}
function Kuwaiticalendar(aDate) {
  var today = new Date(aDate);
  // today.setDate(today.getDate()+adjust-today.getDay());
  // if(adjust) {
  // 	adjustmili = 1000*60*60*24*adjust;
  // 	todaymili = today.getTime()+adjustmili;
  // 	today = new Date(todaymili);
  // }
  day = today.getDate();
  month = today.getMonth();
  year = today.getFullYear();
  m = month + 1;
  y = year;
  if (m < 3) {
    y -= 1;
    m += 12;
  }

  a = Math.floor(y / 100);
  b = 2 - a + Math.floor(a / 4);
  if (y < 1583) b = 0;
  if (y == 1582) {
    if (m > 10) b = -10;
    if (m == 10) {
      b = 0;
      if (day > 4) b = -10;
    }
  }

  jd = Math.floor(365.25 * (y + 4716)) + Math.floor(30.6001 * (m + 1)) + day + b - 1524;

  b = 0;
  if (jd > 2299160) {
    a = Math.floor((jd - 1867216.25) / 36524.25);
    b = 1 + a - Math.floor(a / 4);
  }
  bb = jd + b + 1524;
  cc = Math.floor((bb - 122.1) / 365.25);
  dd = Math.floor(365.25 * cc);
  ee = Math.floor((bb - dd) / 30.6001);
  day = bb - dd - Math.floor(30.6001 * ee);
  month = ee - 1;
  if (ee > 13) {
    cc += 1;
    month = ee - 13;
  }
  year = cc - 4716;

  wd = today.getDay();

  iyear = 10631 / 30;
  epochastro = 1948084;
  epochcivil = 1948085;

  shift1 = 8.01 / 60;

  z = jd - epochastro;
  cyc = Math.floor(z / 10631);
  z = z - 10631 * cyc;
  j = Math.floor((z - shift1) / iyear);
  iy = 30 * cyc + j;
  z = z - Math.floor(j * iyear + shift1);
  im = Math.floor((z + 28.5001) / 29.5);
  if (im == 13) im = 12;
  id = z - Math.floor(29.5001 * im - 29);

  var myRes = new Array(8);

  myRes[0] = day; //calculated day (CE)
  myRes[1] = month - 1; //calculated month (CE)
  myRes[2] = year; //calculated year (CE)
  myRes[3] = jd - 1; //julian day number
  myRes[4] = wd; //weekday number
  myRes[5] = id; //islamic date
  myRes[6] = im - 1; //islamic month
  myRes[7] = iy; //islamic year

  return myRes;
}
// Auto Functions
function onOpen() {
  // "Commands" menu
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("Commands");
  menu.addItem("Kamil Juz Entry", "KamilJuzTasmeeEntry");
  menu.addSeparator();
  var Formats = ui.createMenu("Formatting");
  Formats.addItem("Murajaat/Jadeed Sheet", "MJformatting");
  Formats.addItem("Juz Hali Sheet", "JHFormatting");
  Formats.addItem("Tasmee al Juz Sheet", "TJFormatting");
  menu.addSubMenu(Formats);
  menu.addSeparator();
  var Ikhtebaar = ui.createMenu("Ikhtebaar");
  Ikhtebaar.addItem("Sanah 1", "sanah1");
  Ikhtebaar.addItem("Sanah 2", "sanah2");
  Ikhtebaar.addItem("Sanah 3", "sanah3");
  Ikhtebaar.addItem("Sanah 4", "sanah4");
  Ikhtebaar.addItem("Custom", "specialIkh");
  Ikhtebaar.addItem("Create Mulahazaat", "MulahazaatCreator");
  menu.addSubMenu(Ikhtebaar);
  menu.addToUi();
}
function specialIkh() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Enter desired suwaal amount", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    if (response.getResponseText().valueOf() <= 1) {
      ui.alert("Suwaal amount is too low.");
    } else if (response.getResponseText().valueOf() > 40) {
      ui.alert("Suwaal amount is too high.");
    } else {
      Ikhtebaar(response.getResponseText().valueOf());
    }
  }
}
function onEdit(e) {
  var cell = e.range;
  var spreadsheet = e.source;
  // Cell if edited...
  // Mulahazaat
  if (spreadsheet.getActiveSheet().getName() == "Ikhtebaar al Ajza" && cell.getA1Notation() === "U41") {
    MulahazaatCreator();
  }
  // Auto MJ Tracker
  if (spreadsheet.getActiveSheet().getName() == "Murajaat-Jadeed" && cell.getColumn() == 2.0) {
    SpreadsheetApp.getActiveSheet().getRange("AB11:AC11").setValue(e.value);
  }
}
