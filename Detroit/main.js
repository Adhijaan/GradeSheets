// If you accidently landed on this file, please don't edit anything. - Spreadsheet Dev Team
/**
 * @NotOnlyCurrentDoc
 */
const MJ_SHEET_NAME = "Murajaat-Jadeed";
const JH_SHEET_NAME = "Juz Haali";

function newEntry() {
  SpreadsheetApp.getActive().toast("Creating", "Status");

  // Checks to see if an entry is already created for the current day
  if (doubleEntryCheck()) return;

  // Determines Marhalah number - if unable, throws an error and halts entry creation
  let marhalah = determineMarhalah();

  try {
    //Send yesterday's marks to the database
    if (SpreadsheetApp.getActive().getSheetByName(MJ_SHEET_NAME).getRange("P3").getValue() <= 1000000) {
      SpreadsheetApp.getActive().toast("ITS missing/invalid. Cannot create entry.", "Status");
      return;
    }
    newDatabaseEntry();
  } catch (e) {
    console.error(e);
    SpreadsheetApp.getActive().toast("Data Upload Issue", "Status");
  }
  // Add a new month banner if needed
  let newMonth = isNewMonth();
  if (newMonth) NewMonth();

  //Create Juz Haali Entry
  jhNewEntry();
  jhDate();
  jhMarhalahFormulas(marhalah);
  jhFormat();

  //Create the MJ-JD entry
  mjjdNewEntry();
  mjjdDate();
  mjjdBaseFormulas();
  mjjdMarhalahFormulas(marhalah);
  mjjdFormat();

  //Re-paste MJ tracker to correct position.
  pasteMJTracker(newMonth);

  //uncheck newEntry button
  SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActive().getSheetByName(MJ_SHEET_NAME)).getRange("E5:F5").uncheck();
  SpreadsheetApp.getActive().toast("Complete", "Status", 2);
}

function determineMarhalah() {
  let spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(MJ_SHEET_NAME);

  // Hardcode jadeed value
  sheet.getRange("V8:V10").copyTo(sheet.getRange("V8:V10"), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  // Update Makaan al Haali if applicable
  const jdSurah = sheet.getRange("R8:R10").getValue(); //latest jadeed surah cell
  const jdAyah = sheet.getRange("S8:S10").getValue(); //latest jadeed ayah cell

  //If the last entry has jadeed filled, update Makan al Hali
  if (jdSurah != "" && jdAyah != "") {
    //set cell S5, the surah of Makaan al Haali to the latest jadeed surah.
    sheet.getRange("S5").setValue(jdSurah.toString().trim());
    //set cell V5, the ayah of Makaan al Haali to the latest jadeed ayat.
    sheet.getRange("V5").setValue(jdAyah.toString().trim());
  }

  // Use Makaan al Haali page number (X5) to determine marhalah

  const makaan_al_haali_page = sheet.getRange("X5").getValue();
  if (typeof makaan_al_haali_page !== "number") {
    spreadsheet.toast("Last Jadeed entry has invalid surah name or ayat number.", "Error", 10);
    throw new Error("Page Query for page failed. Last jadeed entry must have invalid values");
  }
  const marhalahs = ["جزء ثلاثون", "جزء اول - ثالث", "جزء رابع - خامس", "جزء خامس+"];
  const marhalah_cell = sheet.getRange("S4:V4"); //current dropdown marhalah

  if (makaan_al_haali_page > 541) {
    // Juz 28 - 30
    marhalah_cell.setValue(marhalahs[0]);
    return 1;
  } else if (makaan_al_haali_page > 101) {
    // Juz 5+
    marhalah_cell.setValue(marhalahs[3]);
    return 4;
  } else if (makaan_al_haali_page > 61) {
    // Juz 4-5
    marhalah_cell.setValue(marhalahs[2]);
    return 3;
  } else {
    // Juz 1-3
    marhalah_cell.setValue(marhalahs[1]);
    return 2;
  }
}

function isNewMonth() {
  var spreadsheet = SpreadsheetApp.getActive();
  sheet = spreadsheet.getSheetByName(MJ_SHEET_NAME);
  var startDateCell = sheet.getRange("B8").getValue(); //last entry date in English
  var endDate = new Date(); //todays english date
  var startDate = new Date(startDateCell);

  // Calculate the difference in milliseconds
  var timeDiff = endDate.getTime() - startDate.getTime();

  // Convert milliseconds to days

  var days = Math.floor(timeDiff / (1000 * 60 * 60 * 24));
  var lastEntrydate = HijriCalander(-days); //last entry hirji date array using the offset calculated above
  var hijri = HijriCalander(); // Get the array for today

  //If the hijri month of the last entry and today's entry dont match - true
  if (hijri[2] != lastEntrydate[2] || hijri[3] != lastEntrydate[3]) {
    return true;
  } else {
    return false;
  }
}
/**
 * function creates a new month header and brings down the old month's header
 */
function NewMonth() {
  let spreadsheet = SpreadsheetApp.getActive();
  const OLD_MONTH_STRING = spreadsheet.getRange("A2:Z2").getValue();
  const OLD_MONTH_COLOR = MonthColor(-1);
  const NEW_MONTH_STRING = `ربع ${convertToArabicIndic(Math.ceil(HijriCalander()[2] / 3))} - ${
    HijriCalander()[5]
  } - ${convertToArabicIndic(HijriCalander()[3])}`;

  // MJ-JD sheet
  let sheet = spreadsheet.getSheetByName(MJ_SHEET_NAME);

  // Old Month Header
  sheet.insertRowsBefore(8, 2);
  sheet
    .getRange("A8:Z9")
    .merge()
    .setFontSize(14)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setValue(OLD_MONTH_STRING)
    .setBackground(OLD_MONTH_COLOR);

  // New Month Header
  sheet.getRange("A2:Z2").setValue(NEW_MONTH_STRING);

  // JH sheet
  sheet = spreadsheet.getSheetByName(JH_SHEET_NAME);

  // Old Month Header
  sheet.insertRowsBefore(7, 2);
  sheet.setRowHeights(7, 2, 27);
  sheet
    .getRange("A7:Q8")
    .merge()
    .setFontSize(14)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setValue(OLD_MONTH_STRING)
    .setBackground(OLD_MONTH_COLOR);

  // New Month Header
  sheet.getRange("A2:Q2").setValue(NEW_MONTH_STRING);
}

function convertToArabicIndic(number) {
  let arabicIndicNum = number.toString().replace(/[0-9]/g, (digit) => {
    const indicDigits = [
      "\u0660",
      "\u0661",
      "\u0662",
      "\u0663",
      "\u0664",
      "\u0665",
      "\u0666",
      "\u0667",
      "\u0668",
      "\u0669",
    ];
    return indicDigits[digit];
  });
  return arabicIndicNum;
}

function MonthColor(monthOffset = 0) {
  const colors = [
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
    "#989898",
    "#e7d0b0",
  ];
  var monthColor = colors[(HijriCalander()[2] + monthOffset + 11) % 12];
  return monthColor;
}

/**
 * Returns an array of useful hijri calender values
 * @param {number} - The number of days to offset from current date for calander values. Defaults to current date.
 * 0 = weekday number
	 1 = islamic day of month num
	 2 = islamic month num
	 3 = islamic year num 
   4 = islamic weekday name string hijri
	 5 = islamic month name string hijri
 */
function HijriCalander(dayOffset = 0) {
  var today = new Date();
  today.setDate(today.getDate() + dayOffset);
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

  var dateinfo = new Array(6);
  dateinfo[0] = wd; //weekday number (0-6)
  dateinfo[1] = id; //islamic day of month num
  dateinfo[2] = im; //islamic month num
  dateinfo[3] = iy; //islamic year num
  dateinfo[4] = wdNames[wd]; //islamic weekday name string
  dateinfo[5] = iMonthNames[im - 1]; //islamic month name string

  return dateinfo;
}

function pasteMJTracker(newMonth) {
  var spreadsheet = SpreadsheetApp.getActive();
  if (newMonth) {
    spreadsheet.getRange("AB8:AE21").activate();
    spreadsheet.getRange("AB13:AE26").moveTo(spreadsheet.getActiveRange());
  } else {
    spreadsheet.getRange("AB8:AE21").activate();
    spreadsheet.getRange("AB11:AE24").moveTo(spreadsheet.getActiveRange());
  }
}

//Inserts one row to the top of the "Daily Hifz Records" tab in database spreadsheet for the latest class entry
function newDatabaseEntry() {
  //students sheet
  let spreadsheet = SpreadsheetApp.getActive();
  //Select the students murajaat sheet as active spreadsheet
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(MJ_SHEET_NAME));

  //find all the data from the last entry
  const databaseID = "1UoF8L8XUvPd5UCJ9g3J5GKkvwz2UxgZj1WdM4A2s3tM";
  var its = spreadsheet.getRange("P3:R3").getValue(); //ITS ID
  var date = spreadsheet.getRange("B8").getValue(); //last date
  var jdAyah = spreadsheet.getRange("S8:S10").getValue(); //last jadeed ayah cell
  var jdSurah = spreadsheet.getRange("R8:R10").getValue(); //last jadeed surah cell
  var jdLines = spreadsheet.getRange("V8:V10").getValue(); //last jadeed numLines cell
  var juz1 = spreadsheet.getRange("C8").getValue(); //last class juz1 murajaat siparo
  var juz1Marks = spreadsheet.getRange("P8").getValue(); //last class juz1 murajaat marks
  var juz2 = spreadsheet.getRange("C9").getValue(); //last class juz2 murajaat siparo
  var juz2Marks = spreadsheet.getRange("P9").getValue(); //last class juz2 murajaat marks
  var juz3 = spreadsheet.getRange("C10").getValue(); //last class juz3 murajaat
  var juz3Marks = spreadsheet.getRange("P10").getValue(); //last class juz3 murajaat marks
  var jH = spreadsheet.getRange("Q8:Q10").getValue(); //last Juz Haali marks

  //switch spreadsheet to the database, and daily hifz record tab

  spreadsheet = SpreadsheetApp.openById(databaseID).getSheetByName("Daily Murajaat Records");
  spreadsheet.insertRowBefore(2); //create a new row at the top, after the headers
  //Set all the previously set variables from latest hifz class into database
  spreadsheet.getRange("A2").setValue(its);
  spreadsheet.getRange("B2").setValue(date);
  spreadsheet.getRange("C2").setValue(juz1);
  spreadsheet.getRange("D2").setValue(juz1Marks);

  //Check if student had a 2nd murajaat siparo, if so add
  if (juz2 != "") {
    spreadsheet.insertRowBefore(2); //create a new row at the top, after the headers
    spreadsheet.getRange("A2").setValue(its);
    spreadsheet.getRange("B2").setValue(date);
    spreadsheet.getRange("C2").setValue(juz2);
    spreadsheet.getRange("D2").setValue(juz2Marks);
  }
  //Check if student had a 3rd murajaat siparo, if so add
  if (juz3 != "") {
    spreadsheet.insertRowBefore(2); //create a new row at the top, after the headers
    spreadsheet.getRange("A2").setValue(its);
    spreadsheet.getRange("B2").setValue(date);
    spreadsheet.getRange("C2").setValue(juz3);
    spreadsheet.getRange("D2").setValue(juz3Marks);
  }

  //switch spreadsheet to the database, and daily Jadeed/JH record tab
  spreadsheet = SpreadsheetApp.openById(databaseID).getSheetByName("Daily Jadeed/JH Records");
  spreadsheet.insertRowBefore(2); //create a new row at the top, after the headers
  //Set all the previously set variables from latest hifz class into database
  spreadsheet.getRange("A2").setValue(its); //ITS
  spreadsheet.getRange("B2").setValue(date); //Date
  spreadsheet.getRange("C2").setValue(jdAyah); //Jadeed Ayat
  spreadsheet.getRange("D2").setValue(jdSurah); //Jadeed Surah
  spreadsheet.getRange("E2").setValue(jdLines); //Jadeed line amount
  spreadsheet.getRange("F2").setValue(jH); //JH marks
}

function doubleEntryCheck() {
  let lastEntryDate = new Date(SpreadsheetApp.getActive().getSheetByName(MJ_SHEET_NAME).getRange("B8").getValue());
  let today = new Date();
  if (lastEntryDate.getDate() == today.getDate() && lastEntryDate.getMonth() == today.getMonth()) {
    let ui = SpreadsheetApp.getUi();
    let response = ui.alert(
      "Warning",
      "A class entry has already been created for today, are you sure you would like to make a new one?",
      ui.ButtonSet.YES_NO
    );
    if (response == ui.Button.NO) {
      SpreadsheetApp.getActive().toast("Stopped", "Status");
      return true;
    }
  }
  return false;
}
