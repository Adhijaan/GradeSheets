/**
 * Once an update is pushed, paste it in a dated comment below for log purposes
 */

const DB_Url = process.env.MASTER_DB_URL;

function getStudentData() {
  let studentDB = SpreadsheetApp.openById(DB_Url).getSheetByName("Student DB");
  return studentDB.getDataRange();
}
function pushUpdate() {
  let studentData = getStudentData().getValues();
  for (let i = 1; i < studentData.length; i++) {
    console.info(`Student: ${studentData[i][1]} ITS: ${studentData[i][0]}`);
    try {
      let studentSpreadsheet = SpreadsheetApp.openByUrl(studentData[i][3]);
      newUpdate(studentSpreadsheet);
      console.info("Updated");
    } catch (e) {
      console.error(e);
    }
  }
  console.log("Updating complete");
}

function newUpdate(spreadsheet) {
  // Don't forget to specify the correct sheet to edit within the spreadsheet
  // Ex. spreadsheet.setActiveSheet(studentSpreadsheet.getSheetByName(MJ_SHEET_NAME));
}

// Log (ascending order)

// Hidden from GitHub
