// Manually inputted spreadsheet links.
const StudentDB = process.env.STUDENT_DB;
var studentsSheets = [StudentDB, "Murajaat-Jadeed"];

// modified updater
function UpdateSheets() {
  for (var i = 0; i < studentsSheets.length; i++) {
    var spreadsheet = SpreadsheetApp.openByUrl(studentsSheets[i]);
    console.log(spreadsheet.getName());
    console.log(spreadsheet.getUrl());
    NewUpdate(spreadsheet);
  }
}
function NewUpdate(spreadsheet) {
  // DON'T FORGET to specify sheet name
  //ex. spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Murajaat-Jadeed'));
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Murajaat-Jadeed"));

  // Rest of Update Maro Here
}
