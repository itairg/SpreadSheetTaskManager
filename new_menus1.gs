function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Tasks')
      .addItem('Add Task', 'add_task_func')
      .addSubMenu(ui.createMenu('Postpone')
          .addItem('To Tomorrow','postpone1')
          .addItem('To Next Week','postpone2')
          .addItem('To New Deadline','postpone3'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Other')
          .addItem('Test This Function','menuItem2')
          .addItem('Test This Function','sendEmails2'))
      .addToUi();
}

function add_task_func() {
  SpreadsheetApp.getUi() 
  
  // gets task information from user
  var task_text = Browser.inputBox('Describe Task');
  var task_deadline = Browser.inputBox('When is the task due?');
  
  // sets useful variables
  //var task_deadline = task_deadline.lower();
  var row_to_add = last_row_A();
  
  // copies the formatting
  var source = SpreadsheetApp.getActiveSheet();
  var range = source.getRange(row_to_add,1,1,4);
  var id1 = range.getGridId();
  
  // add task number to new task and format the new line
  var last_num = SpreadsheetApp.getActiveSheet().getRange(row_to_add,1).getValue();
  var last_num = +last_num+1
  var row_to_add = row_to_add+1  
  SpreadsheetApp.getActiveSheet().getRange(row_to_add,1).setValue(last_num);
  range.copyFormatToRange(id1,1,4,row_to_add,row_to_add);
  
  // add task text to new task
  SpreadsheetApp.getActiveSheet().getRange(row_to_add,2).setValue(task_text);
  
  // add "to do" text to new task
  SpreadsheetApp.getActiveSheet().getRange(row_to_add,3).setValue("לביצוע");
  
  // add task deadline, if exists, to new task
  SpreadsheetApp.getActiveSheet().getRange(row_to_add,4).setValue(get_deadline(task_deadline));
}

function get_deadline(item01) {
switch(item01) {
  case "today":
      var d= new Date();
      break;
  case "Today":
      var d= new Date();
      break;
  case "tomorrow":
      var d= new Date();
      var d_day = d.getDate();
      d.setDate(+d_day+1);
      break;
  case "Tomorrow":
      var d= new Date();
      var d_day = d.getDate();
      d.setDate(+d_day+1);
      break;
    case "0":
      var d = null;
      break;
    default:
      var d = new Date(item01);
      var d_day = d.getDate();
      d.setDate(+d_day+1);
  }
  return d
  } 


function last_row_A() {
  // will return the row number of the last row in Column A that has text, providing there are no spaces
  var Avals = SpreadsheetApp.getActiveSheet().getRange("A1:A").getValues();
  var Alast = Avals.filter(String).length;
  return Alast;
}

function menuItem2() {
  // The code below will show the name of the active sheet.
  Browser.msgBox('Active cell: ' + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getA1Notation());
}


function postpone1(){
  var tt = SpreadsheetApp.getActiveSheet().getActiveCell().getA1Notation();
  var ss = SpreadsheetApp.getActiveSheet().getActiveCell().getValue();
  var dd = new Date(ss);
  var dd_day = dd.getDate();
  dd.setDate(+dd_day+1);
  SpreadsheetApp.getActiveSheet().getRange(tt).setValue(dd);

}

function postpone2(){
  var tt = SpreadsheetApp.getActiveSheet().getActiveCell().getA1Notation();
  var ss = SpreadsheetApp.getActiveSheet().getActiveCell().getValue();
  var dd = new Date(ss);
  var dday = dd.getDay();
  var dd_day = dd.getDate();
  dd.setDate(+dd_day+7-dday);
  SpreadsheetApp.getActiveSheet().getRange(tt).setValue(dd);
}

function postpone3(){
  var new_deadline = Browser.inputBox('When is the task due, now?');
  var tt = SpreadsheetApp.getActiveSheet().getActiveCell().getA1Notation();
  SpreadsheetApp.getActiveSheet().getRange(tt).setValue(get_deadline(new_deadline));
}

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = 2;   // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 2)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (i in data) {
    var row = data[i];
    var emailAddress = row[0];  // First column
    var message = row[1];       // Second column
    var subject = "Sending emails from a Spreadsheet";
    MailApp.sendEmail(emailAddress, subject, message);
  }
}
// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = "EMAIL_SENT";

function sendEmails2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = 2;   // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 3)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0];  // First column
    var message = row[1];       // Second column
    var emailSent = row[2];     // Third column
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
      var subject = "Sending emails from a Spreadsheet";
      MailApp.sendEmail(emailAddress, subject, message);
      sheet.getRange(startRow + i, 3).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
