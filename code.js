var aosReleaseDatesSpreadSheetID = ""
var aosMainCalendarEventKey = "AUTOMATED EVENT! Please Update via OpenShift Release Dates Spreadsheet: https://docs.google.com/spreadsheets/d/";
var aosMainCalendarID = ""
var scheduleSheet = "Schedule";
var importantDatesSheet = "Important Dates";
var sprintDatesSheet = "Sprint Dates";

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Generate Schedule')
      .addItem('Regenerate', 'myFunction')
      .addToUi();
  
  jumpToDate();
}

function jumpToDate() {

  // This call fails
  // var ss = SpreadsheetApp.openById(aosReleaseDatesSpreadSheetID);
 var ss = SpreadsheetApp.getActiveSpreadsheet();

 var sheet = ss.getActiveSheet();
 // Only scroll if it is first tab which is the schedule.
 if (sheet.getName() != scheduleSheet) return;
// Logger.log("in jumpToDate()");
 var range = sheet.getRange("C:C");
 var values = range.getValues();  
 var day = 24*3600*1000;  
 var today = parseInt((new Date().setHours(0,0,0,0))/day);  
 var ssdate; 
 for (var i=0; i<values.length; i++) {
   try {
     ssdate = values[i][0].getTime()/day;
   }
   catch(e) {
   }
   if (ssdate && Math.floor(ssdate) >= today) {
//     Logger.log('Jump to row: %s', i);
     sheet.setActiveRange(range.offset(i,0,1,1));
     break;
   }    
 }
}

function myFunction() {
  var ss = SpreadsheetApp.openById(aosReleaseDatesSpreadSheetID);
  var sched = ss.getSheetByName(scheduleSheet);
  unHideAllRowsAndColumns(sched);
 
  var importantDates = ss.getSheetByName(importantDatesSheet);
  var importantDateData = getDateData(importantDates);
  var sprintDates = ss.getSheetByName(sprintDatesSheet);
  var sprintDateData = getDateData(sprintDates);
 
  updateScheduleTab(sched, importantDateData, sprintDateData); 
  updateAOSMainCalendar(sched) 
}
