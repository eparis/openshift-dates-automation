function unHideAllRowsAndColumns(sched) {
  var fullSheetRange = sched.getRange(1,1,sched.getMaxRows(), sched.getMaxColumns() )  
  sched.unhideColumn( fullSheetRange );
  sched.unhideRow( fullSheetRange ) ;   
}

function removeTrailingRows(ss){
  var maxRows = ss.getMaxRows(); 
  var lastRow = ss.getLastRow();
  var toDelete = maxRows - lastRow
  if (toDelete != 0)
    ss.deleteRows(lastRow+1, maxRows-lastRow);
}

function clearEmptyRows(ss) {
  var values = ss.getDataRange().getValues();
  nextLine: for( var i = values.length-1; i >=0; i-- ) {
    for( var j = 0; j < values[i].length; j++ )
      if( values[i][j] != "" )
        continue nextLine;
    ss.deleteRow(i+1);
  }
}

function getDateData(sourceSheet) {
 var range = sourceSheet.getRange('A2:A');
 return range.getValues();
}

function importData(data, column, sched) {
 sched.getRange(sched.getLastRow()+1, column, data.length, 1).setValues(data);
}

function setColumnCount(wanted, sched) {
 var lastColumn = sched.getMaxColumns()
 needed = wanted - lastColumn
 if (needed > 0) {
   sched.insertColumnsAfter(sched.getLastColumn(),needed);
 } else if (needed < 0) {
   sched.deleteColumns(wanted+1, needed*-1)
 }
}

function createHeader(sched) {
 sched.insertRowsBefore(1,2);
 var header = [
  ["DO NOT EDIT", "AUTOMATICALLY GENERATED", "", "", ""],
  ["Sprints", "Events", "Start Date", "End Date", "Notes"]
];
 range = sched.getRange("A1:E2");
 range.setValues(header);
 range.setFontWeight("bold");
 sched.setFrozenRows(2);
}

function doVLOOKUPs(sched) {
 var formulas = [
  ["=IFERROR(VLOOKUP(B3,'Important Dates'!A:B,2,false),VLOOKUP(A3,'Sprint Dates'!A:B,2,false))", "=IFERROR(VLOOKUP(B3,'Important Dates'!A:C,3,false),VLOOKUP(A3,'Sprint Dates'!A:C,3,false))", "=IFERROR(VLOOKUP(B3,'Important Dates'!A:D,4,false),VLOOKUP(A3,'Sprint Dates'!A:D,4,false))"],
];
  
 range = sched.getRange("C3:E3");
 range.setFormulas(formulas);
 range.setNumberFormat("mmm d, yyyy");
 lastRow = sched.getLastRow()
 for( var i = 4; i <= lastRow; i++ ) {
   var r = "C"+i+":D"+i
   dRange = sched.getRange(r);
   range.copyTo(dRange)
 }
 SpreadsheetApp.flush();
}

function applyConditionalFormating(sched) {
 var range = sched.getRange('B:B');
 var rules = [];
 
 var rule = SpreadsheetApp.newConditionalFormatRule()
     .whenTextContains("MISS")
     .setBackground("#F4C7C3")
     .setRanges([range])
     .build()
 rules.push(rule);
 
 rule = SpreadsheetApp.newConditionalFormatRule()
     .whenFormulaSatisfied('=search("OpenShift",B1)*search("GA",B1)')
     .setBackground("#B7E1CD")
     .setRanges([range])
     .build()
 rules.push(rule)
 
 rule = SpreadsheetApp.newConditionalFormatRule()
     .whenFormulaSatisfied('=search("Kube",B1)*search("GA",B1)')
     .setBackground("#FCE8B2")
     .setRanges([range])
     .build()
 rules.push(rule)
 
 sched.setConditionalFormatRules(rules);
}

function updateScheduleTab(sched, importantDateData, sprintDateData) {
  sched.clear()
  importData(importantDateData, 2, sched);
  importData(sprintDateData, 1, sched);
  
  clearEmptyRows(sched);
 
  setColumnCount(5, sched);
 
  createHeader(sched);
 
  doVLOOKUPs(sched); 
 
  removeTrailingRows(sched);
  applyConditionalFormating(sched);
 
  range = sched.getRange("A3:E");
  range.sort([{column: 3}, {column: 2}]);
 
  sched.autoResizeColumns(1,5);
}
