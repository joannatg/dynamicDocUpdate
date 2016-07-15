

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Functions')
      .addItem('Update Document', 'importDocument')
      .addToUi();
}


function importDocument(){
  var year = getCurrentYear();
  var month = getCurrentMonth();
  var sourceID =  "zzz"; 
  var sourceDoc = SpreadsheetApp.openById(sourceID);    
  var targetTab = "Name of Target Document";  
  var targetID = "yyy";
  var targetDoc = SpreadsheetApp.openById(targetID); 
  var targetSheet = targetDoc.getSheetByName(targetTab);
  if (targetSheet != null) {
     Logger.log("targetSheet=" + targetSheet.getName());
  }
  
  targetSheet.insertColumns(1,1);
  
  targetSheet.setColumnWidth(1, 75);
  targetSheet.setColumnWidth(2, 350);
  targetSheet.setColumnWidth(3, 300);
  targetSheet.setColumnWidth(4, 100);
  targetSheet.setColumnWidth(5, 200);
  targetSheet.setColumnWidth(6, 100);
  targetSheet.setColumnWidth(7, 200);
  targetSheet.setColumnWidth(8, 250);
  targetSheet.setColumnWidth(9, 100);  
  
  targetSheet.clear();   
  var startTargetRow = 1;
  var firstRowSource = 1;
  
  // Import sourceTab "month 1 of year YYYY"
  // -------------------------------------------------------- 
  var currentTabName = month + " " + year;
  
  startTargetRow = importDataForMonth(sourceDoc, currentTabName, firstRowSource, targetSheet, startTargetRow);
  currentTabName = getNextTabName(currentTabName);     
  
  // Import sourceTab "month 2 of year YYYY"
  // --------------------------------------------------------  
  startTargetRow = importDataForMonth(sourceDoc, currentTabName, firstRowSource, targetSheet, startTargetRow);
  currentTabName = getNextTabName(currentTabName);
   
  // Import sourceTab "month 3 of year YYYY"
  // --------------------------------------------------------  
  startTargetRow = importDataForMonth(sourceDoc, currentTabName, firstRowSource, targetSheet, startTargetRow); 
    
 }


function getCurrentYear(){
  var currentDate = new Date();
  var year = currentDate.getYear(); 
  
  return year;
}


function getCurrentMonth(){
  var currentDate = new Date();
  var currentMonthAsNumber = currentDate.getMonth()+1;
  var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  var month = months[currentMonthAsNumber-1];  
  
  return month;
}


function buildImportQuery(sourceID, sourceTab, firstRowSource, lastRowSource) {
  
  var ff = "=(query(IMPORTRANGE(\"" + sourceID + "\", \"" + sourceTab + "!A" + firstRowSource + ":R" + lastRowSource + "\"),";
  ff += "\"" + "select Col1, Col11, Col7, Col18, Col5, Col12, Col16, Col2 where not( Col1 starts with  'IS') and not (Col1 contains '/2016') and not (Col1 contains '/2017') and not (Col1 contains '/2018') and not (Col1 contains 'Theme') and not ( Col3 starts with 'Cancelled') and not ( Col3 starts with 'Cancelled')\"";
  ff += ",1))"; 
  
  return ff;
}

//
//
function getNextTabName(currentTabName){  
  var nextTabName = "";
  var nextMonth = "";  
   
  var mmyy = currentTabName.split(" ");
  var monthAsWord = mmyy[0];
  var yearFromTabName = mmyy[1];
   
  if (monthAsWord == "December") {
    var nextMonth = "January";
    var nextYear = parseInt(yearFromTabName) + 1;
    nextTabName = nextMonth + " " + nextYear;
  }
  else {
    // get next month name         
    var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    var currentMonthAsNumber = months.indexOf(monthAsWord)+1;
    var nextMonth = months[months.indexOf(monthAsWord) + 1];
    nextTabName = nextMonth + " " + yearFromTabName;   
  }
  
  return nextTabName;
}
 
// 
// startTargetRow = importDataForMonth(sourceDoc, sourceTabNName, firstRowSource, targetSheet, startTargetRow)
//
function importDataForMonth(sourceDoc, sourceTabName, firstRowSource, targetSheet, startTargetRow) {    
  
  var sourceSheet = sourceDoc.getSheetByName(sourceTabName);
  if (sourceSheet != null) {
     Logger.log("sourceSheet=" + sourceSheet.getName());
    
  }   
  var sourceID = sourceDoc.getId();
  var lastRowSource = sourceSheet.getLastRow();  
  
  var ff = buildImportQuery(sourceID, sourceTabName, firstRowSource, lastRowSource); 
  
  targetSheet.getRange(startTargetRow,2).setValue(ff);
  targetSheet.getRange(startTargetRow,2,1,8).setBackground('#0059b9');
  targetSheet.getRange(startTargetRow,2).setFontColor("white");
  targetSheet.getRange(startTargetRow +1,2,1,8).setFontWeight("bold");  
  
  var range = targetSheet.getDataRange();  
  var lastRowTarget = targetSheet.getLastRow();       
  var borderRange = targetSheet.getRange(startTargetRow+1, 2, lastRowTarget-startTargetRow, 8);
  
  borderRange.setBorder(true, true, true, true, true, true);
  borderRange.setWrap(true);    
  startTargetRow = lastRowTarget+2; 
  
  return startTargetRow;
  
}
