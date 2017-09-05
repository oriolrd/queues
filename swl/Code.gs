var sheets = ["Pending Cases","PrioIPM","sortedBacklog","LongRunner","Sharing","Closed","TAM (Team availability matrix)","General information","Config"];
var col_Aurora = 1;
var col_pripsi = 2;
var col_GCH = 3;
var col_doc = 4;
var col_caseStat = 5;
var col_swname = 6;
var col_CIRstat = 7;
var col_casekey = 8;
var col_comm = 9;
var col_probdesc = 10;
var col_IPM = 11;
var col_TAT = 12;
var col_OpenDays = 13;
var col_prodnam = 14;
var col_vers=15;
var col_SVC = 16;
var col_PO = 17;
var col_CIRrisk = 18;
var col_fstcont = 19;
var col_creDat = 20;
var col_prodAss = 21;
var col_PRISMA = 22;
var col_country = 23;
var col_customer = 24;
var col_PRApproved = 25;
var col_daysPRA = 26;
var email = Session.getActiveUser().getEmail();

function onEdit(e){
  var lock = LockService.getPublicLock();
  lock.waitLock(10000);
  var s = e.range.getSheet();
  var active_cell = s.getActiveCell();
  var active_column = e.range.getColumnIndex();
  var active_sheet = s.getName();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source_sheet, target_sheet = "";
  
  try{
    if ((active_sheet == sheets[0] || active_sheet == sheets[5]) && e.changeType != "FORMAT") {
      var r = e.source.getActiveRange();
      var row = r.getRow();
      var status = s.getRange(row, col_caseStat).getValue();
      var pri_psi = s.getRange(row, col_pripsi).getValue(), caso = s.getRange(row, col_Aurora).getValue();
      if (active_column == col_caseStat) {      
        if (row != 1) {     
          // If Closed sheet and the status is wrong, move to Pending tab
          // Else If Pending tab and the status is closed, move to the Closed tab
          if (active_sheet == sheets[5] && (status == "Evaluation in Progress" || status == "Investigation In Progress" || status == "Submitted for Peer Review" || status == "Case Closure and Summary")){
            source_sheet = ss.getSheetByName(sheets[5]);
            target_sheet = ss.getSheetByName(sheets[0]);
            moveCase(caso,source_sheet,target_sheet);
            var msg = "The case "+caso+" has been moved to Pending tab";
            ss.toast('NOTE',msg,6);
            lock.releaseLock();
          } else if (active_sheet == sheets[0] && (status == "Case Closed")){
            source_sheet = ss.getSheetByName(sheets[0]);
            target_sheet = ss.getSheetByName(sheets[5]);
            moveCase(caso,source_sheet,target_sheet);
            var msg = "The case "+caso+" has been moved to Closed tab";
            ss.toast('NOTE',msg,6);
            lock.releaseLock();
          }          
        }
      }
      if (active_column == col_Aurora){
        addAllLinks(s.getRange(row, col_Aurora));
      }
      
      if (active_column == col_pripsi) {      
        if (row != 1) {
          s.getRange(row,col_pripsi).setBackground('red');
        }
      }
    }
  } catch (error) { Logger.log(error)}
}




function moveCase(caso,source,dest){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var target_sheet = dest;
  var msg = "";
  var currentrow = findCase(caso,source);
  if (currentrow != 0){ //If exists in current sheet
    if (target_sheet.getName() == sheets[5]) { //Move to Closed sheet
      var source_range = source.getRange(currentrow,1,1,col_lastUser);
      var tat = source.getRange(currentrow, col_TAT); 
      target_sheet.insertRowsAfter(1, 1);
      var target_range = target_sheet.getRange(2,1,1,col_lastUser);
      var target_tat = target_sheet.getRange(2,col_TAT);
      source_range.copyTo(target_range,{contentsOnly: false});
      tat.copyTo(target_tat,{contentsOnly: true});
      target_sheet.getRange(2,24).setFormula("=L2+R2"); //Closed Date calc
      target_sheet.getRange(2,25).setValue(Session.getActiveUser().getEmail());
      
      //Removing format for accepted cases
      target_sheet.getRange("A2:Z").clearFormat().setVerticalAlignment("top").setWrap(true);
      target_sheet.getRange("A1:Y").setHorizontalAlignment("center");
      target_sheet.getRange("Y1:Z").setHorizontalAlignment("left");
      
    } else {
      
      var source_range = source.getRange(currentrow,1,1,col_lastUser);
      var nextAvailableRow = target_sheet.getLastRow()+1;
      target_sheet.insertRowsAfter(nextAvailableRow-1, 1);
      var target_range = target_sheet.getRange(nextAvailableRow,1,1,col_lastUser);
      source_range.copyTo(target_range,{contentsOnly: false});
      
      // Sorting after moving the row
      sortSheet(target_sheet.getName(),caso,0);
    }
    
    msg = "Case "+caso+" moved to the "+target_sheet.getName()+" sheet";
    Logger.log(msg);
    
    //Borrar fila de hoja origen
    source.deleteRows(currentrow);
    ss.toast('INFO', msg, 4); // Show message after moving
    
    /*if (cambiar_hoja == 1) {
    setFocus(target_sheet.getName(),caso,11);
    }*/
  }
}



function findCase(id, sheet){
  Logger.log("Buscando fila real para caso "+id);
  var hoja = sheet.getName();
  Logger.log(sheet + " " + hoja);
  var columnValues = "";
  columnValues = sheet.getRange(2, col_Aurora, sheet.getLastRow()).getValues();
  var len = columnValues.length;
  
  for(var j = 0 ;  j < columnValues.length ; j ++ ) {
    if(columnValues[j][0] == id) {
      var new_row = j + 2;
      
      Logger.log("Found ["+id+"] at row "+new_row+" from sheet "+hoja)
      return new_row;
    }
  }
  Logger.log("Case not found");
  return 0; //no se ha encontrado el caso
}



function sortSheet(sheet, oldcase, origin){
  if (sheet == null) sheet=SpreadsheetApp.getActiveSheet().getName(); //If not getting the sheet name, taking the actual sheet
  if (origin == null) origin=0;
  if (sheet == sheets[0] ){ // Only sorting Pending Cases tab 
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var target_sheet = ss.getSheetByName(sheet);
    var s = ss.getActiveSheet();
    var lastCol = target_sheet.getLastColumn();
    var lastRow = target_sheet.getLastRow();
    var r = target_sheet.getRange(2, 1, lastRow-1, lastCol);
    r.setVerticalAlignment("middle").setWrap(true);
    r.sort([{ column: 1, ascending: true }]);
    target_sheet.getRange("P2:P").setFontSize(9);
    Logger.log("Sorted sheet "+sheet);
  }
}



function addLink(celda){
  var formula = celda.getFormula();
  var i = celda.getActiveRow();
  var a =2;
  if (formula[0] != "=") {
    var caso = celda.getValue();
    var url = "https://rexis.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=a30&str=" + caso;
    celda.setFormula('=HYPERLINK("' + url + '","' + caso + '")');
    celda.setFontFamily("Verdana");
    celda.setFontSize(9);
    celda.setHorizontalAlignment("center");
  }
}


function addAllLinks(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pend = ss.getSheetByName(sheets[0]);
  var cases = pend.getRange(2, col_Aurora, pend.getLastRow()).getValues();
  var pripsis = pend.getRange(2, col_pripsi, pend.getLastRow()).getValues();
  for(var j = 0 ;  j < cases.length -1; j ++ ) {
    // Looking for Case Number in the Aurora column
    var i = j+2;
    var casenum = pend.getRange(j+2, col_Aurora).getValue();

    // Aurora Case
    var formula = pend.getRange(j+2, col_Aurora).getFormula();
    if (formula[0] != "=") {
      var column = col_Aurora;
      var url = "https://rexis.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=a30&str=" + casenum;
      pend.getRange(i, column).setFormula('=HYPERLINK("' + url + '","' + casenum + '")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }
      
    //PRIPSI column
    var formula = pend.getRange(j+2, col_pripsi).getFormula();
    if (formula[0] != "=") {
      var column = col_pripsi;
      pend.getRange(i, column).setFormula('=iferror(if(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Is PRI?",allCases!$1:$1,0))="No","N",if(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Is PRI?",allCases!$1:$1,0))="Yes","Y","")),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }
      
    //GCH column
    var formula = pend.getRange(j+2, col_GCH).getFormula();
    if (formula[0] != "=") {
      var column = col_GCH;
      pend.getRange(i, column).setFormula('=iferror(left(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Global Case Handler",allCases!$1:$1,0)),find(" ",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Global Case Handler",allCases!$1:$1,0)))-1),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }
      
    //Case Status column
    var formula = pend.getRange(j+2, col_caseStat).getFormula();
    if (formula[0] != "=") {
      var column = col_caseStat;
      pend.getRange(i, column).setFormula('=iferror(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Workflow Status",allCases!$1:$1,0)),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }
      
    //Software Name column
    var formula = pend.getRange(j+2, col_swname).getFormula();
    if (formula[0] != "=") {
      var column = col_swname;
      pend.getRange(i, column).setFormula('=iferror(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Software Name",allCases!$1:$1,0)),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }

    //CIR Status column
    var formula = pend.getRange(j+2, col_CIRstat).getFormula();
    if (formula[0] != "=") {
      var column = col_CIRstat;
      pend.getRange(i, column).setFormula('=iferror(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Lot Number",allCases!$1:$1,0)),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }

    //Cause Keywords column
    var formula = pend.getRange(j+2, col_casekey).getFormula();
    if (formula[0] != "=") {
      var column = col_casekey;
      pend.getRange(i, column).setFormula('=iferror(if(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Cause Keywords",allCases!$1:$1,0))="","",hyperlink("https://jira.swlonline.de/browse/"&index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Cause Keywords",allCases!$1:$1,0)),index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Cause Keywords",allCases!$1:$1,0)))),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }

    //Problem Description column
    var formula = pend.getRange(j+2, col_probdesc).getFormula();
    if (formula[0] != "=") {
      var column = col_probdesc;
      pend.getRange(i, column).setFormula('=iferror(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Problem Description",allCases!$1:$1,0)),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }

    //TAT column
    var formula = pend.getRange(j+2, col_TAT).getFormula();
    if (formula[0] != "=") {
      var column = col_TAT;
      pend.getRange(i, column).setFormula('=IF(S'+i+'="",,IFERROR(DATEDIF(S'+i+',TODAY(),"D"),))');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }

    //Open Days column
    var formula = pend.getRange(j+2, col_OpenDays).getFormula();
    if (formula[0] != "=") {
      var column = col_OpenDays;
      pend.getRange(i, column).setFormula('=IF(T'+i+'="",,IFERROR(DATEDIF(T'+i+',TODAY(),"D"),))');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }

    //Product Name column
    var formula = pend.getRange(j+2, col_prodnam).getFormula();
    if (formula[0] != "=") {
      var column = col_prodnam;
      pend.getRange(i, column).setFormula('=iferror(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Product Name",allCases!$1:$1,0)),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }

    //Software Version column
    var formula = pend.getRange(j+2, col_vers).getFormula();
    if (formula[0] != "=") {
      var column = col_vers;
      pend.getRange(i, column).setFormula('=iferror(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Software Version",allCases!$1:$1,0)),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }

    //First Contact column
    var formula = pend.getRange(j+2, col_fstcont).getFormula();
    if (formula[0] != "=") {
      var column = col_fstcont;
      pend.getRange(i, column).setFormula('=iferror(date(right(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0)),4),left(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0)),find("/",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0)))-1),MID(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0)),find("/",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0)))+1,len(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0)))-5-find("/",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0))))),date(right(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0)),4),MID(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0)),find(".",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0)))+1,len(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0)))-5-find(".",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0)))),left(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0)),find(".",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Date of First Contact",allCases!$1:$1,0)))-1)))');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }

    //Created Date column
    var formula = pend.getRange(j+2, col_creDat).getFormula();
    if (formula[0] != "=") {
      var column = col_creDat;
      pend.getRange(i, column).setFormula('=iferror(date(right(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0)),4),left(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0)),find("/",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0)))-1),MID(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0)),find("/",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0)))+1,len(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0)))-5-find("/",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0))))),date(right(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0)),4),MID(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0)),find(".",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0)))+1,len(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0)))-5-find(".",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0)))),left(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0)),find(".",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Escalated Issue: Created Date",allCases!$1:$1,0)))-1)))');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }

    //Product Assessment column
    var formula = pend.getRange(j+2, col_prodAss).getFormula();
    if (formula[0] != "=") {
      var column = col_prodAss;
      pend.getRange(i, column).setFormula('=iferror(left(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Product Assessment(Com. Substantiated)",allCases!$1:$1,0)),find(" / ",index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Product Assessment(Com. Substantiated)",allCases!$1:$1,0)))-1),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }

    //Prisma ID column
    var formula = pend.getRange(j+2, col_PRISMA).getFormula();
    if (formula[0] != "=") {
      var column = col_PRISMA;
      pend.getRange(i, column).setFormula('=iferror(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Prisma Case ID",allCases!$1:$1,0)),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }

    //Country column
    var formula = pend.getRange(j+2, col_country).getFormula();
    if (formula[0] != "=") {
      var column = col_country;
      pend.getRange(i, column).setFormula('=iferror(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Prisma Country of Origin",allCases!$1:$1,0)),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }

    //Customer column
    var formula = pend.getRange(j+2, col_customer).getFormula();
    if (formula[0] != "=") {
      var column = col_customer;
      pend.getRange(i, column).setFormula('=iferror(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Prisma Customer ID",allCases!$1:$1,0)),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");      
      
    }

    //Days since Peer Review Approved
    var formula = pend.getRange(j+2, col_daysPRA).getFormula();
    if (formula[0] != "=") {
      var column = col_daysPRA;
      pend.getRange(i, column).setFormula('=if(Y'+i+'<>"", iferror(datedif(Y'+i+',today(),"D"),""),"")');
      pend.getRange(i, column).setFontFamily("Verdana");
      pend.getRange(i, column).setFontSize(9);
      pend.getRange(i, column).setHorizontalAlignment("center");
    }
    
    // If a case is closed, renaming the formula in order to move it to Closed cases tab
    var lock = LockService.getPublicLock();
    var column = col_caseStat;
    var caseStatus = pend.getRange(i, column).getValue();
    if (caseStatus == "Case Closed") {
      //pend.getRange(i, column).setFormula('=iferror(index(allCases!$A:$S,match($A'+i+',allCases!$A:$A,0),match("Workflow Status",allCases!$1:$1,0))," ")');
      var caso = pend.getRange(i, col_Aurora).getValue();
      var source_sheet = ss.getSheetByName(sheets[0]);
      var target_sheet = ss.getSheetByName(sheets[5]);
      moveCase(caso,source_sheet,target_sheet);
      var msg = "The case "+caso+" has been moved to Closed tab";
      ss.toast('NOTE',msg,6);
      lock.releaseLock();
    }
    
    //When a Peer Review has been Approved, add the date to the column 'Peer Review approved date'.
    //It takes the date of the day when the list was updated with that status.
    if (caseStatus == "Peer Review Approved") {
      var formula = pend.getRange(j+2,col_PRApproved).getFormula();
      if (formula[0] != "=") {
        // Aurora Case
        var d = new Date();
        var timeStamp = d.getTime();
        //The timeStamp is in Epoch format. Needs conversion
        var column = col_PRApproved;
        pend.getRange(i, column).setFormula('=(('+timeStamp+'-21600000)/86400000)+25569');
        pend.getRange(i, column).setFontFamily("Verdana");
        pend.getRange(i, column).setFontSize(9);
        pend.getRange(i, column).setHorizontalAlignment("center"); 
      }
    }
  }
}



function moveWrongCasesfromClosed(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pend = ss.getSheetByName(sheets[0]);
  var closed = ss.getSheetByName(sheets[5]);
  var casestat = closed.getRange(2, col_caseStat, closed.getLastRow()).getValues();
  // loop in reverse sort so the rows deleted don't affect in the count
  for(var j = casestat.length -1 ;  j > 0 ; j -- ) {
    var caso = closed.getRange(j+1, col_caseStat).getValue();
    var casenumb = closed.getRange(j+1, col_Aurora).getValue();
    if ((casestat[j-1] == "Evaluation in Progress") || (casestat[j-1][1] == "Investigation in Progress")) {
      moveCase(casenumb,closed,pend);
      Utilities.sleep(3000);
    }
  }
}



function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CIR Utilities')
  .addItem('Sort sheet', 'sortSheet')
  .addItem('Add links', 'addAllLinks')
  //.addItem('Wrong closed cases', 'moveWrongCasesfromClosed')
  .addToUi();
}