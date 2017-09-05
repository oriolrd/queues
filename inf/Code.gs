//variables ###############################################
var sheets = ["Pending Cases","Pending Closure","Closed"];
var col_postit = 2;
var col_pripsi = 3; 
var col_prior = 4;
var col_case = 5;
var col_user = 6;
var col_tat = 7;
var col_reminder = 8;
var col_status = 9;
var col_version = 10;
var col_country = 11;
var col_module = 12;
var col_description = 13;
var col_comment = 14;
var col_customer = 15;
var col_region = 16;
var col_sw = 17;
var col_date = 18;
var col_change = 19;
var col_oldprior = 20;

//var email = Session.getActiveUser().getEmail();
var ui = SpreadsheetApp.getUi();
//#########################################################

function onEdit(e){

  //var s = e.source.getActiveSheet();
  var s = e.range.getSheet();
  var active_cell = s.getActiveCell();
  //var active_column = active_cell.getColumn();
  var active_column = e.range.getColumnIndex();
  var active_sheet = s.getName();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source_sheet, target_sheet = "";
  
  
  try{
    if ((active_sheet == sheets[0] || active_sheet == sheets[1]) && e.changeType != "FORMAT") {
      
      var r = e.source.getActiveRange();
      var row = r.getRow();
      var status = s.getRange(row, col_status).getValue();
      //var region = s.getRange(row, col_region).getValue();
      //var user = s.getRange(row, col_user).getValue();
      var pri_psi = s.getRange(row, col_pripsi).getValue(), priority = s.getRange(row, col_prior).getValue(),caso = s.getRange(row, col_case).getValue();
      
      //si se modifican las columnas PRI, PSI, Priority y Current Status entonces aplicamos script
      if (active_column == col_pripsi || active_column == col_prior || active_column == col_status) {      
        if (row != 1) {     
          if (caso != "") s.setRowHeight(row, 21); // misma altura para todas las filas
          
          
          //format priorities, no hace falta cuando cambiamos de estado
          if (active_column != col_status){
            formatPriorities(pri_psi,priority,row);
          } else {
            var now = new Date();
            Logger.log("Añadir fecha de cambio");
            s.getRange(row, col_change).setValue(now).setNumberFormat('yyyy-mm-dd');
            
            if (status == "Case Closed"){
              source_sheet = s;
              target_sheet = ss.getSheetByName(sheets[2]);
              moveCase(caso,source_sheet,target_sheet);
            }
          }
           
          //Si estamos en pestaña Pending Closure y volvemos al estado Investigation in Progress o Submitted for Peer Review, mover caso a la hoja Pending Cases
          if (active_sheet == sheets[1] && (status == "Evaluation in Progress" || status == "Investigation In Progress" || status == "Submitted for Peer Review" || status == "Case Closure and Summary")){
            source_sheet = ss.getSheetByName(sheets[1]);
            target_sheet = ss.getSheetByName(sheets[0]);
            moveCase(caso,source_sheet,target_sheet);
          } else if (active_sheet == sheets[0] && (status == "Peer Review Approved" || status == "Submitted for Affiliate Review")) {
            if (priority == 5) {
              var old_priority = s.getRange(row, col_oldprior).getValue();
              s.getRange(row, col_prior).setValue(old_priority);
            }
            source_sheet = ss.getSheetByName(sheets[0]);
            target_sheet = ss.getSheetByName(sheets[1]);
            moveCase(caso,source_sheet,target_sheet);
            var msg = "Please move case "+caso+" to the Aff.Pending column in the whiteboard";
            ss.toast('NOTE',msg,6);
          } else if (active_sheet == sheets[0] && (status == "Submitted for Peer Review") && (pri_psi == "N")) {
            s.getRange(row, col_oldprior).setValue(priority)
            s.getRange(row, col_prior).setValue("5");
            orderSheet(active_sheet,caso,0);
            var msg = "Please move case "+caso+" to the PR column in the whiteboard";
            ss.toast('NOTE',msg,6); 
          } else if (status == "Pending answer"){
            var ndate = Utilities.formatDate(new Date(), "GMT", "dd/MM");
            var msg = "Please update case "+caso+" in the whiteboard with date "+ndate;
            ss.toast('NOTE',msg,6); 
          }
        }
      }
      
      //Añadimos el link al caso de Aurora
      if (active_column == col_case){
        addLink(s.getRange(row, col_case));
      }
      
      if (active_column == col_reminder){
        addCommentReminder(active_cell);
      }
      
      if (active_column == col_pripsi || active_column == col_prior || active_column == col_date){
        var old_case = caso;
        orderSheet(active_sheet,old_case,0);
      }
      
      //volvemos a poner el foco donde estábamos después de ordenar
      //setFocus(active_sheet,old_case,old_column);
    }
  } catch (error) { Logger.log(error)}

}

//filter=num columna que se usará para filtrar para los casos que aparecen arriba
//value = valor del filtro
function freezesRows(sheet,filter,value){
  var lrow = sheet.getLastRow();
  var r = sheet.getRange(2, filter, lrow-1, 1); // lrow-1 to not count header
  var data = r.getValues();
  var count=1
  for (var i = 0; i < data.length ; i++) {
      if (data[i][0] >= value) {
        count++;
    }
  }
  sheet.setFrozenRows(count);
};

function onOpen() {
  ui.createMenu('CIR Utilities')
  .addItem('New case', 'formNewCase')
  .addItem('Order sheet', 'orderSheet')
  .addToUi();
}

function setCellFormat(cell, fontfamily, fontsize, fontcolor, weight, style) {
  //formato por defecto
  if (fontfamily == null){ fontfamily = "Verdana";}
  if (fontsize == null){fontsize = 9;}
  if (fontcolor == null){fontcolor = "black";}
  if (weight == null){weight = "normal";}
  if (style == null){style = "normal";}
  
  cell.setFontFamily(fontfamily);
  cell.setFontSize(fontsize);
  cell.setFontColor(fontcolor);
  cell.setFontWeight(weight);
  cell.setFontStyle(style);
 }

//Mueve los casos entre pestañas
//dest = nombre de la hoja destino
function moveCase(caso,source,dest){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var target_sheet = dest;
  var msg = "";
  var currentrow = findCase(caso,source);
  if (currentrow != 0){ //si existe el caso en la hoja actual
    if (target_sheet.getName() == sheets[2]) { //movemos a Closed sheet
      var source_range1 = source.getRange(currentrow, col_pripsi, 1, col_date); //desde la columna PRI a la Date of First contact
      var source_range2 = source.getRange(currentrow, col_status, 1, col_change); //desde la columna Status a la Change
      var tat = source.getRange(currentrow, col_tat); 
      target_sheet.insertRowsAfter(1, 1);
      var target_range1 = target_sheet.getRange(2, 1, 1, col_pripsi-1);
      var target_range2 = target_sheet.getRange(2, col_status-3, 1, col_change-1);
      var target_tat = target_sheet.getRange(2,col_tat-2);
      source_range1.copyTo(target_range1,{contentsOnly: false});
      source_range2.copyTo(target_range2,{contentsOnly: false});
      tat.copyTo(target_tat,{contentsOnly: true});
      target_sheet.getRange(2,16).setFormula("=O2+E2"); //calculamos el closed date
      
      //Quitar formato a casos aceptados
      target_sheet.getRange("A2:P").clearFormat().setVerticalAlignment("top").setWrap(true);
      target_sheet.getRange("A1:N").setHorizontalAlignment("center");
      target_sheet.getRange("N1:P").setHorizontalAlignment("left");
      
    } else {
      
      var source_range = source.getRange(currentrow,1,1,col_change);
      var nextAvailableRow = target_sheet.getLastRow()+1;
      target_sheet.insertRowsAfter(nextAvailableRow-1, 1);
      var target_range = target_sheet.getRange(nextAvailableRow,1,1,col_change);
      source_range.copyTo(target_range,{contentsOnly: false});
      
      //Ordenamos la hoja después de mover la fila
      orderSheet(target_sheet.getName(),caso,0);
    }
    
    msg = "Case "+caso+" moved to the "+target_sheet.getName()+" sheet";
    Logger.log(msg);
    
    //Borrar fila de hoja origen
    source.deleteRows(currentrow);
    ss.toast('INFO', msg, 4); //mostramos mensaje después de mover
    
    /*if (cambiar_hoja == 1) {
    setFocus(target_sheet.getName(),caso,11);
    }*/
  }
}


//Damos formato a los campos de prioridades (PRI, PSI, priority)
function formatPriorities(pri_psi,priority,row){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  if (pri_psi == "Y"){
    setCellFormat(s.getRange(row, col_pripsi),null,null,"red","bold",null);
    setCellFormat(s.getRange(row, col_case),null,null,"red","bold",null); 
    setCellFormat(s.getRange(row, col_prior),null,null,"red","bold",null);      
    if (priority != "100"){
      s.getRange(row, col_prior).setValue("100"); //marcar como prioritario
    } 
    
    //no PRI/PSI
  }else{  
    setCellFormat(s.getRange(row, col_pripsi),null,null,null,"normal",null);
    /*Marcar prioritarios*/
    if (priority > 0){
      setCellFormat(s.getRange(row, col_prior),null,null,null,null,"italic"); //setCellFormat(s.getRange(row, col_case),null,null,null,"bold",null);
    }else if (priority == 0){
      setCellFormat(s.getRange(row, col_prior),null,null,null,"normal",null); setCellFormat(s.getRange(row, col_case),null,null,null,"normal",null);
    } else {
      if (pri_psi != "Y" && (priority != 1 && priority != 2 && priority != 3 && priority != 4 && priority != 5 && priority != 100)){
        s.getRange(row, col_pripsi).setValue("N");
        setCellFormat(s.getRange(row, col_case),null,null,null,"normal",null);
      }
    }
  }
}

function setFocus(sheet,caso, origin){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tSheet = ss.getSheetByName(sheet);

  //volvemos a poner el foco donde estábamos después de ordenar  
  var columnValues = tSheet.getRange(2, col_case, tSheet.getLastRow()).getValues();
  var new_row;
  for(var i = 0 ;  i < columnValues.length ; i ++ ) {
    if(columnValues[i][0].toUpperCase() == caso.toUpperCase()) {
      new_row = i + 2;
      Logger.log("Poniendo el foco para el caso "+caso+" ["+sheet+"]-Col:"+col_case+" Row:"+new_row);
      break;
    }
  }
  
  if (origin==0)  var new_col = tSheet.getActiveCell().getColumn(); //si queremos la columna que estaba marcada antes de ordenar
  else var new_col = col_case;
  var new_range = tSheet.getRange(new_row, new_col);
  tSheet.setActiveSelection(new_range);
}

//order sheet by priority and TAT
//origin = {0: sheet ; 1: new case form}
function orderSheet(sheet, oldcase, origin){
  if (sheet == null) sheet=SpreadsheetApp.getActiveSheet().getName(); //si no recibimos la hoja por parámetro, cogemos el nombre actual
  if (origin == null) origin=0;
  if (sheet == sheets[0]  || sheet == sheets[1] ){ //solo ordernar las hojas Pending cases y Pending closure
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var target_sheet = ss.getSheetByName(sheet);
    var s = ss.getActiveSheet();
    var lastCol = target_sheet.getLastColumn();
    var lastRow = target_sheet.getLastRow();
    var r = target_sheet.getRange(2, 1, lastRow-1, lastCol);
    r.setVerticalAlignment("middle").setWrap(true);
    r.sort([{ column: col_prior, ascending: false },{ column: col_tat, ascending: false }, { column: col_case, ascending: true }]);
    target_sheet.getRange("P2:P").setFontSize(9);
    Logger.log("Ordered sheet "+sheet);
    //setFocus(sheet, oldcase, origin);
    if (sheet == sheets[0]){ //congelamos las filas de arriba solo en el pending cases
      freezesRows(s,4,5);
    }
  } else {
    ss.toast('ERROR', "Function not valid for current sheet", 5);
  }
  
}

//Busca duplicados en las hojas cuando se introduce un caso nuevo. 
//Si lo encuentra, sale aviso y borra el caso introducido.
//row=X buscamos desde el formulario
function checkDuplicate(id, sheet, row){
  Logger.log("Buscando duplicados para caso "+id);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  if (s.getName() == sheets[0] || s.getName() == sheets[1] ) {
    for (var i=0; i < sheets.length; i++){
      Logger.log("Buscando en hoja "+sheets[i]);
      var tSheet = ss.getSheetByName(sheets[i]);
      if (sheets[i] == "Closed") {
        var column = col_case - 1; //en Closed hay una columna menos
      } else {
        var column = col_case; //case column 
      }
      var columnValues = tSheet.getRange(2, column, tSheet.getLastRow()).getValues(); //buscar solo en la columna caso
      for(var j = 0 ;  j < columnValues.length ; j ++ ) {
        if(columnValues[j][0].toUpperCase() == id.toUpperCase()) {
          var new_row = j + 2;
          var msg = "Case already exists in '"+sheets[i]+"' row "+new_row
          Logger.log(msg);
          ss.toast('ERROR',msg,4);
          if (row == "X") {
            return 2; //si se ha encontrado un duplicado al introducir caso desde el formulario
          } else {
            if (tSheet != sheet && row != new_row){
              return 1;
            }
          }
        }
      }
    }
    return 0;
  }
}

function findCase(id, sheet){
  Logger.log("Buscando fila real para caso "+id);
  var hoja = sheet.getName();
  Logger.log(sheet + " " + hoja);
  var columnValues = "";
  if (hoja == sheets[2]) {
    columnValues = sheet.getRange(2, 4, sheet.getLastRow()).getValues(); //buscar solo en la columna caso
  } else {
    columnValues = sheet.getRange(2, col_case, sheet.getLastRow()).getValues(); //buscar solo en la columna caso
  }
  
  for(var j = 0 ;  j < columnValues.length ; j ++ ) {
    if(columnValues[j][0].toUpperCase() == id.toUpperCase()) {
      var new_row = j + 2;
      
      Logger.log("Encontrado ["+id+"] en fila "+new_row+" de hoja "+hoja)
      return new_row; //retornamos fila del caso
    }
  }
  return 0; //no se ha encontrado el caso
}

function addLink(celda){
  var formula = celda.getFormula();
  if (formula[0] != "=") {
    var caso = celda.getValue();
    var url = "https://rexis.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=a30&str=" + caso;
    celda.setFormula('=HYPERLINK("' + url + '","' + caso + '")');
    celda.setFontFamily("Verdana");
    celda.setFontSize(9);
    celda.setHorizontalAlignment("center");
  }
}

function addCommentReminder(celda) {
  var valor = celda.getValue();
  var comments = celda.getComment();
  var msg = "";
  
  if (valor == "1R" || valor == "2R" || valor == "3R"){
    if (valor == "1R") {
      msg = "First reminder: "; 
    } else if (valor == "2R") {
      msg = "\nSecond reminder: ";
    } else if (valor == "3R") {
      msg = "\nThird reminder: ";
    }
  
    var ndate = Utilities.formatDate(new Date(), "GMT", "dd/MM/YYYY");
    comments = comments + msg + ndate; 
    celda.setComment(comments);
  } else if (valor == "-"){
    celda.clearNote();
  }
}

