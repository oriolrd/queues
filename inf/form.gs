//abre el formulario para caso nuevo
function formNewCase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet(), sheet = sh.getSheetName(); 
  var config_sheet = ss.getSheetByName("Config");
  if (sheet == sheets[0] || sheet == sheets[1]){ //mostrar formulario solo si estamos en la hoja pending cases o pending closure
    //Create the user interface
    var app = UiApp.createApplication().setTitle('Add new case');
    var ispri = app.createCheckBox("PRI/PSI").setName('ispri').setStyleAttribute('font-weight',"bold"); //PRI
    var postit = app.createCheckBox("PostIt").setName('postit').setStyleAttribute('font-weight',"bold");  //PostIt Y/N
    var label_urg = app.createLabel('Urg:').setStyleAttribute('font-weight',"bold"); //priority
    var urgency = app.createListBox().setWidth('35px').setName('urgency').setStyleAttribute('font-size',14);
    lista = getListValues("Urgency");
    urgency.addItem(0); //añadimos prioridad 0 (default)
    for(var i=0; i<lista.length; i++){
      urgency.addItem(lista[i])
    }
    
    var separador = app.createLabel('|')
    
    //case
    var label_case = app.createLabel('*Case ID:').setStyleAttribute('font-weight',"bold");
    var caseid = app.createTextBox().setWidth('120px').setName('caseid').setValue('CN-');
    //date
    var label_date = app.createLabel('*Date of first contact:').setStyleAttribute('font-weight',"bold");
    var created = app.createDateBox().setWidth('120px').setName('created');
    
    //country
    var label_country = app.createLabel('Country:').setStyleAttribute('font-weight',"bold");
    var country = app.createListBox().setWidth('200px').setName('country').setStyleAttribute('font-size',14);
    
    lista = getListValues("Country");
    for(var i=0; i<lista.length; i++){
      country.addItem(lista[i])
    }
    
    //customer
    var label_customer = app.createLabel('Customer:').setStyleAttribute('font-weight',"bold");
    var customer = app.createTextBox().setWidth('200px').setName('customer');
    
    //status
    var label_status = app.createLabel('Status:').setStyleAttribute('font-weight',"bold");
    var status = app.createListBox().setWidth('210px').setName('status').setStyleAttribute('font-size',14);;
    lista = getListValues("Status");
    for(var i=0; i<lista.length; i++){
      status.addItem(lista[i])
    }
    
    //software
    var label_sw = app.createLabel('Software:').setStyleAttribute('font-weight',"bold");
    var software = app.createListBox().setWidth('120px').setName('software').setStyleAttribute('font-size',14);
    lista = getListValues("Software");
    for(var i=0; i<lista.length; i++){
      software.addItem(lista[i])
    }
    
    //version
    var label_version = app.createLabel('Version:').setStyleAttribute('font-weight',"bold");
    var version = app.createTextBox().setWidth('120px').setName('version');
    
    //module
    var label_module = app.createLabel('Module:').setStyleAttribute('font-weight',"bold");
    var module = app.createListBox().setWidth('120px').setName('module').setStyleAttribute('font-size',14);
    lista = getListValues("Module");
    for(var i=0; i<lista.length; i++){
      module.addItem(lista[i])
    }
    
    //Problem description
    var label_problem = app.createLabel('Problem:').setStyleAttribute('font-weight',"bold");
    var problem  = app.createTextArea().setWidth('405px').setHeight('70px').setName('problem');
    
    //info
    var info = app.createLabel('Case number and creation date are mandatory').setStyleAttribute("color", "#F00").setStyleAttribute('font-style',"italic").setStyleAttribute('text-align',"right").setId('info').setVisible(false);
    
    var okHandler = app.createServerHandler('respondToOK');
    var btnOK = app.createButton('CONFIRM', okHandler);
    
    var absPanel = app.createAbsolutePanel().setWidth('600').setHeight('400');
    absPanel.add(ispri, 0, 10);  absPanel.add(postit, 75, 10);    absPanel.add(separador,135,13);absPanel.add(label_urg, 145, 13);  absPanel.add(urgency, 175, 10);absPanel.add(info, 225, 10);
    absPanel.add(label_case, 0, 45);absPanel.add(caseid, 75, 45);
    absPanel.add(label_date, 220, 45);absPanel.add(created, 360, 45);
    absPanel.add(label_sw, 0, 80);absPanel.add(software, 75, 80);
    absPanel.add(label_country,220 , 80);absPanel.add(country, 280, 80);
    absPanel.add(label_version,0 , 115);absPanel.add(version, 75, 115);
    absPanel.add(label_customer,210 , 115);absPanel.add(customer,280 , 115);
    absPanel.add(label_status,0, 150);absPanel.add(status, 75, 150);
    absPanel.add(label_module,300, 150);absPanel.add(module, 360, 150);
    absPanel.add(label_problem,0 , 185);absPanel.add(problem, 75, 185);
    absPanel.add(btnOK, 215, 270);
    
    okHandler.addCallbackElement(absPanel);
    
    app.add(absPanel);
    ss.show(app);
  }
}

//si se confirma en la ventana, añadimos los datos a la fila
function respondToOK(e) {
 
  var pri_psi = e.parameter.ispri;
  var postit = e.parameter.postit;
  //var priority = e.parameter.priority;
  var priority = e.parameter.urgency;
  var caso = e.parameter.caseid;
  var created = e.parameter.created;
  var country = e.parameter.country;
  var customer = e.parameter.customer;
  var software = e.parameter.software;
  var version = e.parameter.version;
  var status = e.parameter.status;
  var module = e.parameter.module;
  var problem = e.parameter.problem;
  var comment = "-"; //de momento no lo capturamos
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var app = UiApp.getActiveApplication();

  var template = "Complaint";
  //Destino Pending Cases
  if (status == "Evaluation in Progress" || status == "Investigation In Progress" || status == "Submitted for Peer Review" || status == "Case Closure and Summary" || status == "Pending answer"){
          var s = ss.getSheetByName(sheets[0]); ;
    
  //destino Pending Closure
  } else if (status == "Peer Review Approved" || status == "Submitted for Affiliate Review") {
          var s = ss.getSheetByName(sheets[1]);   
  }
  
  //si el caso no se ha indicado
  if ((caso == "CN-" || caso=="") || created==""){ //mostramos mensaje de info
    if ((caso == "CN-" || caso=="") && created==""){
      app.getElementById("info").setVisible(true);
    } else if ((caso == "CN-" || caso=="") && created!=""){
      app.getElementById("info").setText("Please enter case number").setVisible(true);
    } else {
      app.getElementById("info").setText("Please enter date of first contact").setVisible(true);
    }
    return app;
  
  //si id del caso se ha indicado
  } else {
    
    var exists = checkDuplicate(caso, s, "X");
    if (exists == 2){
      app.getElementById("info").setText("Case already exists").setVisible(true);
      return app;
    } else {
  
      if (pri_psi == 'true'){ pri_psi = 'Y';template="PCC";priority="100";} else {pri_psi='N'}
      if (postit == 'true'){ postit = '☑';} else {postit='✘'}
      //if (priority == 'true'){ priority = 'Y'} else {priority='N'}
      if (customer=="") customer = "-";
      if (version=="") version = "-";
            
      var lRow = s.getLastRow();
      var lCol = s.getLastColumn(), range = s.getRange(lRow,1,1,lCol);
      
      s.insertRowsAfter(lRow, 1);
      range.copyTo(s.getRange(lRow+1, 1, 1, lCol), {contentsOnly:false});
      var nextRow=lRow+1;
      
      s.getRange(nextRow, col_postit).setValue(postit);
      s.getRange(nextRow, col_pripsi).setValue(pri_psi);
      s.getRange(nextRow, col_prior).setValue(priority);
      s.getRange(nextRow, col_date).setValue(created);
      s.getRange(nextRow, col_reminder).setValue("-");
      s.getRange(nextRow, col_status).setValue(status);
      s.getRange(nextRow, col_case).setValue(caso);
      s.getRange(nextRow, col_country).setValue(country);
      s.getRange(nextRow, col_customer).setValue(customer);
      s.getRange(nextRow, col_user).clearContent();
      s.getRange(nextRow, col_description).setValue(problem);
      s.getRange(nextRow, col_comment).setValue(comment);
      s.getRange(nextRow, col_module).setValue(module);
      s.getRange(nextRow, col_sw).setValue(software);
      s.getRange(nextRow, col_version).setValue(version);
      s.getRange(nextRow, col_change).clearContent();
      
      var now = new Date(); //añadimos fecha en la que hemos añadido el caso
      s.getRange(nextRow, col_change).setValue(now).setNumberFormat('yyyy-mm-dd');
      
      addLink(s.getRange(nextRow, col_case));
      s.getRange(nextRow,1,1,lCol).setVerticalAlignment("middle").setWrap(true);
      
      ss.toast('INFO', "Case "+caso+" succesfully created", 3);
      
      //damos formato a prioritarios
      formatPriorities(pri_psi,priority,nextRow);
      
      //ordenamos
      orderSheet(s.getName(),caso,1);

      //Close the user interface
      return app.close();
    }
  }
}

//Genera array con los datos en las columnas de la hoja Config
function getListValues(column) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName("Config");
  var data = s.getDataRange().getValues();
  var headers = 1; // number of header rows to skip at top
  var lista = [], item;
  
  if (column != "Status" && column != "Urgency") lista.push(""); //si no son Status o Urgency, añadimos un valor en blanco
  var col = getColumnByName(s,column);
  for (var row = headers; row < data.length; row++) {
    item = data[row][col-1];
    if (item && item != "Case Closed") lista.push(item);
  }

  return lista;
}

//Devuelve el número de columna a partir del título en la primera fila
function getColumnByName(sheet, name) {
  var range = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var values = range.getValues();
  
  for (var row in values) {
    for (var col in values[row]) {
      if (values[row][col] == name) {
        return parseInt(col)+1;
      }
    }
  }
  
  throw 'failed to get column for '+name;
}

function getRegion(country){
var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName("Config");
  var col = getColumnByName(s,"Country");
  var values = s.getRange(2, col, s.getLastRow(),2).getValues();
  for(var row in values) {
    if(values[row][0].toUpperCase() == country.toUpperCase()) {
      Logger.log("Found region for "+country+": "+values[row][1]);
      return values[row][1]; //devolvemos la region
    }
  }
  Logger.log("No region found for country "+country);
  return 0;
}


