function formRanger_preconfig() {
  setformRangerSid();
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  ScriptProperties.setProperty('ssId', ssId);
  // if you are interested in sharing your complete workflow system for others to copy (with script settings)
  // Select the "Generate preconfig()" option in the menu and
  //#######Paste preconfiguration code below before sharing your system for copy#######
  
  
  
  
  
  

  //#######End preconfiguration code#######
  ScriptProperties.setProperty('preconfigStatus', 'true'); 
  onOpen();
}


function formRanger_extractorWindow () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var propertyString = '';
  for (var key in properties) {
    if ((properties[key]!='')&&(key!="preconfigStatus")&&(key!="formRanger_sid")&&(key!="ssId")) {
      var keyProperty = properties[key].replace(/[/\\*]/g, "\\\\");                                     
      propertyString += "   ScriptProperties.setProperty('" + key + "','" + keyProperty + "');\n";
    }
  }
  var app = UiApp.createApplication().setHeight(500).setTitle("Export preconfig() settings");
  var panel = app.createVerticalPanel().setWidth("100%").setHeight("100%");
  var labelText = "Copying a Google Spreadsheet copies scripts along with it, but without any of the script settings saved.  This normally makes it hard to share full, script-enabled Spreadsheet systems. ";
  labelText += " You can solve this problem by pasting the code below into a script file called \"formRanger_preconfig\" (go to formRanger in the Script Editor and select \"preconfig.gs\" in the left sidebar) prior to publishing your Spreadsheet for others to copy. \n";
  labelText += " After a user copies your spreadsheet, they will select \"Run initial configuration.\"  This will preconfigure all needed script settings.  If you got this workflow from someone as a copy of a spreadsheet, this has probably already been done for you.";
  var label = app.createLabel(labelText);
  var window = app.createTextArea();
  var codeString = "//This section sets all script properties associated with this formRanger profile \n";
  codeString += "var preconfigStatus = ScriptProperties.getProperty('preconfigStatus');\n";
  codeString += "if (preconfigStatus!='true') {\n";
  codeString += propertyString; 
  codeString += "};\n";
  codeString += "ScriptProperties.setProperty('preconfigStatus','true');\n";
  codeString += "var triggerTypesVar = ScriptProperties.getProperty('triggerTypes');\n";
  codeString += "if (triggerTypesVar) {\n";
  codeString += "  formRanger_checkSetTriggers(triggerTypesVar);\n";
  codeString += "}\n"
  codeString += "ss.toast('Custom formRanger preconfiguration ran successfully. Please check formRanger menu options to confirm system settings.');\n";
  window.setText(codeString).setWidth("100%").setHeight("350px");
  app.add(label);
  panel.add(window);
  app.add(panel);
  ss.show(app);
  return app;
}
