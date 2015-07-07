// Written by Andrew Stillman for New Visions for Public Schools
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html
// Support and contact at http://www.youpd.org/formranger
// #modified v2: rlabok

var formUrl = ''
var scriptName = "formRanger"
var scriptTrackingId = "UA-40688501-1"
var FORMRANGERIMAGEID = "0B2vrNcqyzernTzhZT0JZYTVFTWc";

function onInstall() {
  onOpen();
}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scriptProperties = PropertiesService.getScriptProperties();
  var preconfigStatus = scriptProperties.getProperty('preconfigStatus');
  var menuItems = [];
  menuItems[0] = {name:'What is formRanger?', functionName: 'formRanger_whatIs'};
  menuItems[1] = null;
  if (preconfigStatus=="true") {
    menuItems[2] = {name:'Assign form item(s) to column(s)', functionName: 'formRangerUi'};
    menuItems[3] = {name:'Refresh form', functionName: 'formRanger_populateForm'};
    menuItems[4] = {name:'Set triggers for form refresh', functionName:'formRanger_triggerSettings'};
    menuItems[5] = null;
    menuItems[6] = {name:'Export settings', functionName:'formRanger_extractorWindow'};
  } else {
    menuItems[2] = {name:'Run initial configuration',functionName:'formRanger_preconfig'};
  }
  ss.addMenu('formRanger', menuItems);
}

function formRangerUi() {
  var allowedTypes = ["MULTIPLE_CHOICE","LIST","CHECKBOX"];
  var scriptProperties = PropertiesService.getScriptProperties();
  var properties = scriptProperties.getProperties();
  if (properties.questionRanges) {
    var questionRanges = Utilities.jsonParse(properties.questionRanges);
  }
  var app = UiApp.createApplication().setHeight(375);
  var imageId = FORMRANGERIMAGEID;
  var url = 'https://drive.google.com/uc?export=download&id='+imageId;
  var waitingImage = app.createImage(url).setWidth('150px').setHeight('150px').setStyleAttribute('position', 'absolute').setStyleAttribute('left', '130px').setStyleAttribute('top', '20px').setVisible(false);
  app.add(app.createLabel("Columns are identified by sheet and header name, with values starting in row 2.").setStyleAttribute('marginBottom', '5px'));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  app.setTitle("Assign form item(s) to column(s)");
  var formUrl = ss.getFormUrl();
  if (!formUrl) {
    Browser.msgBox("It appears you have no Google Form attached to this Spreadsheet.  Please create your form before running this script.");
    return;
  }
  var form = FormApp.openByUrl(formUrl);
  var items = form.getItems();
  var mainPanel = app.createVerticalPanel();
  var mainScrollPanel = app.createScrollPanel().setHeight("300px");
  var sheets = ss.getSheets();
  var sheetNames = [];
  var sheetIds = [];
  for (var i=0; i<sheets.length; i++) {
    sheetNames.push(sheets[i].getName());
    sheetIds.push(sheets[i].getSheetId());
  }
  var subPanels = [];
  var checkBoxes = [];
  for (var i = 0; i<items.length; i++) {
    mainPanel.add(app.createLabel("Q" + (i+1)).setWidth("100%").setStyleAttribute('padding', '4px').setStyleAttribute('backgroundColor', 'grey').setStyleAttribute('color', 'white'));
    subPanels[i] = app.createVerticalPanel().setHeight("100px").setStyleAttribute('marginLeft', '10px');
    var titlePanel = app.createHorizontalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setWidth("450px").setStyleAttribute('padding', '10px');
    var qTitle = app.createLabel(items[i].getTitle()).setStyleAttribute('textAlign', 'left');
    var type = items[i].getType().toString();
    var qId = items[i].getId();
    var qType = app.createLabel("Type: " + type).setStyleAttribute('textAlign', 'right');
    var checkBoxHandler = app.createServerHandler('formRanger_refreshEnabled').addCallbackElement(subPanels[i]);
    checkBoxes[i] = app.createCheckBox('Populate options from column').setName('enabled-'+qId).addClickHandler(checkBoxHandler);
    titlePanel.add(qTitle).add(qType);
    subPanels[i].add(titlePanel);
    subPanels[i].add(app.createHidden('qId', qId));
    var sheetSelect = app.createListBox().setName('sheetId-'+qId).setId('sheetIdSelect-'+qId).setEnabled(false);
    var sheetLabel = app.createLabel('Sheet').setId('sheetLabel-'+qId).setStyleAttribute('color', '#C0C0C0');
    var headerLabel = app.createLabel('Column').setId('headerLabel-'+qId).setStyleAttribute('color', '#C0C0C0');
    var sheetChangeHandler = app.createServerHandler('formRanger_refreshHeaders').addCallbackElement(subPanels[i]);
    for (var j=0; j<sheetNames.length; j++) {
      sheetSelect.addItem(sheetNames[j],sheetIds[j]);
    }
    sheetSelect.addChangeHandler(sheetChangeHandler);
    var headerSelect = app.createListBox().setName('header-'+qId).setId('headerSelect-'+qId).setEnabled(false);
    if (questionRanges) {
      if (!questionRanges['sheetId-'+qId]) {
        sheetSelect.setSelectedIndex(0);
        try {
          var headers = sheets[0].getRange(1, 1, 1, sheets[0].getLastColumn()).getValues()[0];
        } catch(err) {
          headers = ['No data in sheet'];
        }
        for (var h=0; h<headers.length; h++) {
          headerSelect.addItem(headers[h]);
        }
      } else {
        checkBoxes[i].setValue(true);
        var sheetId = questionRanges['sheetId-'+qId];
        var sheetIndex = formRanger_getSheetIndexFromId(ss, sheetId);
        sheetSelect.setEnabled(true).setSelectedIndex(sheetIndex);
        var sheet = formRanger_getSheetFromId(ss, sheetId);
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var headerIndex = headers.indexOf(questionRanges['header-'+qId]);
        for (var k=0; k<headers.length; k++) {
          headerSelect.addItem(headers[k]);
        }
        headerSelect.setEnabled(true).setSelectedIndex(headerIndex);
        sheetLabel.setStyleAttribute('color', 'black');
        headerLabel.setStyleAttribute('color', 'black');
      }
    } else { 
      sheetSelect.setSelectedIndex(0);
      try {
        var headers = sheets[0].getRange(1, 1, 1, sheets[0].getLastColumn()).getValues()[0];
      } catch(err) {
        headers = ['No data in sheet'];
      }
      for (var h=0; h<headers.length; h++) {
        headerSelect.addItem(headers[h]);
      }
    }
  if (allowedTypes.indexOf(type)!=-1) {
    var subGrid = app.createGrid(2, 3).setCellPadding(5);
    subGrid.setWidget(0, 0, checkBoxes[i]);
    subGrid.setWidget(0, 1, sheetLabel);
    subGrid.setWidget(0, 2, sheetSelect);
    subGrid.setWidget(1, 1, headerLabel);
    subGrid.setWidget(1, 2, headerSelect);
    subPanels[i].add(subGrid);
  } else {
    subPanels[i].add(app.createLabel('Range validation not available for this question type').setStyleAttribute('color', 'grey').setStyleAttribute('margin', '35px'));
    }
    mainPanel.add(subPanels[i]);
  }
  mainScrollPanel.add(mainPanel);
  var saveHandler = app.createServerHandler('formRanger_saveSettings').addCallbackElement(mainScrollPanel);
  var saveButton = app.createButton('Save settings', saveHandler).setStyleAttribute('marginTop', '15px');
  var waitingHandler = app.createClientHandler().forTargets(mainPanel).setStyleAttribute('opacity', '0.5').forTargets(waitingImage).setVisible(true);
  saveButton.addClickHandler(waitingHandler);
  app.add(mainScrollPanel);
  app.add(saveButton);
  app.add(waitingImage);
  ss.show(app);
  return app;
}


function formRanger_saveSettings(e) {
  setSid();
  var app = UiApp.getActiveApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var form = FormApp.openByUrl(ss.getFormUrl());
  var items = form.getItems();
  var questionRanges = new Object();
  var scriptProperties = PropertiesService.getScriptProperties();
  for (var i=0; i<items.length; i++) {
    var qId = items[i].getId();
    var thisSheetId = e.parameter['sheetId-'+qId];
    var thisHeader = e.parameter['header-'+qId];
    var enabled = e.parameter['enabled-'+qId];
    if ((thisSheetId)&&(thisHeader!="No data in this sheet")&&(enabled=="true")) {
      questionRanges['sheetId-'+qId] = thisSheetId;
      questionRanges['header-'+qId] = thisHeader;
    }
  }
  try {
    formRanger_logRangeReferenceSet();
  } catch(err) {
  }
  questionRanges = Utilities.jsonStringify(questionRanges);
  
  scriptProperties.setProperty('questionRanges', questionRanges);
  scriptProperties.setProperty('ssId', ssId);
  formRanger_populateForm();
  app.close();
  return app;
}

function formRanger_refreshEnabled(e) {
  var app = UiApp.getActiveApplication();
  var qId = e.parameter.qId;
  var enabled = e.parameter['enabled-'+qId];
  var sheetIdSelect = app.getElementById('sheetIdSelect-'+qId);
  var sheetLabel = app.getElementById('sheetLabel-'+qId);
  var headerSelect = app.getElementById('headerSelect-'+qId);
  var headerLabel = app.getElementById('headerLabel-'+qId);
  if (enabled=="true") { 
    sheetIdSelect.setEnabled(true);
    sheetLabel.setStyleAttribute('color', 'black');
    headerSelect.setEnabled(true);
    headerLabel.setStyleAttribute('color', 'black');
    
  } else {
    sheetIdSelect.setEnabled(false);
    sheetLabel.setStyleAttribute('color', '#C0C0C0');
    headerSelect.setEnabled(false);
    headerLabel.setStyleAttribute('color', '#C0C0C0');
  }
  return app;
}

function formRanger_refreshHeaders(e) {
  var app = UiApp.getActiveApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qId = e.parameter.qId;
  var sheetId = e.parameter['sheetId-'+qId];
  var headerSelect = app.getElementById('headerSelect-'+qId);
  var sheet = formRanger_getSheetFromId(ss, sheetId);
  try {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  } catch(err) {
    headers = ['No data in sheet'];
  }
  headerSelect.clear();
  for (var i=0; i<headers.length; i++) {
    headerSelect.addItem(headers[i]);
  }
  return app;
}

function formRanger_getSheetFromId(ss, sheetId) {
  var sheets = ss.getSheets();
  for (var i=0; i<sheets.length; i++) {
    if (sheetId == sheets[i].getSheetId()) {
      return sheets[i];
    }
  }
  return;
}


function formRanger_getSheetIndexFromId(ss, sheetId) {
  var sheets = ss.getSheets();
  for (var i=0; i<sheets.length; i++) {
    if (sheetId == sheets[i].getSheetId()) {
      return i;
    }
  }
  return;
}


function formRanger_populateForm() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var ssId = scriptProperties.getProperty('ssId');
  var ss = SpreadsheetApp.openById(ssId);
  var questionRanges = scriptProperties.getProperty('questionRanges');
  if (!questionRanges) {
    Browser.msgBox("Oops. Something is wrong with your settings.");
    return;
  } else {
    questionRanges = Utilities.jsonParse(questionRanges);
  }
  var form = FormApp.openByUrl(ss.getFormUrl());
  var items = form.getItems();
  for (var i=0; i<items.length; i++) {
    var qId = items[i].getId();
    if ((questionRanges['sheetId-'+qId])&&(questionRanges['header-'+qId])) {
      var sheet = formRanger_getSheetFromId(ss, questionRanges['sheetId-'+qId]);
      var values = formRanger_getColValues(sheet, questionRanges['header-'+qId]);
      if (values.length == 0) {
        values[0] = "No values found in column \"" + questionRanges['header-'+qId] + "\"";
      }
      var type = items[i].getType().toString();
      if (type == "LIST") {
        items[i].asListItem().setChoiceValues(values);
      }
      if (type == "MULTIPLE_CHOICE") {
        items[i].asMultipleChoiceItem().setChoiceValues(values);
      }
      if (type == "CHECKBOX") {
        items[i].asCheckboxItem().setChoiceValues(values);
      }
    }
  }
  formRanger_logFormUpdated();
}


function formRanger_getColValues(sheet, header) {
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var col = headers.indexOf(header) + 1;
  try {
    var values = sheet.getRange(2,col,sheet.getLastRow()-1,1).getValues();
    var valueArray = [];
    for (var i=0; i<values.length; i++) {
      if (values[i][0]!='') {
        valueArray.push(values[i][0]);
      }
    }
    return valueArray;
  } catch(err) {
    return ['no data in column:' + header];
  }
}

function formRanger_triggerSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle('Set triggers for form refresh').setHeight(180);
  var panel = app.createVerticalPanel();
  var label = app.createLabel("Decide how you want to trigger the refresh of Spreadsheet-linked form options").setStyleAttribute('marginBottom', '15px');
  panel.add(label);
  var scriptProperties = PropertiesService.getScriptProperties();
  var triggerTypes = scriptProperties.getProperty('triggerTypes');
  if (triggerTypes) {
    triggerTypes = triggerTypes.split(",");
  } else {
    triggerTypes = [];
  }
  var onFormCheckBox = app.createCheckBox('On every form submit').setName('onFormSubmit');
  var onEditCheckBox = app.createCheckBox('On every spreadsheet edit').setName('onEdit');
  var everyFiveCheckBox = app.createCheckBox('Every 5 minutes').setName('everyFive');
  if (triggerTypes.indexOf("onFormSubmit")!=-1) {
    onFormCheckBox.setValue(true);
  }
  if (triggerTypes.indexOf("onEdit")!=-1) {
    onEditCheckBox.setValue(true);
  }
  if (triggerTypes.indexOf("everyFive")!=-1) {
    everyFiveCheckBox.setValue(true);
  }
  var triggerSaveHandler = app.createServerHandler('formRanger_saveTriggers').addCallbackElement(panel);
  var button = app.createButton('Save trigger settings', triggerSaveHandler).setStyleAttribute('marginTop', '15px');
  panel.add(onFormCheckBox).add(onEditCheckBox).add(everyFiveCheckBox).add(button);
  app.add(panel);
  ss.show(app);
  return app;
}

function formRanger_saveTriggers(e) {
  var app = UiApp.getActiveApplication();
  var onFormSubmit = e.parameter.onFormSubmit;
  var onEdit = e.parameter.onEdit;
  var everyFive = e.parameter.everyFive;
  var triggerTypes = [];
  var scriptProperties = PropertiesService.getScriptProperties();
  if (onFormSubmit=="true") {
    triggerTypes.push('onFormSubmit');
  }
  if (onEdit=="true") {
    triggerTypes.push('onEdit');
  }
  if (everyFive=="true") {
    triggerTypes.push('everyFive');
  }
  formRanger_checkSetTriggers(triggerTypes);
  triggerTypes.join();
  scriptProperties.setProperty('triggerTypes', triggerTypes.toString());
  app.close();
  return app;
}



function formRanger_checkSetTriggers(triggerTypesVar) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var ssId = scriptProperties.getProperty('ssId');
  if (!ssId) {
    ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  }
  var triggers = ScriptApp.getProjectTriggers();
  var existingTriggerTypes = [];
  for (var i=0; i<triggers.length; i++) {
    if ((triggerTypesVar.indexOf("onFormSubmit")==-1)&&(triggers[i].getEventType().toString()=="ON_FORM_SUBMIT")) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    if ((triggerTypesVar.indexOf("onEdit")==-1)&&(triggers[i].getEventType().toString()=="ON_EDIT")) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    if ((triggerTypesVar.indexOf("everyFive")==-1)&&(triggers[i].getEventType().toString()=="CLOCK")) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    existingTriggerTypes.push(triggers[i].getEventType().toString());
  }
  if ((triggerTypesVar.indexOf("onFormSubmit")!=-1)&&(existingTriggerTypes.indexOf("ON_FORM_SUBMIT")==-1)) {
    ScriptApp.newTrigger('formRanger_populateForm').forSpreadsheet(ssId).onFormSubmit().create();
  } 
  if ((triggerTypesVar.indexOf("onEdit")!=-1)&&(existingTriggerTypes.indexOf("ON_EDIT")==-1)) {
    ScriptApp.newTrigger('formRanger_populateForm').forSpreadsheet(ssId).onEdit().create();
  }
  if ((triggerTypesVar.indexOf("everyFive")!=-1)&&(existingTriggerTypes.indexOf("CLOCK")==-1)) {
    ScriptApp.newTrigger('formRanger_populateForm').timeBased().everyMinutes(5).create();
  }
  return;
}
