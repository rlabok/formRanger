function formRanger_whatIs() {
  var imageId = FORMRANGERIMAGEID;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var url = 'https://drive.google.com/uc?export=download&id='+imageId;
  var app = UiApp.createApplication().setHeight(400);
  var grid = app.createGrid(1, 2);
  var image = app.createImage(url).setWidth("120px").setHeight("120px");
  var label = app.createLabel("formRanger: Define the options in list, checkbox, and multiple choice Google Form questions by referencing any column in your Spreadsheet").setStyleAttribute('verticalAlign', 'top').setStyleAttribute('fontSize', '16px');
  grid.setWidget(0, 0, image).setWidget(0, 1, label).setStyleAttribute(0, 1, 'verticalAlign', 'top');
  var html = "Features:";
  html += "<ul><li>Prepopulate any list, checkbox, or multiple choice Google Form question with values from any column in any sheet in the attached spreadsheet.</li>";
  html += "<li>Easily set the trigger of form option refresh to occur after every form submission (allows form options to be informed by all prior form submissions), after every spreadsheet edit, or every 5 minutes.</li>";
  var description = app.createHTML(html);
  var sponsorLabel = app.createLabel("Brought to you by");
  var sponsorImage = app.createImage("http://www.youpd.org/sites/default/files/acquia_commons_logo36.png");
  var supportLink = app.createAnchor('Watch the tutorial!', 'http://www.youpd.org/formranger');
  var bottomGrid = app.createGrid(3, 1);
  bottomGrid.setWidget(0, 0, sponsorLabel);
  bottomGrid.setWidget(1, 0, sponsorImage);
  bottomGrid.setWidget(2, 0, supportLink);
  app.add(grid);
  app.add(description);
  app.add(bottomGrid);
  ss.show(app)
  return app;
}
