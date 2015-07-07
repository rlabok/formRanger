// This code was borrowed and modified from the Flubaroo Script author Dave Abouav
// It anonymously tracks script usage to Google Analytics, allowing our non-profit to report our impact to funders
// For original source see http://www.edcode.org


function formRanger_logRangeReferenceSet()
{
  var scriptProperties = PropertiesService.getScriptProperties();
  var systemName = scriptProperties.getProperty("systemName")
  NVSL.log("Range%20Reference%20Set", scriptName, scriptTrackingId, systemName)
}


function formRanger_logFormUpdated()
{
  var scriptProperties = PropertiesService.getScriptProperties();
  var systemName = scriptProperties.getProperty("systemName")
  NVSL.log("Form%20Updated", scriptName, scriptTrackingId, systemName)
}


function logRepeatInstall()
{
  var scriptProperties = PropertiesService.getScriptProperties();
  var systemName = scriptProperties.getProperty("systemName")
  NVSL.log("Repeat%20Install", scriptName, scriptTrackingId, systemName)
}

function logFirstInstall()
{
  var scriptProperties = PropertiesService.getScriptProperties();
  var systemName = scriptProperties.getProperty("systemName")
  NVSL.log("First%20Install", scriptName, scriptTrackingId, systemName)
}


function setSid()
{ 
  var scriptNameLower = scriptName.toLowerCase();
  var scriptProperties = PropertiesService.getScriptProperties();
  var sid = scriptProperties.getProperty(scriptNameLower + "_sid");
  if (sid == null || sid == "")
  {
    var dt = new Date();
    var ms = dt.getTime();
    var ms_str = ms.toString();
    scriptProperties.setProperty("formranger_sid", ms_str);
    var userProperties = PropertiesService.getUserProperties();
    var uid = userProperties.getProperty(scriptNameLower + "_uid");
    if (uid) {
      logRepeatInstall();
    } else {
      logFirstInstall();
      userProperties.setProperty(scriptNameLower + "_uid", ms_str);
    }      
  }
}
