// This code was borrowed and modified from the Flubaroo Script author Dave Abouav
// It anonymously tracks script usage to Google Analytics, allowing our non-profit to report our impact to funders
// For original source see http://www.edcode.org


function formRanger_logRangeReferenceSet()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Range%20Reference%20Set", scriptName, scriptTrackingId, systemName)
}


function formRanger_logFormUpdated()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Form%20Updated", scriptName, scriptTrackingId, systemName)
}


function formRanger_logRepeatInstall()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Repeat%20Install", scriptName, scriptTrackingId, systemName)
}

function formRanger_logFirstInstall()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("First%20Install", scriptName, scriptTrackingId, systemName)
}


function setformRangerSid()
{ 
  var formRanger_sid = ScriptProperties.getProperty("formRanger_sid");
  if (formRanger_sid == null || formRanger_sid == "")
  {
    // user has never installed formRanger before (in any spreadsheet)
    var dt = new Date();
    var ms = dt.getTime();
    var ms_str = ms.toString();
    ScriptProperties.setProperty("formRanger_sid", ms_str);
    var formRanger_uid = UserProperties.getProperty("formRanger_uid");
    if (formRanger_uid != null || formRanger_uid != "") {
      formRanger_logRepeatInstall();
    }else {
      formRanger_logFirstInstall();
    }      
  }
}
