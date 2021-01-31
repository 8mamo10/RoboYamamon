/////////////////////////////////////////////////////////////////
// Notify today's person with note from previous person
/////////////////////////////////////////////////////////////////
function NotifyPersonToday() {
  var today = new Date();
  var formattedToday = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd');

  // Get today's person by date
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list = ss.getRange('Schedule!A1:A1000').getValues();
  var index = 0;
  for (var i = 1; i < list.length; i++) {
    if (list[i] == "") break;
    var date = new Date(list[i]);
    var formattedDate = Utilities.formatDate(date, 'JST', 'yyyy/MM/dd');
    if (formattedToday == formattedDate) {
      index = i + 1;
      break;
    }
  }
  if (index == 0) {
    Logger.log("Person not found.");
    return;
  }

  // This is today's person
  var name = ss.getRange('Schedule!C' + index).getValue();
  var userId = ss.getRange('Schedule!D' + index).getValue();

  // This is previous person and the note from him/her
  var preIndex = index - 1;
  var preName = ss.getRange('Schedule!C' + preIndex).getValue();
  var note = "";
  if (ss.getRange('Schedule!E' + preIndex).isBlank()) {
    note = "Nothing."
  } else {
    note = ss.getRange('Schedule!E' + preIndex).getValue();
  }

  var text = "Today's Person is " + name + "(<@" + userId + ">)\n\n";
  text += "Note from previous person(";
  text += preName;
  text += ")\n";
  text += "```" + note + "```";

  // Send to slack
  var jsonData =
  {
     "username"  : "Reminder",
     "icon_emoji": ":shield:",
     "text"      : text
  };
  var payload = JSON.stringify(jsonData);

  var options = {
    "method"      : "post",
    "contentType" : "application/json",
    "payload"     : payload
  };
  var hookUrl = "[Webhook URL of slack Incoming WebHooks]";
  UrlFetchApp.fetch(hookUrl, options);
}


/////////////////////////////////////////////////////////////////
// Notify tomorrrow's person with note from previous person
/////////////////////////////////////////////////////////////////
function NotifyPersonTomorrow() {
  var today = new Date();
  var formattedToday = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd');

  // Get tomorrow's person by date
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list = ss.getRange('Schedule!A1:A1000').getValues();
  var index = 0;
  for (var i = 1; i < list.length; i++) {
    if (list[i] == "") break;
    var date = new Date(list[i]);
    var formattedDate = Utilities.formatDate(date, 'JST', 'yyyy/MM/dd');
    if (formattedToday == formattedDate) {
      index = i + 1;
      break;
    }
  }
  if (index == 0) {
    Logger.log("Person not found.");
    return;
  }

  // This is tomorrow's person
  var nextIndex = index + 1;
  var name = ss.getRange('Schedule!C' + nextIndex).getValue();
  var userId = ss.getRange('Schedule!D' + nextIndex).getValue();

  // This is previous person and the note from him/her
  var preIndex = index;
  var preName = ss.getRange('Schedule!C' + preIndex).getValue();
  var note = "";
  if (ss.getRange('Schedule!E' + preIndex).isBlank()) {
    note = "Nothing."
  } else {
    note = ss.getRange('Schedule!E' + preIndex).getValue();
  }

  var text = "Tomorrow's Person is " + name + "(<@" + userId + ">)\n\n";
  text += "Note from previous person(";
  text += preName;
  text += ")\n";
  text += "```" + note + "```";

  // Send to slack
  var jsonData =
  {
     "username"  : "Reminder",
     "icon_emoji": ":shield:",
     "text"      : text
  };
  var payload = JSON.stringify(jsonData);

  var options = {
    "method"      : "post",
    "contentType" : "application/json",
    "payload"     : payload
  };
  var hookUrl = "[Webhook URL of slack Incoming WebHooks]";
  UrlFetchApp.fetch(hookUrl, options);
}


/////////////////////////////////////////////////////////////////
// Notify person list for the next 7 days with note url
/////////////////////////////////////////////////////////////////
function NotifyPersonWeekly() {
  var today = new Date();
  var formattedToday = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd');

  // Get today's person by date
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list = ss.getRange('Schedule!A1:A1000').getValues();
  var index = 0;
  for (var i = 1; i < list.length; i++) {
    if (list[i] == "") break;
    var date = new Date(list[i]);
    var formattedDate = Utilities.formatDate(date, 'JST', 'yyyy/MM/dd');
    if (formattedToday == formattedDate) {
      index = i + 1;
      break;
    }
  }
  if (index == 0) {
    Logger.log("Person not found.");
    return;
  }

  // This is tomorrow's person
  var nextIndex = index + 1;
  var userId = ss.getRange('Schedule!D' + nextIndex).getValue();

  // Load person list for seven days starting tomorrow
  var personList = "";
  for (var i = 1; i <= 7; i++) {
    var targetIndex = index + i;
    var date = ss.getRange('Schedule!A' + targetIndex).getValue();
    var formattedDate = Utilities.formatDate(date, 'JST', 'yyyy/MM/dd(E) ');
    var name = ss.getRange('Schedule!C' + targetIndex).getValue();
    personList += formattedDate;
    personList += name;
    personList += "\n";
  }

  // Mention to tomorrow's person
  var text = "<@" + userId + "> Remind: Person schedule\n";
  text += "```" + personList + "```\n";
  text += "[Note URL]"

  // Send to slack
  var jsonData =
  {
     "username"  : "Reminder",
     "icon_emoji": ":shield:",
     "text"      : text
  };
  var payload = JSON.stringify(jsonData);

  var options = {
    "method"      : "post",
    "contentType" : "application/json",
    "payload"     : payload
  };
  var hookUrl = "[Webhook URL of slack Incoming WebHooks]";
  UrlFetchApp.fetch(hookUrl, options);
}
