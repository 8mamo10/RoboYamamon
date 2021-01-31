function doAnnounce() {
  Logger.log('闇練メールの配信を開始します。');

  var today = new Date();
  //Logger.log(today);
  var formattedToday = Utilities.formatDate(today, 'VST', 'yyyy/MM/dd');
  //Logger.log(formattedToday);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list = ss.getRange('Schedule!A1:A1000').getValues();

  var index = 0;
  for (var i = 1; i < list.length; i++) {
    if (list[i] == "") break;

    var date = new Date(list[i]);
    var formattedDate = Utilities.formatDate(date, 'VST', 'yyyy/MM/dd');
    //Logger.log(formattedDate);

    if (formattedToday == formattedDate) {
      index = i + 1;
      break;
    }
  }

  // title
  var title = getMailTitle(formattedDate);
  Logger.log(title);

  var text = ss.getRange('Schedule!B' + index).getValue();
  Logger.log(text);

  // densuke
  var url = createDensuke(formattedDate);
  if (url == null) {
    errorReport("densuke の作成に失敗しました。");
    return null;
  }
  //Logger.log(url);

  // store densuke url
  ss.getRange('Schedule!C' + index).setValue(url);

  // body
  var body = getMailBody(formattedDate, text, url);
  //Logger.log(body);

  if (title == null || body == null) {
    errorReport("メールの作成に失敗しました。");
    return;
  }
  // 20200720: not send mail, but send line instead
  //sendMail(title, body);
  //Logger.log('闇練メールの配信を正常終了しました。');

  // line
  var lineBody = getMailBody(formattedDate, text, url);
  var lineText = getLineText(formattedDate, lineBody, url);
  //Logger.log(lineText);
  sendLine(lineText);
  Logger.log('LINE の配信を正常終了しました。');
}

function getMailTitle(date) {
  //return '闇練 ' + date + ' 週';
  var origin = new Date(date);
  var originYear = origin.getYear();
  var originMonth = origin.getMonth();
  var nextMonth = (originMonth + 2) % 12;
  var nextYear = 1900 + originYear;
  if (nextMonth == 1) {
    nextYear++;
  }
  return '闇練 ' + nextYear + '年' + nextMonth + '月';
}

function getMailBody(date, text, url) {
  var prefix = '闇練ジャーの皆様\n\nおつかれさまです。[Sender]です。\n来月の練習予定をご連絡します。\n\n';
  var suffix = '\n\n伝助へ予定入力お願いします。\n\n';
  var ad = '\n-----\n[Sender]';

  return prefix + text + suffix + url + ad;
}

function createDensuke(date) {

  var formData = {
    'eventname'   : getMailTitle(date),
    'schedule'    : getMonthlySchedule(date),
    'explain'     : getMailTitle(date),
    'email'       : 'xxxxxxxx+densuke@gmail.com',
    'pw'          : 0,
    'password'    : "",
    'eventchoice' : 1
  };
  var options = {
    'method'          : 'post',
    'followRedirects' : false,
    'payload'         : formData
  };

  // see request
  /*
  var response = UrlFetchApp.getRequest('https://www.densuke.biz/create', options);
  for(i in response) {
  Logger.log(i + ": " + response[i]);
  }
  return "";
  */

  var response = UrlFetchApp.fetch('https://www.densuke.biz/create', options);
  //Logger.log(response);
  var headers = response.getAllHeaders();
  //Logger.log(headers);
  var location = headers['Location'];
  //Logger.log(location);

  // dummy header
  //var location = 'complete?cd=b6Uv7nNsELrJDPWm&sd=indOnTxVr9j0U';

  //var regex = /complete\?cd\=+\&sd/g;
  var result = location.match(/cd\=(.+)\&sd/g);
  //Logger.log(result);
  if (result == null) {
    return null;
  }
  var cd = result[0].slice(0, -3); // cut &sd
  var url = 'http://densuke.biz/list?' + cd;
  //Logger.log(url);

  return url;
}

function sendMail(title, body) {
  // dummy
  // title = 'dummy title';
  // body = 'dummy body';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list = ss.getRange('ML!C1:C1000').getValues();
  var address = [];
  for (var i = 1; i < list.length; i++) {
    if (list[i] == "") break;
    //var address = list[i];
    Logger.log(list[i]);
    //GmailApp.sendEmail(address, title, body);
    address.push(list[i]);
  }

  // sender info
  var options = {
    from: 'xxxxxxxx+yamiren@gmail.com',
    name: '[Sender]'
  };

  GmailApp.sendEmail(address, title, body, options);
}

function getWeeklySchedule(date) {
  var schedule = "";
  // dummy date
  //date = '2017/07/03';

  // Yamiren entries (base is Monday)
  //var entryList = [
  //  [1, '20-22'], // Tue 20-22
  //  [5, '8-14'], // Sat 8-14
  //  [5, '20-22（全体練習）']  // Sat 20-22
  //];

  var entryList = [
    [1, '20-22時'], // Tue 20-22
    [5, '10-13時'], // Sat 10-13
  ];

  for (var i = 0; i < entryList.length; i++) {
    entry = entryList[i];
    var diff = entry[0];
    var time = entry[1];

    var base = new Date(date);
    base.setDate(base.getDate() + diff);
    var formattedDate = Utilities.formatDate(base, 'VST', 'MM/dd(EEE)');
    //Logger.log(formattedDate);
    schedule += formattedDate + ' ' + time + '\n';
  }
  //Logger.log(schedule);

  return schedule;
}

function getMonthlySchedule(date) {
  var schedule = "";
  // dummy date
  //date = '2020/12/25';
  var origin = new Date(date);
  var originYear = origin.getYear();
  //Logger.log(originYear);
  var originMonth = origin.getMonth();
  //Logger.log(originMonth);
  var nextMonth = (originMonth + 2) % 12;
  //Logger.log(nextMonth);
  var nextYear = 1900 + originYear;
  if (nextMonth == 1) {
    nextYear++;
  }
  //Logger.log(nextYear);

  for (var i = 1; i <= 31; i++) {
    var strDate = nextYear + '/' + nextMonth + '/' + i;
    //Logger.log(strDate);
    var timestamp = Date.parse(strDate);
    if (isNaN(timestamp)) continue;

    var aDate = new Date(strDate);
    var dow = aDate.getDay();
    if (dow == 2) {
      var formattedDate = Utilities.formatDate(aDate, 'VST', 'MM/dd(EEE)');
      schedule += formattedDate + ' 20-22時\n';
    } else if (dow == 6) {
      var formattedDate = Utilities.formatDate(aDate, 'VST', 'MM/dd(EEE)');
      schedule += formattedDate + ' 10-13時\n';
    }
  }
  Logger.log(schedule);
  return schedule;
}


function errorReport(msg) {
  Logger.log(msg);
  GmailApp.sendEmail('xxxxxxxx+yamiren@gmail.com', '闇練メール配信エラー', msg);
}

function getLineText(date, text, url) {
  return getMailTitle(date) + '\n' + text;
}

function sendLine(text) {
  var headers = {
    //'Authorization': 'Bearer xxxxxxxx' // sandbox
    'Authorization': 'Bearer yyyyyyyy' // official
  };
  var payload = "message=" + text;
  var options = {
    'method'  : 'post',
    'headers' : headers,
    'payload' : payload
  };
  var response = UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}