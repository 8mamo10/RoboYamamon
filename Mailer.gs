// Mail
function doAnnounce() {
  Logger.log('アナウンスメールの配信を開始します。');
  var mailList = loadMailList();
  var mailBody = loadMailBody();
  sendMail(mailList, mailBody);
  Logger.log('アナウンスメールの配信を終了します。');
}

function loadMailList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list1 = ss.getRange('参加者リスト!C2:C1000').getValues();
  var list2 = ss.getRange('参加者リスト!E2:E1000').getValues();
  list1 = dropNullItemFromArray(list1);
  list2 = dropNullItemFromArray(list2);
  var list = list1.concat(list2);
  //Logger.log(list);
  return list;
}

function loadMailBody() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var title = ss.getRange('メール本文!B1').getValue();
  //Logger.log(title);
  var body = ss.getRange('メール本文!B2').getValue();
  //Logger.log(body);
  return {'title': title, 'body': body};
}

function sendMail(mailList, mailBody)
{
  var size = mailList.length;
  Logger.log("宛先合計数: " + size);
  var unit = 40;
  for (var i = 0; i < size; i+=unit) {
    var sendList = mailList.slice(i, i + unit);
    //Logger.log(sendList);
    // sender info
    var options = {
      from: 'xxxxxxxx@gmail.com',
      name: '[Sender]',
      bcc : array2csv(sendList)
    };
    //Logger.log(options);
    //GmailApp.sendEmail('xxxxxxxx@gmail.com', mailBody.title, mailBody.body, options);
    Logger.log('メールを送信しました。');
    Logger.log('-- 件数 --');
    Logger.log(sendList.length);
    Logger.log('-- 宛先 --');
    Logger.log(options.bcc);
    Logger.log('----------');
    Logger.log('-- Subject --');
    Logger.log(mailBody.title);
    Logger.log('-- Body --');
    Logger.log(mailBody.body);
  }
}

function dropNullItemFromArray(array)
{
  var newArray = [];
  for each (var value in array) {
    if(value != null && value != "") {
      newArray.push(value);
    }
  }
  return newArray;
}

function array2csv(array){
  var csv="";
  for(var i = 0; i < array.length; i++){
    csv += array[i]+",";
  }
  csv = csv.slice(0,-1);
  return csv;
}

// UI
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('メール送信');
  menu.addItem('メールを送信する', 'onClickItemSendMail');
  menu.addToUi();
}

function onClickItemSendMail() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('本当にメールを送信しますか？', ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    Logger.log('The user clicked "Yes."');
    doAnnounce();
  } else {
    Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
  }
}
