var colIdentity = "F";
var mailCountLimit = 100;

function parseMailToLog() {
  var query = "label:Apple"
  var myThreads = GmailApp.search(query, 0, mailCountLimit);
  var myMsgs = GmailApp.getMessagesForThreads(myThreads);

  for (var i = 0 ; i < myMsgs.length ; i++) {
    var myMsg = myMsgs[i][0];
    var mailBody = myMsg.getPlainBody();

    var version = mailBody.match(/App\sVersion\sNumber:\s(.*)/);
    if (version === null) continue;
    if (!(version.length > 0)) continue;

    var isWaitingForReview        = mailBody.match(/The status (for the following|of your) app has changed to Waiting For Review/);
    var isInReview                = mailBody.match(/The status (for the following|of your) app has changed to In Review/);
    var isPendingDeveloperRelease = mailBody.match(/The status (for the following|of your) app has changed to Pending Developer Release/);
    var isProcessing              = mailBody.match(/The status (for the following|of your) app has changed to Processing for App Store/);
    var isForSale                 = mailBody.match(/The following app has been approved/);

    var state = ""
    if (isWaitingForReview) {
      state = "審査提出"
    } else if (isInReview) {
      state = "レビュー開始"
    } else if (isPendingDeveloperRelease) {
      state = "レビュー通過"
    } else if (isProcessing) {
      state = "公開準備開始"
    } else if (isProcessing) {
      state = "公開完了"
    } else {
      continue;
    }
    var objSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("審査記録");
    var objDate = myMsg.getDate()
    var fullDate = Utilities.formatDate(objDate, "JST", "yyyy/MM/dd (E) HH:mm:ss Z")
    if (checkExsitingLog(objSheet, fullDate)) continue;

    var date = Utilities.formatDate(objDate, "JST", "yyyy/MM/dd");
    var time = Utilities.formatDate(objDate, "JST", "HH:mm:ss");
    var dayOfWeek = ["日", "月", "火", "水", "木", "金", "土"][objDate.getDay()]

    objSheet.insertRowAfter(1);
    objSheet.getRange("A2").setValue(date);
    objSheet.getRange("B2").setValue(dayOfWeek);
    objSheet.getRange("C2").setValue(time);
    objSheet.getRange("D2").setValue(version[1]);
    objSheet.getRange("E2").setValue(state);
    objSheet.getRange(colIdentity + "2").setValue(fullDate);
  }
}

function checkExsitingLog(sheet, d) {
  var firstRow = 2;
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(firstRow, colIdentity.charCodeAt(0) - 65 + 1, lastRow - firstRow + 1).getValues();
  for (var rowIndex = 0; rowIndex < range.length; rowIndex++) {
    var row = range[rowIndex];
    // NOTE: なぜかincludesがreference errorになったのでfor文で回している
    for (var colIndex = 0; colIndex < row.length; colIndex ++) {
      if (row[colIndex] == d) return true;
    }
  }
  return false;
}