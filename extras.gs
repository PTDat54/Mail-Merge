function reportQuota(){
var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
//Logger.log("Remaining email quota: " + emailQuotaRemaining);
SpreadsheetApp.getActiveSpreadsheet().toast(emailQuotaRemaining,'Email Quota Remaining');

}
