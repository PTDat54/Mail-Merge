const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Status";
const CC_COL = 'CC';
const BCC_COL = 'BCC';
const ATTACHMENT_COL1 = "Attachment ID 1";
const ATTACHMENT_COL2 = "Attachment ID 2";
const ATTACHMENT_COL3 = "Attachment ID 3";
const SUBJECT_COL = "Email Subject";
const TIME_COL = "Time";
const TRIGGER_COL = "Trigger ID";
 
// Creates the menu item "Mail Merge" for user to run scripts on drop-down.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Send Emails', 'sendEmails')
      .addItem('Send Emails with Timers','createTriggers')
      .addItem('Get Quota','reportQuota')     
      .addItem('List files in a folder','listAllFilesInFolder')     
     .addToUi();

  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  PropertiesService.getScriptProperties().setProperty("spreadsheetId", ssId);
}

/**
 * Fill template string with data object
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
*/
function getGmailTemplateFromDrafts_(subject_line){
  try {
    // get drafts
    const drafts = GmailApp.getDrafts();
    // filter the drafts that match subject line
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    // get the message object
    const msg = draft.getMessage();
    // getting attachments so they can be included in the merge
    const attachments = msg.getAttachments();
    return {message: {subject: subject_line, text: msg.getPlainBody(), html:msg.getBody()}, 
            attachments: attachments};
  } catch(e) {
    throw new Error("Can't find Gmail draft");
  }
}

/**
 * Filter draft objects with the matching subject linemessage by matching the subject line.
 * @param {string} subject_line to search for draft message
 * @return {object} GmailDraft object
*/
function subjectFilter_(subject_line){
  return function(element) {
    if (element.getMessage().getSubject() === subject_line) {
      return element;
    }
  }
}
  
/**
 * Escape cell data to make JSON safe
 * @see https://stackoverflow.com/a/9204218/1027723
 * @param {string} str to escape JSON special characters from
 * @return {string} escaped string
*/
function escapeData_(str) {
  return str
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
}

/**
 * Fill template string with data object
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
*/
function fillInTemplateFromObject_(template, data) {
  // we have two templates one for plain text and the html body
  // stringifing the object means we can do a global replace
  let template_string = JSON.stringify(template);

  // token replacement
  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    let token = key.replace(/[{}]+/g, "").trim();
    let match = token.match(/<[^>]+>(.*?)<\/[^>]+>/g); // Find all HTML tags

    if (match) {
      let dataKey = token.replace(/<[^>]+>|<\/[^>]+>/g, ""); // Remove all HTML tags
      let value = escapeData_(data[dataKey] || "");

      // Replace the data key within the original HTML tags
      return token.replace(token.replace(/<[^>]+>|<\/[^>]+>/g, ""), value);
    } else {
      return escapeData_(data[token] || "");
    }
  });
  return  JSON.parse(template_string);
}

function sendEmails(subjectLine, sheet=SpreadsheetApp.getActiveSheet()) {
  // option to skip browser prompt if you want to use this code in other projects
  if (!subjectLine){
    subjectLine = Browser.inputBox("Mail Merge", 
                                      "Type or copy/paste the subject line of the Gmail " +
                                      "draft message you would like to mail merge with:",
                                      Browser.Buttons.OK_CANCEL);
                                      
    if (subjectLine === "cancel" || subjectLine == ""){
      return;
    }
  }
  
  // get the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  
  // get the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetch displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // assuming row 1 contains our column headings
  const heads = data.shift(); 
  
  // get the index of column named 'Email Status' (Assume header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  
  // convert 2d array into object array
  // @see https://stackoverflow.com/a/22917499/1027723
  // for pretty version see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // used to record sent emails
  const out = [];

  // loop through all the rows of data
  obj.forEach(function(row, rowIdx){
    // only send emails is email_sent cell is blank and not hidden by filter
    if (row[EMAIL_SENT_COL] == ''){
      try {
        SpreadsheetApp.getActiveSpreadsheet().toast('Attempting to send row:' + (rowIdx+2),'Status');
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
        
        var filelist = []; //create and array to store files  //get the three files
        filelist[0] = row[ATTACHMENT_COL1];
        filelist[1] = row[ATTACHMENT_COL2];
        filelist[2] = row[ATTACHMENT_COL3];
        var blobs = [];
        var x = 0;
        for (var i = 0; i < filelist.length; i++) {
   
        //check if file ID is on sheet
          if (filelist[i] == '') 
            {
            // no ID found
            }
          else 
            {
            blobs[x] = DriveApp.getFileById(filelist[i]).getBlob();
            x++;
            }    
        }       
        
        // if (row[ATTACHMENT_COL1] != ''){  //old code for single file
        // var attach_file = DriveApp.getFileById(row[ATTACHMENT_COL1]);
        // }
        // @see https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
        // if you need to send emails with unicode/emoji characters change GmailApp for MailApp
        // Uncomment advanced parameters as needed (see docs for limitations)
        GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          bcc: row[BCC_COL],
          cc: row[CC_COL],
          // from: 'he@wcg.ac.uk',
          // name: 'WCG HE',
          // replyTo: 'a.reply@email.com',
          // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
          // attachments: emailTemplate.attachments
          attachments: blobs
        });
        // modify cell to record email sent date
        out.push([new Date()]);
     
      //  var currentdate = new Date();
      //  var gbdatetime = currentdate.toLocaleDateString("en-GB") + " " + currentdate.toLocaleTimeString("en-GB")
      //  out.push(gbdatetime);

      } catch(e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });
  
  // updating the sheet with new data
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);
}

function sendScheduledEmailForRow(e) {
  var triggerID = e.triggerUid;

  // Get sheet from information written in Project Properties
  var ssId = PropertiesService.getScriptProperties().getProperty("spreadsheetId");
  var ss = SpreadsheetApp.openById(ssId);
  var sheetId = PropertiesService.getScriptProperties().getProperty("sheetId");
  var sheet = ss.getSheetById(parseInt(sheetId));

  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  var heads = data.shift();

  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  var obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));  

  obj.forEach(function(row, rowIdx) {
    if (row[TRIGGER_COL] == triggerID && row[EMAIL_SENT_COL] == '') {
      var subject = row[SUBJECT_COL];
      var emailTemplate = getGmailTemplateFromDrafts_(subject);

      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        // array to store files
        var filelist = [];
        filelist[0] = row[ATTACHMENT_COL1];
        filelist[1] = row[ATTACHMENT_COL2];
        filelist[2] = row[ATTACHMENT_COL3];

        var blobs = [];
        var x = 0;
        for (var i = 0; i < filelist.length; i++) {  
        //check if file ID is on sheet
          if (filelist[i] == '') 
            {
            // no ID found
            }
          else 
            {
            blobs[x] = DriveApp.getFileById(filelist[i]).getBlob();
            x++;
            }    
        }

        // @see https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
        // if you need to send emails with unicode/emoji characters change GmailApp for MailApp
        // Uncomment advanced parameters as needed (see docs for limitations)
        GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          bcc: row[BCC_COL],
          cc: row[CC_COL],
          // from: 'he@wcg.ac.uk',
          // name: 'WCG HE',
          // replyTo: 'a.reply@email.com',
          // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
          // attachments: emailTemplate.attachments
          attachments: blobs
        });
        // modify cell to record email sent date
        sheet.getRange(rowIdx + 2, emailSentColIdx + 1).setValue("Đã gửi");
        // delete used trigger
        ScriptApp.deleteTrigger(e);

      } catch(e) {
        // modify cell to record error
        sheet.getRange(rowIdx + 2, emailSentColIdx + 1).setValue([e.message]);
      }
    }
  });
}

function createTriggers() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Write sheet's information to Project Properties
  PropertiesService.getScriptProperties().setProperty("sheetId", sheet.getSheetId());

  var dataRange = sheet.getDataRange();
  var data = dataRange.getDisplayValues();
  var heads = data.shift();
  
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  var obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  var allTriggers = ScriptApp.getProjectTriggers();

  obj.forEach(function(row, rowIdx) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Attempting to create trigger for row: ' + (rowIdx + 1),'Status');

    // Skip if no "Time" to send email
    if (row[TIME_COL]) {
      var sendTime = Utilities.parseDate(row[TIME_COL], "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss");
      
      // Check if "Time" is valid and in the future
      if (sendTime > new Date()) {
        // Find existing trigger for this row (if any)
        for (var t of allTriggers) {
          var s = String(t.getUniqueId());
          if (s == row[TRIGGER_COL]) {
            ScriptApp.deleteTrigger(t);
            break;
          }
        }

        try {
          var trigger = ScriptApp.newTrigger("sendScheduledEmailForRow")
            .timeBased()
            .at(sendTime)
            .create();
          var cell = sheet.getRange(rowIdx + 2, emailSentColIdx);
          cell.setValue(String(trigger.getUniqueId()));
        } catch (e) {
          Logger.log("Error creating trigger for row " + (rowIdx + 1) + ": " + e.toString());
        }
      }
    }
  });
}