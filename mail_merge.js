function onOpen() {
  var ui=SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
  .setTitle('Instructions')
  .setWidth(300);
  ui.createMenu('Mail Merge')
  .addItem('Show instructions','instructions')
  .addItem('Send test email','sendTestEmail')
  .addItem('Send emails','sendEmail')
  
  //.addItem('test','test') //testing out functions
  
  .addToUi();
  ui.showSidebar(html);
}

//================================================================================
// functions for inline images to work
//================================================================================
//Google Apps Script lacks native replaceAll function, this one is custom built
String.prototype.replaceAll = function(search, replacement) {
        var target = this;
        return target.replace(new RegExp(search, 'g'), replacement);
};

//return an object that contains all inline images from draft message
function getInlineImg(message){
  var attachments = message.getAttachments({
    includeInlineImages: true,
    includeAttachments: false
  });
  
  return attachments;
}

//return a list of unique cid of inline images from draft message body
function cidExtract(body){
  var cids = [];
  var inlineImages = body.match(/<img [^>]*src="[^"]*"[^>]*>/gm);
  inlineImages.forEach(function(img) {
        x = img.match(/src="cid:([^"]*)"/);
        cids.push(x[1]);
  });
  return cids;
}

//return an object composed of inline images matched with their corresponding cid
//used for the inlineImages field of sendEmail()
function assignInlineImg(cids,imgs){ 
  var inlineParam ={};
  for (var i = 0; i < cids.length; i++) {
    inlineParam[cids[i]] = imgs[i];
  }
  return inlineParam;
}

//================================================================================

function test(){
  var ui=SpreadsheetApp.getUi();
  var draft = GmailApp.getDrafts()[0]; //gets the first draft in your email 
  var message=draft.getMessage(); //gets the message of your draft
  var subject=message.getSubject(); //gets the subject in your draft
  var body=message.getBody(); //gets the body of the message in your draft
  var attachment=message.getAttachments(); //gets the attachments in your draft
  
  var imgs = getInlineImg(message); //get all the inline images from draft
  var cids = cidExtract(body); //get all the unique inline images cids from draft
  var inlineParam = assignInlineImg(cids,imgs); //this is the parameter for MailApp.sendEmail(...,inlineImages:inlineParam)
      
  var response = ui.alert("body: "+body+"\n"
                          +"cids: "+cids+"\n"
                          , ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) {
    throw ("exit successfully");
  }
 
  var ss=SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Mail Merge Sheet');
  var fname=sheet.getRange(3,6).getValue();
  var lname=sheet.getRange(3,7).getValue();
  var email = sheet.getRange(3,8).getValue();
  var subject_new = subject.replaceAll("{{FirstName}}",fname).replaceAll("{{LastName}}",lname);
  var body_new = body.replaceAll("{{FirstName}}",fname).replaceAll("{{LastName}}",lname);
 
  MailApp.sendEmail(
    email,
    subject_new,
    ' ',
    {
     htmlBody: body_new,
     attachments:attachment,    
     inlineImages:inlineParam
    }
  );
  
  //flag for completing all the codes above
  var ss=SpreadsheetApp.getActive();
  SpreadsheetApp.getActiveSheet().getRange(1,10).setValue("Test complete");
}

////////////////////////////////////////////////////////////////////////

function instructions(){
  var ui=SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
  .setTitle('Instructions')
  .setWidth(300);
  ui.showSidebar(html);
}

function sendTestEmail() {
  var draft = GmailApp.getDrafts()[0]; //gets the first draft in your email 
  var message=draft.getMessage(); //gets the message of your draft
  var subject=message.getSubject(); //gets the subject in your draft
  var body=message.getBody(); //gets the body of the message in your draft
  var attachment=message.getAttachments(); //gets the attachments in your draft
  
  var imgs = getInlineImg(message); //get all the inline images from draft
  var cids = cidExtract(body); //get all the unique inline images cids from draft
  var inlineParam = assignInlineImg(cids,imgs); //this is the parameter for MailApp.sendEmail(...,inlineImages:inlineParam)
  
  var ss=SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Mail Merge Sheet');
  var fname = sheet.getRange(3,6).getValue();
  var lname = sheet.getRange(3,7).getValue();
  var email = sheet.getRange(3,8).getValue();
  var subject_new = subject.replaceAll("{{FirstName}}",fname).replaceAll("{{LastName}}",lname);
  var body_new = body.replaceAll("{{FirstName}}",fname).replaceAll("{{LastName}}",lname);
  
  MailApp.sendEmail(
    email,
    subject_new,
    ' ',
    {
     htmlBody: body_new,
     attachments:attachment,    
     inlineImages:inlineParam
    }
  );
  
  SpreadsheetApp.getActiveSheet().getRange(3,9).setValue("Sent");
}

function sendEmail() {
  var ui=SpreadsheetApp.getUi();
  var confirm=ui.alert("Please confirm","Do you want to proceed with the mail merge? This cannot be undone.",ui.ButtonSet.YES_NO);
  switch(confirm){
    case ui.Button.YES:  
      var draft = GmailApp.getDrafts()[0]; //gets the first draft in your email
      var message=draft.getMessage(); //gets the message of your draft
      var subject=message.getSubject(); //gets the subject in your draft
      var body=message.getBody(); //gets the body of the message in your draft
      var attachment=message.getAttachments(); //gets the attachments in your draft
      
      var imgs = getInlineImg(message); //get all the inline images from draft
      var cids = cidExtract(body); //get all the unique inline images cids from draft
      var inlineParam = assignInlineImg(cids,imgs); //this is the parameter for MailApp.sendEmail(...,inlineImages:inlineParam)
      
      var ss=SpreadsheetApp.getActive();
      var sheet = ss.getSheetByName('Mail Merge Sheet');
      var last = sheet.getLastRow();

      
      for(var i=2; i<last+1; i++){
        var fname=sheet.getRange(i,1).getValue();
        var lname=sheet.getRange(i,2).getValue();
        var email = sheet.getRange(i,3).getValue();
        var subject_new = subject.replaceAll("{{FirstName}}",fname).replaceAll("{{LastName}}",lname);
        var body_new = body.replaceAll("{{FirstName}}",fname).replaceAll("{{LastName}}",lname);

        MailApp.sendEmail(
          email,
          subject_new,
          ' ',
          {
            htmlBody: body_new,
            attachments:attachment,    
            inlineImages:inlineParam
          }
        );
        
        SpreadsheetApp.getActiveSheet().getRange(i,4).setValue("Sent");
      }
      ui.alert("Mail merge completed")
      break;
    default:
      ui.alert("Mail merge did not complete");
      break;
  }
}
