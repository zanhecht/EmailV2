/* Send Form by Email V2 */

/* Contact the author at http://email.ahecht.com */
 
/* Originally based on the script by Amit Agarwal at http://www.labnol.org/?p=20884 */
 
function Initialize() {
     
    var triggers = ScriptApp.getScriptTriggers();
    
	for(var i in triggers) {
		ScriptApp.deleteTrigger(triggers[i]);
    }
	
    ScriptApp.newTrigger("SendGoogleForm")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
}
     
function SendGoogleForm(e) {
    try
    {
		// Email address to send the message to.
        // You may also replace this with another email address.
		var email = Session.getActiveUser().getEmail();
        // var email = "";
		
        // You may replace SheetName with another name.
        // Set to "" to disable putting the sheet name in the email subject.
        var SheetName = SpreadsheetApp.getActiveSpreadsheet().getName().replace(/\s*\(.*?\)\s*/g, '')
        //var SheetName = "";
        
        // Change the following variables to have a custom default subject
		// reply-to email address, and sent from name for Google Docs emails
		var form_subject = "Form Submission"
        var reply_to="noreply@google.com";
        var sent_by="Google Docs Form";

        // This will set the reply_to email address, sender name, and
        // subject to the values entered in the 'Email Address', 'Name'
        // and 'Subject' fields, if they exist.
        // Change the following variables to use custom field names.
        var email_field = 'Email Address';
        var name_field = 'Name';
        var subject_field = 'Subject';
      
        /* DO NOT EDIT ANYTHING BELOW THIS LINE */
		
		var s = SpreadsheetApp.getActiveSheet();
		var headers = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
		var message = "";
        var htmlMessage = "";
		
		for(var i in headers) {
	        message += headers[i] + ' :: '+ e.namedValues[headers[i]].toString() + "\n\n";
            htmlMessage += '<p><b>' + headers[i] + ':</b><br />'
            htmlMessage += e.namedValues[headers[i]].toString().replace(/\n/g, "<br />") + '</p>';
		}
		
        if(e.namedValues[subject_field][0]) {
            form_subject = e.namedValues[subject_field][0];
        }
      
        
        if(SheetName) {   
            subject = "[" + SheetName + "] " + form_subject;
        } else {
            subject = form_subject;
        }
        
          
        if(e.namedValues[email_field][0]) {
            reply_to = e.namedValues[email_field][0];
            htmlMessage += "<p><b>Reply:</b><br /><a href='mailto:" + reply_to;
            htmlMessage += "?subject=" + encodeURIComponent("Re: " + form_subject).replace(/\'/g, "%27");
            htmlMessage += "&body=" + encodeURIComponent("====FORM SUBMISSION====\n\n" + message).replace(/\'/g, "%27") + "'>Click Here</a></p>";
        }
        
        if(e.namedValues[name_field][0]) {
            sent_by = e.namedValues[name_field][0];
        } else if(reply_to!="noreply@google.com") {
            sent_by = reply_to;
        }
      
    
        message += "Sheet URL :: " + SpreadsheetApp.getActiveSpreadsheet().getUrl() + "\n\n";
        htmlMessage += '<p><b>Sheet URL:</b><br /><a href="'
        htmlMessage += SpreadsheetApp.getActiveSpreadsheet().getUrl() +'">'
        htmlMessage += SpreadsheetApp.getActiveSpreadsheet().getUrl() + '</p>';
        
        // This is the MailApp service of Google Apps Script
        // that sends the email. You can also use GmailApp here.
        MailApp.sendEmail(email, subject, message, {
            name: sent_by,
            replyTo: reply_to,
            htmlBody: htmlMessage
        });
		
    } catch (e) {
		Logger.log(e.toString());
    }
	
}