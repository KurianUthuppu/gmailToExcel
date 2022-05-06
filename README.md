# gmail_To_Excel (Work-In-Progress)
Run in an hourly manner to capture details from emails being sent with a specific subject which gets labelled and then write it into an associated excel file & send reply to the user, and finally putting a completed label on the specific email
- In this specific context, I am trying to capture details such as 'customer', 'requirements', 'purpose' raised by the users via email with the subject 'Requirement'

### Requirements
* Valid Google account
* Browser - Chrome / Firefox

### Resources
- Google excel sheets
- Google scripts - https://script.google.com/u/1/home/start

### Setup
- Create a new excel sheet and fill in the requisite column headers
- Go to Extensions and click Apps script, this will create a new apps script file
- You may use the new apps scripts editor as per your comfort; I found the new editor to be much more user-friendly

### Code
#### Declaring global constants
- Declaring global constants that are to be used in the rest of the code
- The incoming emails that matches the specific subject and content criteria are labelled as 'pending requirement'
- Those emails whose details are copied to the excel sheet are labelled as 'done requiement'
```
var LABEL_PENDING = "pending_requirement";
var LABEL_DONE = "done_requirement";
var parentFolder = DriveApp.getFolderById('1KGOsAXofZZasJC0nqtOzHbfBslhLR-P9')
var fileTypesToExtract = ['jpg', 'jpeg', 'tif', 'png', 'gif', 'bmp', 'svg'];
```
#### Setting the active excel sheet
- Getting the active spreadsheet and calling the requisite funcion
```
// Starter function; to be scheduled regularly
function main_emailDataToSpreadsheet() {
  // Get the active spreadsheet and make sure the first
  // sheet is the active one
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.setActiveSheet(ss.getSheets()[0]);

  // Process the pending emails
  processPending_(sh);
}
```

#### Getting the requisite email threads
- Declare the processPending_ function and inserting the requisite code
- Finding the requisite emails that are marked with the label
- Extracting the requisite email threads
```
// Get out labels by name
  var label_pending = GmailApp.getUserLabelByName(LABEL_PENDING);
  var label_done = GmailApp.createLabel(LABEL_DONE);

// The threads currently assigned to the 'pending' label
  var threads = label_pending.getThreads();
```

#### Reading the id of the last requirement raised from the excel sheet
```
   // Read the last SRID    
    lr = sheet.getLastRow();
    lc = sheet.getLastColumn();
  
    if (lr > 1) {
    rid_cell_hist = sheet.getRange(lr, 1);
    rid_cell = sheet.getRange(lr+1, 1); 
    rid = rid_cell_hist.getValue()+1;
    rid_cell.setValue(rid).setNumberFormat("000000");
    } else {
    rid_cell = sheet.getRange(lr+1, 1); 
    rid = 1;
    rid_cell.setValue(rid).setNumberFormat("000000");
    } 
```

#### Getting the message body and attachments
- Assigning the first message in the thread to the variable
- Declare variables to capture the message body and attachments
```
var thread = threads[t];

// Gets the message body
var message = thread.getMessages()[0];
var content = thread.getMessages()[0].getPlainBody();
var attachments = message.getAttachments();
```

#### Writing date and username/email-id to the excel sheet
```
 // Add message to sheet 
 date_cell = sheet.getRange(lr+1, 2)
 date_cell.setValue(message.getDate());
    
 username = message.getFrom();

 from_cell = sheet.getRange(lr+1, 3)
 from_cell.setValue(username);
```

#### Capturing the requisite phrases shared by the user
- Regexp formula is used to match with the 'catch' (Word to be matched) word and then extract the content shared by used against the same
- "i" is used to make the formula case insensitive
- The same formula is repeated to capture other relevant content including requirement, purpose etc:-
```
    regExp = new RegExp("(?<=" + "Customer:" + ").*","i");
    customer = content.match(regExp);
    if (customer === null){
      customer = '';
    } else {
      customer = customer.toString().trim();
      customer_cell = sheet.getRange(lr+1, 4);
      customer_cell.setValue(customer);
    }
```

#### Saving the attachments in the requisite folder
- Retrieve all the attachments
- Store the attachments with the requisite filename in the rquisite folder
- Setting the requisite access permissions for the file
```
    var root = DriveApp.getRootFolder();

    for(var k in attachments){
      var attachment = attachments[k];
      var isDefinedType = checkIfDefinedType_(attachment);
      if(!isDefinedType) continue;
      var attachmentBlob = attachment.copyBlob();
      var file = DriveApp.createFile(attachmentBlob).setName('SR_'+rid+'-'+k+'_'+username);
      parentFolder.addFile(file);
      stored_file = DriveApp.getFolderById('1ABOsBXogYYcsJC0nqtIyHbfAslhLM-Z9').getFilesByName(file);

    // Get the id of the saved attachment
    while (stored_file.hasNext()) {
      var file = stored_file.next();
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      image_file_id = file.getId();
    };
```

#### Writing the saved filename link to the excel sheet
- Write the link to the file in the excel
```
   image_link = 'https://drive.google.com/file/d/'+image_file_id+'/view';

    attachment_cell = sheet.getRange(lr+1, i+7);
    attachment_cell.setValue(image_link);
    root.removeFile(file);
    i += 1;
    
    }
```

#### Replying back to the email request
- Using the noReply option to reply back in the same email 
```
    message.replyAll("",{
    htmlBody: "Hello,<br/>We have recieved your requirement with the below details:"
    +"<br/><b>Customer</b>: "+customer+"<br/><b>Requirement</b>: "+requirement+"<br/><b>Purpose</b>: "+purpose
    +"<br/><b>Attachment</b>: "+attachment_Status
    +"<br/><br/><b>RID</b>: "+ Utilities.formatString("%06d", rid)+"<br/> Please save this RID (Requirement Id) for any future correspondence."
    +"<br/><br/>Now you please sit back, relax and enjoy your day as the team has started to cook your requirement right away..!"
    +"<br/><br/>Regards,<br/>SPEG Team",
    noReply: true
    })
```

You may find the full script in the file - Gmail2Excel 
