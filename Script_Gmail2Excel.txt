// Modified from http://pipetree.com/qmacro/blog/2011/10/automated-email-to-task-mechanism-with-google-apps-script/

// Globals, constants
var LABEL_PENDING = "pending_requirement";
var LABEL_DONE = "done_requirement";
var parentFolder = DriveApp.getFolderById('1FGOsBXofYYcsJC0nqtIzHbfAslhLK-P2')
var fileTypesToExtract = ['jpg', 'jpeg', 'tif', 'png', 'gif', 'bmp', 'svg'];

// processPending(sheet)
// Process any pending emails and then move them to done
function processPending_(sheet) {

  
  // Get out labels by name
  var label_pending = GmailApp.getUserLabelByName(LABEL_PENDING);
  var label_done = GmailApp.createLabel(LABEL_DONE);

  // The threads currently assigned to the 'pending' label
  var threads = label_pending.getThreads();
    
  // Process each one in turn, assuming there's only a single
  // message in each thread
  for (var t in threads) {
      
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
  
    var thread = threads[t];

    // Gets the message body
    var message = thread.getMessages()[0];
    var content = thread.getMessages()[0].getPlainBody();
    var attachments = message.getAttachments();
    
    // TODO: Process the messages here
    
    // Add message to sheet 
    date_cell = sheet.getRange(lr+1, 2)
    date_cell.setValue(message.getDate());
    
    username = message.getFrom();

    from_cell = sheet.getRange(lr+1, 3)
    from_cell.setValue(username);
    
    regExp = new RegExp("(?<=" + "Customer:" + ").*");
    customer = content.match(regExp);
    if (customer === null){
      customer = '';
    } else {
      customer = customer.toString().trim();
      customer_cell = sheet.getRange(lr+1, 4);
      customer_cell.setValue(customer);
    }
        
    regExp = new RegExp("(?<=" + "Requirement:" + ").*");
    requirement = content.match(regExp);
    if (requirement === null){
      customer = '';
    } else {
    requirement = requirement.toString().trim();
    requirement_cell = sheet.getRange(lr+1, 5);
    requirement_cell.setValue(requirement);
    }
    
    regExp = new RegExp("(?<=" + "Purpose:" + ").*");
    purpose = content.match(regExp);
    if (purpose === null){
      customer = '';
    } else {
    purpose = purpose.toString().trim();
    purpose_cell = sheet.getRange(lr+1, 6);
    purpose_cell.setValue(purpose);   
    }

    i = 0;
    var root = DriveApp.getRootFolder();

    for(var k in attachments){
      var attachment = attachments[k];
      var isDefinedType = checkIfDefinedType_(attachment);
      if(!isDefinedType) continue;
      var attachmentBlob = attachment.copyBlob();
      var file = DriveApp.createFile(attachmentBlob).setName('SR_'+rid+'-'+k+'_'+username);
      parentFolder.addFile(file);
      stored_file = DriveApp.getFolderById('1FGOsBXofYYcsJC0nqtIzHbfAslhLK-P2').getFilesByName(file);

    // Get the id of the saved attachment
    while (stored_file.hasNext()) {
      var file = stored_file.next();
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      image_file_id = file.getId();
    };

    // Generate the image link and store the value
    image_link = 'https://drive.google.com/file/d/'+image_file_id+'/view';

    attachment_cell = sheet.getRange(lr+1, i+7);
    attachment_cell.setValue(image_link);
    root.removeFile(file);
    i += 1;
    
    }

    thread.removeLabel(label_pending); 
    thread.addLabel(label_done);
    if(i>0){
      var attachment_Status = "Yes";
    } else {
      var attachment_Status = "No";
    }

    message.replyAll("",{
    htmlBody: "Hello,<br/>We have recieved your requirement with the below details:"
    +"<br/><b>Customer</b>: "+customer+"<br/><b>Requirement</b>: "+requirement+"<br/><b>Purpose</b>: "+purpose
    +"<br/><b>Attachment</b>: "+attachment_Status
    +"<br/><br/><b>RID</b>: "+ Utilities.formatString("%06d", rid)+"<br/> Please save this RID (Requirement Id) for any future correspondence."
    +"<br/><br/>Now you please sit back, relax and enjoy your day as the team has started to cook your requirement right away..!"
    +"<br/><br/>Regards,<br/>SPEG Team",
    noReply: true
    })
    
  }
}

function checkIfDefinedType_(attachment){
  var fileName = attachment.getName();
  var temp = fileName.split('.');
  var fileExtension = temp[temp.length-1].toLowerCase();
  if(fileTypesToExtract.indexOf(fileExtension) !== -1) return true;
  else return false;
}

// main()
// Starter function; to be scheduled regularly
function main_emailDataToSpreadsheet() {
  // Get the active spreadsheet and make sure the first
  // sheet is the active one
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.setActiveSheet(ss.getSheets()[0]);

  // Process the pending emails
  processPending_(sh);
}