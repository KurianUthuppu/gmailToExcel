# gmail_To_Excel
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
