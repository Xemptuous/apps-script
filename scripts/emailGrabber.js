/* Grabs all email attachments that meet certain criteria
In this case, standalone pdfs or those within .zips.
*/
function allBatchCalls() {
  // performing all batch thread calls here to avoid time-delays with "newer_than" Gmail operator
  const firstEmail  = [ GmailApp.search("newer_than:1h AND from:@EMAIL1.COM", 0, 50), "NAME1" ];
  const secondEmail = [ GmailApp.search("newer_than:1h AND {replyto:@EMAIL2.COM, EMAIL2.5.COM}", 0, 50), "NAME2" ];
  [...]
  
  let threads = [];
  let threadList = [];
  threadList.push(firstEmail, secondEmail, ..., lastEmail);
  
  // filtering out calls with no responses
  for (let i = 0; i < threadList.length; i++) {
    if (threadList[i][0].length) {
      threads.push(threadList[i]);
    }
  }
  threadList = null;
  if (!threads.length) { return; }
  
  try {
    for (let i = 0; i < threads.length; i++) {
      getBatchEmails(threads[i][0], threads[i][1]);
    }
  }
  catch (Exception) {
    let err = new Error();
    errorHandlerLog(Exception, err);
  }
}


function getBatchEmails(threads, name) {
  let employees = getEmployeeEmails();
  let total = 0, count = 0;
  let dupe_names = [];
  try {
    const folderId = getFolderId(name);
    if (!folderId) { return; }
    const attachmentFolder = DriveApp.getFolderById(folderId);
    const msgs = GmailApp.getMessagesForThreads(threads);
    
    // getting all files recursively to ensure we don't make duplicates.
    const allFiles = recursiveListFiles(folderId)
    
    for (let i = 0; i < msgs.length; i++) {
      for (let j = 0; j < msgs[i].length; j++) {
        let data = msgs[i][j];
        let efrom = data.getFrom();
        // ensuring we don't grab outbound emails within email chains from our own employees.
        if (!employees.some(el => efrom.includes(el)) {
          let attachments = data.getAttachments();
          for (let k = 0; k < attachments.length; k++) {
            let blob = attachments[k].copyBlob();
            if (blob.getContentType().includes("pdf") {
              total += 1;
              if (!allFiles.includes(attachments[k].getName())) {
                attachmentFolder.createFile(blob);
                count += 1;
                continue;
              }
              dupe_names.push(attachments[k].getName().toString());
            }
            else if (blob.getContentType().includes("zip")) {
              let zipBlob = blob.setContentTypeFromExtension();
              let unzipped = Utilities.unzip(zipBlobs);
              for (let l = 0; l < unzippedBlobs.length; l++) {
                if (unzipped[l].getContentType().includes("pdf")) {
                  total += 1;
                  if (!allFiles.includes(unzipped[l].getName())) {
                    attachmentFolder.createFile(unzipped[l]);
                    count += 1;
                    continue;
                  }
                  dupe_names.push(unzipped[l].getName().toString());
                }
              }
            }
          }
        }
      }
    }
  }
  catch (Exception) {
    let err = new Error();
    errorHandlerLogger(Exception, err);
    return;
  }
  finally {
    if (total) {
      if (dupe_names.length) {
        writeToSheet(name, total, count, dupe_names.length, dupe_names);
        return;
      }
      writeToSheet(name, total, count, 0);
    }
  }
}


function getFolderId(name) {
  // grab parent folder that holds folders with unique names
  const driveFolders = DriveApp.getFolderById("FOLDER_ID_HERE").getFolders();
  while (driveFolders.hasNext()) {
    let folder = driveFolders.next();
    if (folder.getName().toLowerCase() === name.toLowerCase()) {
      return folder.getId();
    }
  }
  return false;
}

/**
* Assumes there is a spreadsheet containing master data on employees that engage in back-and-forth communications.
*/
function getEmployeeEmails() {
  // various titles/positions of employees to be grabbed.
  const conditions = ["SPA", "SPA Team Lead", "SR. SPA"];
  const ss = SpreadsheetApp.openById("ID_HERE");
  const sh = ss.getSheets()[0]; // first sheet
  const data = sh.getDataRange().getValues();
  
  // column names should include names below
  const r_col = data[0].indexOf("Role");
  const e_col = data[0].indexOf("Email");
  
  return data.filter(row => conditions.some(el => row[r_col] == el)).map(a => a[e_col]);
}


/**
* Recursive function that lists all files within the specified folder, as well as all subfolders.
* @param {string} id - the ID of the parent folder to perform the recursive search on. 
*/
function recursiveListFiles(id) {
  const _FOLDERS = DriveApp.getFolderById(id).getFolders();
  const _FILES = DriveApp.getFolderById(id).getFiles();
  let allFiles = [];
  
  // pushes immediate files in parent folder into allFiles list.
  while (_FILES.hasNext()) {
    let file = _FILES.next();
    let name = file.getName();
    allFiles.push(name);
  }
  // iterate through every folder in the parent folder. 
  while (_FOLDERS.hasNext()) {
    let foldersNext = _FOLDERS.next();
    // iterate through every subfolder and get all files and subfolders.
    listFilesAndSubfolders(foldersNext);
  }
  function listFilesAndSubfolders(folder) {
    // get all files in current folder.
    listFiles(folder);
    // go through each subfolder in current folder.
    let subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      let subfolder = subfolders.next();
      // run this function again on all subfolders inside the given folder.
      listFilesAndSubfolders(subfolder);
    }
  }
  // get all files in the given folder.
  function listFiles(foldersNext) {
    let files = foldersNext.getFiles();
    while (files.hasNext()) {
      let file = files.next();
      let name = file.getName();
      allFiles.push(name);
    }
  }
  return allFiles;
}


/**
* Writes the data gathered from the email scripts to a Log file in the Drive.
* @param {string} name - The name of the folder within the parent; also the uid-name for this email
* @param {number} total - The total number of emails looped through that match all conditional criteria.
* @param {number} count - The number of files which were created into the drive from the script.
* @param {number} [duplicates=0] - The number opf duplicate files which were found but not created. 
* @param {string[]} names - An array of strings representing the names of duplicate files.
*/
function writeToSheet(name, total, count, duplicates, names) {
  let date = new Date();
  let month = Utilities.formatDate(date, "America/Los_Angeles", "MMM");
  let year = date.getFullYear();
  const ss = SpreadsheetApp.openById("ID_HERE");
  const sh = ss.getSheets()[0];
  if names() {
    names = names.join(", ");
    sh.appendRow([date, month, year, name, total, count, duplicates, names]);
    return;
  }
  sh.appendRow([date, month, name, total, count, duplicates]);
}


/**
 * Handles try-catch errors and appends the data to another spreadsheet for logging purposes.
 * @typedef Exception
 * @type {Object}
 * @param {Exception} x - The caught exception within the try-catch block.
 * @param {Error} err - The new Error class object instance.
 * @param {SpreadsheetApp.Spreadsheet} [xss="Standalone Script"] - The active spreadsheet. If left blank, defaults to `Standalone Script`.
 * @param {SpreadsheetApp.Sheet} [xsh="Standalone Script"] - The spreadsheet's active sheet. If left blank, defaults to `Standalone Script`.
 */
function errorHandlerLogger(x, err, xss, xsh)
{
  //Change this to the spreadsheet you want to use as your log
  const ss = SpreadsheetApp.openById("YOUR LOGGING SHEET");
  const sh = ss.getSheets()[0];
  const timestamp = Utilities.formatDate(new Date(), "America/Los_Angeles", "MM/dd/yyyy hh:mm:ss a");
  
  //Grabs the name of the script this log is attached to
  const script    = DriveApp.getFileById(ScriptApp.getScriptId()).getName();
  const xName     = x.toString().split(":")[0].trim();
  const xDetails  = x.toString().split(":")[1].trim();
  
  //Appends rows to the ss
  //Change format of rows to how you want it laid out
  
  if (!xss || !xsh) {
    xss = "Standalone Script", xsh = "Standalone Script";
    sh.appendRow([timestamp, script, xss, xsh, xName, xDetails, err.stack]);
    return;
  }
  sh.appendRow([timestamp, script, xss.getName(), xsh.getName(), xName, xDetails, err.stack]);
}
