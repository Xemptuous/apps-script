//EXAMPLE:
function test() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheets()[0];
    let error = error_test; //ReferenceError
  }
  catch (Exception) {
    //declaring new Error here so that the error point to this line of code.
    let err = new Error();
    errorHandlerLog(Exception, err, ss, sh)
  
    //or if you don't want to log the Spreadsheet and Sheet:
    //errorHandlerLog(Exception, err)
  }  
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
