// GLOBAL Variables
var uniqueValues = [];
var shareWith =[];
var permisions = [];
var sendEmail = [];
var colHeader;
var filterColumnNumber;
var newFileName;
var gDriveFolderID;
var gDriveNewLocation;
var sheetList = []
var selectedSheet;
var report =[];
var formats =[];

// NOTE ----------------------------------------------
// If you activate the variable below, it breaks the addon...
// Seems like the variable calls for a function before oAuthMode is read.
// If you activate it the onOpen and the onEdit stop working... probably as a security meassure.
// I placed "sheetList = sheetnms()" within the onEdit function
// var sheetList = sheetnms(); // do not activate here.


// Constants
const SetupSheet = 'S&S Parameters';
const sharingOptions = ['Edit', 'View Only', 'Make Comments'];
const notificationOptions = ['Yes - Send', ''];

const uniqueCellNote = 'This column will autofill. A new file will be generated for each unique value displayed here. You can remove some of them. Minimum: 1 value needed';
const shareCellNote = '(optional) Email address of the person to share with (multiple emails must be separated by ", " (comma and space)';
const permisionCellNote = '(optional) Type of sharing permissions. Required ONLY if you have entered an email in the same Row of "Share with User" (Column E)';
const notificationCellNote = '(optional) Standard notification that a new file has been shared with them. If left empty, no notification will be sent';
const colHeaderCellNote = 'Column with unique criteria used to create new files';
const sheetColCellNote = 'Name of the sheet, within this spreadsheet, containing the data you want copied/split into new files';
const newFileNameCellNote = 'Name you want to give to new files (will be followed by a unique identifier)';
const folderNameCellNote = 'Google Drive folder ID  where you want all new files to be generated (you must have editors rights on this folder). Numbers and letters following: https://drive.google.com/drive/folders/...';
const newLocationCellNote = 'A different Google Drive folder ID where this specific new file should be placed (you must have editors rights on this folder)';

function setUpInicial(){
try {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var yourNewSheet = activeSpreadsheet.getSheetByName(SetupSheet);
  
  var sheetList = [];
  activeSpreadsheet.getSheets().forEach(function(val){sheetList.push(val.getName())});
    
  if (yourNewSheet === null) {
    var addOnSheet = activeSpreadsheet.insertSheet().setName(SetupSheet).setTabColor('#000000')
    
    // 1. Headers: Specific headers I want created + frozen rows
    var headersH1 =[['SPLIT PARAMETERS (REQUIRED)',,,,'SHARE PARAMETERS (OPTIONAL)',,,,,'ACTIVITY REPORT']]
    var headersH2 = [['Fill in each concept below',,'Unique Values:','(a) Share with user/s','(a.1) Sharing Permisions:','(a.2) Send Notifications:','(b) Different gDrive Location','','Timestamp','New Spreadsheet URL','Sheet with data','Column Header','Unique Value in file','# Rows' ,'# Columns','Shared with','Sharing Permisions','Notification','Full name of new file','Google Drive Folder']]; 
    var headersV = [['Sheet with data:'],['Column Header (with unique criteria):'],['Name for your new files:'], ['Google Drive Folder ID:']]
    addOnSheet.getRange(1, 1, headersH1.length, headersH1[0].length).setValues(headersH1);   
    addOnSheet.getRange(2, 2, headersH2.length, headersH2[0].length).setValues(headersH2);   
    addOnSheet.getRange(3, 1, headersV.length, headersV[0].length).setValues(headersV);   
    
       
      
    // 2. Constant Data Validation 
    addOnSheet.getRange('F3:F').setDataValidation(SpreadsheetApp.newDataValidation()
                                                .setAllowInvalid(false)
                                                .requireValueInList(sharingOptions)
                                                .build()
                                                );    
    
    addOnSheet.getRange('G3:G').setDataValidation(SpreadsheetApp.newDataValidation()
                                                .setAllowInvalid(false)
                                                .requireValueInList(notificationOptions)
                                                .build()
                                                );    
    
    // 3. Initial Data Validation that will be later on updated by the onEdit() function
    addOnSheet.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation()
                                                .setAllowInvalid(true)
                                                .requireValueInList(sheetList)
                                                .build()
                                                );
    
    
    
    // 4. Look & Feel Customization 
    addOnSheet.setFrozenRows(2);
    addOnSheet.getRange('E1:H1').activate();
    addOnSheet.getActiveRangeList().setBackground('#c9daf8');
    addOnSheet.getRange('E2:G2').activate();
    addOnSheet.getActiveRangeList().setBackground('#6fa8dc');
    addOnSheet.getRange('H2').activate();
    addOnSheet.getActiveRangeList().setBackground('#3d85c6');
    addOnSheet.getRange('E1:H1').activate().merge().setHorizontalAlignment('center');
    addOnSheet.getRange('B2:H2').activate().setHorizontalAlignment('center');

    addOnSheet.getRange('J1:U1').activate();
    addOnSheet.getActiveRangeList().setBackground('#b6d7a8');
    addOnSheet.getRange('J2:U2').activate();
    addOnSheet.getActiveRangeList().setBackground('#93c47d');
    
    addOnSheet.getRange('A1:D1').activate();
    addOnSheet.getActiveRangeList().setBackground('#d9d9d9');
    addOnSheet.getRange('A2:D2').activate();
    addOnSheet.getActiveRangeList().setBackground('#000000').setFontColor('#FFFFFF')
    addOnSheet.getRange('B3:B6').activate();
    addOnSheet.getActiveRangeList().setBackground('#efefef');
    addOnSheet.getRange('A1:D1').activate().merge().setHorizontalAlignment('center');
        
    addOnSheet.getRange('F:G').activate();
    addOnSheet.getActiveRangeList().setHorizontalAlignment('center');  
    
    addOnSheet.getRange('B1:B4').activate();
    addOnSheet.getActiveRangeList().setHorizontalAlignment('center');
    addOnSheet.getRange('B5:B6').activate();
    addOnSheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    
    addOnSheet.getRange('A3:A').activate();
    addOnSheet.getActiveRangeList().setHorizontalAlignment('right');  
    
    addOnSheet.getRange('D3:H').activate();
    addOnSheet.getActiveRangeList().setBorder(null, null, null, null, null, true, '#000000', SpreadsheetApp.BorderStyle.DASHED);
    
    addOnSheet.getRange(1, 1, addOnSheet.getMaxRows(), addOnSheet.getMaxColumns()).activate();
    addOnSheet.autoResizeColumns(1, addOnSheet.getMaxColumns());
    addOnSheet.setColumnWidth(3, 30);
    addOnSheet.setColumnWidth(5, 175);
    addOnSheet.setColumnWidth(9, 30);
    addOnSheet.setColumnWidth(10, 120);
    addOnSheet.setColumnWidth(22, 20);
    addOnSheet.setColumnWidth(23, 20);
    addOnSheet.setColumnWidth(24, 20);
    addOnSheet.setColumnWidth(25, 20);
    addOnSheet.setColumnWidth(26, 20);
    
    addOnSheet.getRange('A3').setNote(sheetColCellNote);
    addOnSheet.getRange('A4').setNote(colHeaderCellNote);
    addOnSheet.getRange('A5').setNote(newFileNameCellNote);
    addOnSheet.getRange('A6').setNote(folderNameCellNote);
    addOnSheet.getRange('D2').setNote(uniqueCellNote);
    addOnSheet.getRange('E2').setNote(shareCellNote);
    addOnSheet.getRange('F2').setNote(permisionCellNote);
    addOnSheet.getRange('G2').setNote(notificationCellNote);
    addOnSheet.getRange('H2').setNote(newLocationCellNote);
    
    addOnSheet.getRange('B10').activate();

    
    // 5. All Logs
//     Logger.log("SheetList ="+sheetList);
//     Logger.log("HeadersH1 =" + headersH1);
//     Logger.log("HeadersH2 =" + headersH2);
//     Logger.log("HeadersV =" + headersV);
//     Logger.log("HeadersV Length =" + headersV.length);
//     Logger.log("HeadersV 0 =" + headersV[0]);
//     Logger.log("HeadersV 0 Length =" + headersV[0].length); 
    
  } else {
   // Go to the parameters sheet
  activeSpreadsheet.getSheetByName(SetupSheet).getRange('B3:B6').activate();
  
  }
      
  
//  var mainSSID = activeSpreadsheet.getId();


} catch (e) {
  Logger.log("Error logged (setUpInicial): "+e)
}


}

function runAddOn(){
try{  
 var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 var addOnSheet = activeSpreadsheet.getSheetByName(SetupSheet);
   
  if (addOnSheet === null) {
    SpreadsheetApp.getUi().alert('Please run the "Initial Setup" first', '(Step 1) Run the Initial Setup.  (Step 2) Fill in parameters', SpreadsheetApp.getUi().ButtonSet.OK) 
    
  } else { 
  
    // var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();  
    var activeSSID = activeSpreadsheet.getId();
    
    // 1. Get Input Parameters needed
    var thisSheet = addOnSheet.getRange("B3").getValue();
    var colHeader = addOnSheet.getRange("B4").getValue();
    var filterColumnNumber = headerIndex(addOnSheet, colHeader);
    var newFileName = addOnSheet.getRange("B5").getDisplayValue();
    var gDriveFolderID = addOnSheet.getRange("B6").getDisplayValue(); // rather than getValue() which returns undefined // Source: https://stackoverflow.com/questions/34691425/difference-between-getvalue-and-getdisplayvalue-on-google-app-script
    
    var uniqueValues = addOnSheet.getRange("D3:D").getValues().filter(String);
    var shareWith = addOnSheet.getRange("E3:E").getValues();
    var permisions = addOnSheet.getRange("F3:F").getValues();
    var sendEmail = addOnSheet.getRange("G3:G").getValues(); 
    var gDriveNewLocation = addOnSheet.getRange("H3:H").getValues(); 
       
    
    // 2. Verify that all required parameters are full  
    
   if(uniqueValues == "" || filterColumnNumber == "" || thisSheet == "" || newFileName == "" || gDriveFolderID == ""){
     goToParameters(addOnSheet) 
    } else{
//      Logger.log("No Alert needed: Run");
//      Logger.log("uniqueValues = "+uniqueValues);
//      Logger.log("shareWith = " + shareWith);
//      Logger.log("permisions = " + permisions);
//      Logger.log("sendEmail = " + sendEmail);
//      Logger.log("filterColumnNumber = " + filterColumnNumber);
//      Logger.log("thisSheet = " + thisSheet);    
//      Logger.log("newFileName = " + newFileName);
//      Logger.log("gDriveFolderID = " + gDriveFolderID);
//      Logger.log("gDriveNewLocation = " + gDriveNewLocation);
      
      // 3. Verify no Filter is active in the selected Sheet
      if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(thisSheet).getFilter()) {  //returns null if fitler does not exist
        Logger.log('No previous filter'); 
      }
      else {
        // ACTIVE filter >> remove filter
        Logger.log('Removing previous filter');
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(thisSheet).getFilter().remove()
      }
      
      // 4. Run Add-on with parameters  
      MakeCopyPerUniqueValue(activeSSID,thisSheet,uniqueValues,shareWith,permisions,filterColumnNumber,gDriveFolderID,newFileName,sendEmail, colHeader,gDriveNewLocation )
    
      // 5. Paste Report
      pasteAndSortReport(addOnSheet, report);
 
    }
    
  } 
} catch (e) {
  Logger.log("Error logged (runAddOn): "+e)
}
  
}


function goToParameters(addOnSheet){  
  if (addOnSheet === null) {
//    Logger.log('Please run the "Initial Setup" first' );
    // SpreadsheetApp.getUi().alert('Please run the "Initial Setup" first', '(Step 1) Run the Initial Setup.  (Step 2) Fill in the required information', SpreadsheetApp.getUi().ButtonSet.OK) 
    
  } else {  
    addOnSheet.getRange('B3:B6').activate();  
  SpreadsheetApp.getUi().alert('Required Parameters are missing', 'Please fill all 4 required parameters in column B (sheet: '+SetupSheet+')', SpreadsheetApp.getUi().ButtonSet.OK)
 }
}


// Finds the Index value for the specified column header.
function headerIndex (addOnSheet, header){
try { 
   var sheetWithData = addOnSheet.getRange('B3').getValue();
   var SheetwithColumn = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetWithData);
   var lastCol = SheetwithColumn.getLastColumn();
   var possibleCols = SheetwithColumn.getRange(1, 1, 1, lastCol).getValues()
   var colValues = possibleCols[0];
   var index = colValues.indexOf(header)+1;
   return index;
//  Logger.log(index);

} catch (e) {
  Logger.log("Error logged (headerIndex): "+e)
  }
  
}


function MakeCopyPerUniqueValue(thisID,thisSheet,uniqueValues,shareWith,permisions,filterColumnNumber,gDriveFolderID,newFileName,sendEmail, colHeader, gDriveNewLocation) {
try {
  var ss = SpreadsheetApp.openById(thisID);
  var sheet = ss.getSheetByName(thisSheet);
  var sheetID = sheet.getSheetId();
  
  var lastRow = sheet.getLastRow(); // with data
  var maxRow = sheet.getMaxRows() // in the sheet
  var lastColumn = sheet.getLastColumn(); // with data
  var maxColumn = sheet.getMaxColumns(); // in the sheet
   
 
 for (i = 0; i < uniqueValues.length; i++){
//  Logger.log("Filtering per UniqueValue = "+uniqueValues[i]);   
   perUniqueValue(thisID, thisSheet, uniqueValues[i], sheetID, filterColumnNumber, gDriveFolderID, newFileName, shareWith[i], permisions[i],sendEmail[i],colHeader, gDriveNewLocation[i] ); // Filtra per Unique Value
  }
 

} catch (e) {
  Logger.log("Error logged (MakeCopyPerUniqueValue): "+e)
}
}



function perUniqueValue (thisID, thisSheet, specificValue, sheetID, filterColumnNumber, gDriveFolderID, newFileName, specificShare, specificPermision, specificEmail, colHeader, specificGDfolder){
//  Logger.log('START Logs fx perUniqueValue below ==>');
//  Logger.log(thisID);
//  Logger.log(thisSheet);
//  Logger.log(specificValue);
//  Logger.log(sheetID);
//  Logger.log(filterColumnNumber);
//  Logger.log(gDriveFolderID);
//  Logger.log(newFileName);
//  Logger.log(specificShare);
//  Logger.log(specificPermision);
//  Logger.log(specificEmail);
//  Logger.log(colHeader);
//  Logger.log(specificGDfolder);
//  Logger.log('END Logs fx perUniqueValue below ==>');
  
try{
  var ss = SpreadsheetApp.openById(thisID);
  var sheet = ss.getSheetByName(thisSheet);
 
 // 1. Creates Spreadsheet Filter and applys specific filter Criteria
  //var filter = ss.getDataRange().createFilter(); 
  var filter = sheet.getDataRange().createFilter(); 
  var filterCriteria = SpreadsheetApp.newFilterCriteria();
  var defineCriteria = filterCriteria.whenTextEqualTo(specificValue);
  var applyfilter = filter.setColumnFilterCriteria(filterColumnNumber, filterCriteria);
  
  
  // 2. Saves de visible data for that unique value into an array
  var returnArray = visibleData(thisID, sheetID, specificValue); // returns array with visibleData to paste
  Logger.log("returnArray.length ="+returnArray.length);
  Logger.log("returnArray[0].length ="+returnArray[0].length);
  
  
  // 3. Creates a new Spreadsheet
  var finalName = (newFileName+"-"+specificValue);
  var finalFolder;
  if (specificGDfolder == ""){
    finalFolder = gDriveFolderID;
  } else {
    finalFolder = specificGDfolder;
  }
  
  var destinationSS = create_New_SS(thisSheet, DriveApp.getFolderById(finalFolder), specificValue, finalName);
  var destinationSSID = destinationSS.getId();
  var destinationSSURL = destinationSS.getUrl();
//  Logger.log("destinationSSURL => "+destinationSSURL);
  
  
  // 4. Pastes Array defined by filter (and uniqueValue) into the new Spreadsheet
  var destSheet = destinationSS.getSheetByName(specificValue);
  destSheet.getRange(1, 1, returnArray.length, returnArray[0].length).setValues(returnArray);
  
  // 5. Campture amount of rows and columns
  var amountofCols = destSheet.getLastColumn();
  var amountofRows = destSheet.getLastRow();
  
  // 6.Shares destSheet with the specificShare and specificPermisions
  var DriveShareDestinationSS = DriveApp.getFileById(destinationSSID) // Share via DriveApp ==> Notifications ON
  var SsShareDestinationSS = SpreadsheetApp.openById(destinationSSID) // Share via SpreadsheetApp ==> Notifications OFF

// Logger.log("specificShare => below");
// Logger.log(specificShare);  // if separated by coma: [jpingib@gmail.com, jp@liceobritanico.com]. Still length 1... not 2...
// Logger.log("specificShare.length => below");
// Logger.log(specificShare.length);  
//  
// Logger.log("specificEmail => below");
// Logger.log(specificEmail);
// Logger.log("specificPermision => below");
// Logger.log(specificPermision);
  
 
 // var sharesArray = specificShare.toString().split(',');  // Split multiple emails within one cell if separated by coma, into an array
 var arraySpecific = specificShare.join().split(', ');
 Logger.log(arraySpecific);
 Logger.log(arraySpecific.length);
 Logger.log(arraySpecific[0]);
 Logger.log(arraySpecific[1]); 
 
    if(specificEmail == 'Yes - Send'){                                   // Share via DriveApp ==> Notifications ON
      if(specificPermision == "" || specificShare == ""){
        // do not share & Change for REPORT
            specificShare = "";
            specificPermision = "";
            specificEmail = "";  
      } else if (specificPermision == 'Edit'){
        DriveShareDestinationSS.addEditors(arraySpecific);  
      } else if (specificPermision == 'View Only'){
        DriveShareDestinationSS.addViewers(arraySpecific)  
      } else if (specificPermision == 'Make Comments'){
        DriveShareDestinationSS.addCommenters(arraySpecific);
      }   
    } else {                                                           // Share via SpreadsheetApp ==> Notifications OFF (Will share Commenters as Viewers)
      if(specificPermision == "" || specificShare == ""){
        // do not share & Change for REPORT
            specificShare = "";
            specificPermision = "";
            specificEmail = "";   
      } else if (specificPermision == 'Edit'){
        SsShareDestinationSS.addEditors(arraySpecific);  
      } else if (specificPermision == 'View Only' || specificPermision == 'Make Comments'){
        SsShareDestinationSS.addViewers(arraySpecific)
        // specificPermission Change for REPORT
              if (specificPermision == 'Make Comments'){
                  specificPermision = 'View Only (Cannot "Make Comments" without sending notification)'
                 }
      }         
    }

  
  // 7. Push results into Report variable (array)
  var today = getTime();   
  report.push([today, destinationSSURL, thisSheet, colHeader, specificValue,amountofRows ,amountofCols ,specificShare, specificPermision, specificEmail, finalName, 'https://drive.google.com/drive/folders/'+finalFolder])
   
    
  // 8. Removes filter so the loop can restart without an error.
  var removeFilter = filter.remove()

} catch (e) {
  Logger.log("Error logged (perUniqueValue): "+e)
  // 7. Removes filter so the loop can restart without an error.
  var removeFilter = filter.remove()
  }
  
}


// Returns the array of data visible when applying the specificValue
function visibleData(ssID, sheetID, specificValue) {
try{
  var spreadsheetId = ssID; 
  var sheetId = sheetID; 
  var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/gviz/tq?tqx=out:csv&gid=" + sheetId + "&access_token=" + ScriptApp.getOAuthToken();
  var res = UrlFetchApp.fetch(url);
  var array = Utilities.parseCsv(res.getContentText());
//  Logger.log("URL Fetched: " + url);
//  Logger.log(specificValue+ " 's array = "+ array);
  return array;
} catch (e) {
  Logger.log("Error logged (visibleData): "+e)
  }  
}


// Creates a new Spreadsheet with specified parameters in a specific Google Drive folder.
function create_New_SS(thisSheet, gdfolder, tabName, finalName){
try{
    var ss = SpreadsheetApp.create(finalName);
    var id = ss.getId();
    var file = DriveApp.getFileById(id);
    gdfolder.addFile(file);
    
    var newSSxSpecificValue = SpreadsheetApp.openById(id);
    newSSxSpecificValue.setActiveSheet(newSSxSpecificValue.getSheets()[0]).setName(tabName).setTabColor('000000')
   
    return newSSxSpecificValue; 
} catch (e) {
  Logger.log("Error logged (create_New_SS): "+e)
  }
}


function getTime(){
 var now = new Date();
 var timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
 var addedTime = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd - HH:mm');
 return addedTime
 }


function pasteAndSortReport(addOnSheet, report) {
try{
  var reportingCol = addOnSheet.getRange('J3:J').getDisplayValues().filter(String);
  var reportingLastRow = reportingCol.length;
//  Logger.log(reportingCol)
//  Logger.log(reportingLastRow)
//  
  addOnSheet.getRange(3+reportingLastRow, 10, report.length, report[0].length).setValues(report);
//  Logger.log('PRINTED');
  
  // Order Descending
  var rangeToOrder = addOnSheet.getRange(3, 10, reportingLastRow + report.length, report[0].length);
  rangeToOrder.sort({column: 10, ascending: false}); // Column from spreadsheet, not from rangeToOrder.
//  Logger.log('& SORTED');

} catch (e) {
  Logger.log("Error logged (pasteAndSortReport): "+e)
  }  
  
}


