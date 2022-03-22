////////////////////////////////////////////////////////////////////////////////////
////////// SIMPLE TRIGGERS /////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////

function onInstall(e){
  onOpen(e);
}


function onOpen(e) {
//  Logger.log(e.authMode);
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    // Add a normal menu item (works in all authorization modes).
     menu.addItem('Initial Setup', 'setUpInicial')
} else {
    // Add a menu item based on properties (doesn't work in AuthMode.NONE).
     menu.addItem('Initial Setup', 'setUpInicial')
     menu.addSeparator()
     menu.addItem('Run Add-on', 'runAddOn')
  }
  menu.addToUi();
  
}

function onEdit(e){
//   Logger.log(e);
//   Logger.log(e.authMode);
   if (e && e.authMode == ScriptApp.AuthMode.LIMITED || e && e.authMode == ScriptApp.AuthMode.FULL){
try {
//    Logger.log("RUN ONEDIT(e) => Add-on INSTALLED and ENABLED => e.authMode is LIMITED or FULL");
    sheetList = sheetnms();
    var addOnSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SetupSheet);
    var currentSelectedSheet = addOnSheet.getRange('B3').getValue();
    var currentSelectedHeader = addOnSheet.getRange('B4').getValue();
    
    var range = e.range
    var ss = e.source;
    var sheetName = ss.getActiveSheet().getName();
    var column = range.getColumn();
    var row = range.getRow();
    var input = e.value;
    var oldValue = e.oldValue;
    var HeaderOfInput = ss.getActiveSheet().getRange(1,column).getValue();
    
    
    // IF Selected Sheet not found among sheets
    if (sheetList.indexOf(currentSelectedSheet) == -1 && currentSelectedSheet !== ""){
//     Logger.log("currentSelectedSheet was not found. Clear B3 and update Validation!")
     
     var cellB3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SetupSheet).getRange('B3');
     cellB3.clearContent()
     cellB3.setDataValidation(SpreadsheetApp.newDataValidation()
                                                .setAllowInvalid(true)
                                                .requireValueInList(sheetList)
                                                .build()
                                                );
     clearSelectedSheetandHeader(addOnSheet);
     clearUniqueValues(addOnSheet) 
    }      

    
    // IF they select the SetupSheet as source of data
    if (sheetName == SetupSheet && column === 2 && row === 3 && input == SetupSheet){
//      Logger.log("Marcador 1")
      displayToast('You cannot select this sheet', SetupSheet);
      clearSelectedSheetandHeader(addOnSheet);
      clearUniqueValues(addOnSheet)
    }
    
    // If the Selected Sheet changes (B3) or if the header in the Selected Sheet changes, then =>
    if (sheetName == SetupSheet && column === 2 && row === 3 && input !== "" && input !== SetupSheet){
//      Logger.log("Marcador 2")
      displayToast('Updating ...', SetupSheet);
      clearSelectedHeader(addOnSheet);
      clearUniqueValues(addOnSheet);
      createHeaderValidation(input);  
    } 
    
    // If any headers in the already selected sheet change
    if (sheetName == currentSelectedSheet && row == 1){
//      Logger.log("Marcador 3")
      createHeaderValidation(currentSelectedSheet);  
    } 
    
    // When cell B4 Changes: Brings unique values for the header column B4 specifies.
    if (sheetName == SetupSheet && column == 2 && row == 4 && input !== "") {
//      Logger.log("Marcador 4")
      displayToast('Updating "Unique Values"', SetupSheet);
      getUniqueColValues (addOnSheet, input)
    }
    
    // If the header that changed is the one that was selected
    if (oldValue === currentSelectedHeader){
//      Logger.log("Marcador 5")
      addOnSheet.getRange('B4').setValue(input)
    }
    
    //If the column with displayed unique values change.. reset unique values
    if (HeaderOfInput == currentSelectedHeader && row > 1 && sheetName !== SetupSheet ){
//      Logger.log("Marcador 6")
      getUniqueColValues (addOnSheet, currentSelectedHeader)
    } 
    
} catch (e) {
  Logger.log("Error logged (onEdit): "+e);
}
    
  } else {
//    Logger.log("DO NOT RUN ONEDIT(e) => e.authMode is NONE (not LIMITED, not FULL)");
    }   
}



/////////////////////////////////////////////////////////////////////////////////////////////////////////
////// Support onEdit actions ///////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////////


// Transforms an array in row format to an array in Col format.
function row2Col(row) {
  return row.map(function(elem) {return [elem];});
  // return row[0].map(function(elem) {return [elem];});
}

// Maps all the sheets in the active spreadsheet
function sheetnms() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(
    function(x) {return x.getName();}
  );
}

function clearUniqueValues(addOnSheet){
   addOnSheet.getRange('D3:D').clearContent();
}

function clearSelectedHeader(addOnSheet){
   addOnSheet.getRange('B4').clearContent();
}

function clearSelectedSheetandHeader(addOnSheet){
   addOnSheet.getRange('B3:B4').clearContent();
}

function displayToast(message, sheet){
  SpreadsheetApp.getActiveSpreadsheet().toast(sheet +' sheet', message, 1);
}

// Generates a Data validation for the different column that selected sheet currently has.
function createHeaderValidation(input){
//   Logger.log("Marcador fx1");
try {
   var selectedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(input);
   var lastCol = selectedSheet.getLastColumn();
   var possibleCols = selectedSheet.getRange(1, 1, 1, lastCol).getValues()
   var colValues = possibleCols[0];
    
//    Logger.log(possibleCols); // Result: [[Nombre, Responsable, ORG, Grades, Attendance, , , , , , ]]
//    Logger.log(colValues); // Result: [Nombre, Responsable, ORG, Grades, Attendance, , , , , , ]
    
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SetupSheet).getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation()
                                                .setAllowInvalid(true)
                                                .requireValueInList(colValues)
                                                .build()
                                                );
} catch (e) {
  Logger.log("Error logged (createHeaderValidation): "+e);
  }  
}



function getUniqueColValues (addOnSheet, input){
   Logger.log("Marcador fx2");
try {
   clearUniqueValues (addOnSheet);
   var sheetWithData = addOnSheet.getRange('B3').getValue();
   var SheetwithColumn = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetWithData);
   var lastCol2 = SheetwithColumn.getLastColumn();
   var lastRow2 = SheetwithColumn.getLastRow();
   var possibleCols2 = SheetwithColumn.getRange(1, 1, 1, lastCol2).getValues()
   // var colValues = possibleCols[0].filter(String);
   var colValues = possibleCols2[0];
   var index2 = colValues.indexOf(input)+1;
  
//   Logger.log("index2 Below");
//   Logger.log(index2);
//   Logger.log("----------");
    
   var uniqueSet = SheetwithColumn.getRange(2, index2, lastRow2).getValues(); 
   var uniqueSet0 = uniqueSet[0]
   
//   Logger.log("uniqueSet Below");
//   Logger.log(uniqueSet); // Returns [[1991.0], [1992.0], [1993.0], [1991.0], [1992.0], [1993.0], [1991.0], [1992.0], [1993.0], []] Formato Columna. Perfecto.
//   Logger.log("----------");
    
   // var uniqueSet = SheetwithColumn.getRange(2, index, lastRow, index+1).getValues();
   var newArray = [];
   
   
// Gets unique values but transforms that array into a ROW array
   uniqueSet.forEach(function(x){ 
    if(newArray.indexOf(x[0]) === -1){
        Logger.log(x[0])
        newArray.push(x[0]);
        }                   
       }); 
    
//    Logger.log("NewArray below");
//    Logger.log(newArray); // [1991.0, 1992.0, 1993.0, ]
//    Logger.log("----------");
 
 // Transforms Row array into a Column array.
    var newArrayInCol = row2Col(newArray);
//    Logger.log("newArrayInCol below");
//    Logger.log(newArrayInCol); // [[1991.0], [1992.0], [1993.0], []]
//    Logger.log("----------");
    

    addOnSheet.getRange(3, 4, newArrayInCol.length).setValues(newArrayInCol); 
} catch (e) {
  Logger.log("Error logged (getUniqueColValues): "+e);
  }  
}


