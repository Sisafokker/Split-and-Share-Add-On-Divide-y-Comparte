/*  

OAuth verification video requested: https://www.youtube.com/watch?v=LtrDt6Ig-L4


3 OAuth Scopes functionality descriptions:
https://www.googleapis.com/auth/drive
https://www.googleapis.com/auth/script.external_request
https://www.googleapis.com/auth/spreadsheets

ScriptApp: just for the Addon menu based on authMode
UrlFetchApp: Copies filtered data from the table the user wants to copy into a new spreadsheet.
SpreadsheetApp: 
  - Reads data in the spreadsheet
  - Gets values from specific cells
  - Creates data validation and filters
  - Reads sheet names in the spreadsheet
  - Shares files without sending notifications
  - Creates Activity report for the user
  
DriveApp:
  - Creates new spreadsheets
  - adds/places new files in the gDrive folders the user specifices
  - Sends notifications to those users whom the new files have been shared with

*/


/*
 * This add-on will create new Spreadsheets for groups of data that match 
 * specific Filter Criteria (unique values in specific columns & rows).
 * User specifies which sheet contains the dataset
 * User specifies which column dictates the unique values
 * User specifies in which gDrive folder the new files will be created
 * User specifies whom to share these new files with (and the sharing permissions)
 * User specifies where new files are stored (specific gDrive folders)



Introduction 
Split a table into multiple new files (based on criteria in a column) and share those with specific users. Go to "Add-ons" > "Split and Share" and run de Initial Setup

Description
"Split and Share" helps you sort and distribute specific information in your sheets (based on criteria in a column) to specific users or shared folders.

Easily separate data from a database / table / Form Response sheet into individual spreadsheets, 
and shares them (optional) with users and/or place the new files in specific folders. The Addon does NOT COLLECT any user data.


Post Install tip:
Run the initial setup, set your parameters and run the addon. 
*/
