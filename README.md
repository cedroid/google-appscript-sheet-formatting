A Google Apps Script for Automated Formatting in Google Sheets. Designed for sheets with inconsistent column order and formatting, especially useful for large spreadsheets with multiple tabs, such as those containing hotel reservation data. It includes functions to sort columns alphabetically, apply predefined formatting rules, and log processing status.

The success of the entire script, especially the formatting, hinges on the correct setup of the getColumnFormats() function. It's essential that the column names and their corresponding formats in this function accurately reflect the actual data structure in your sheets for the script to work effectively.
 
This script is based on a Google Sheet that contains hotel reservations data but the columns are not sorted and the cell formatting was not correct, wheb the issue happend across hundreds of tabs within the same sheet.
 
1. processSheets()
This is the main function that initiates the processing of all sheets in the active spreadsheet. It performs several key tasks:
 
- Retrieves all the sheets in the spreadsheet.
- Creates or accesses the 'Formatting Logs' sheet for tracking progress.
- Processes each sheet (up to 30 at a time for resource management).
- Logs the processing status (success or error) for each sheet.
- Displays a notification summarizing the number of sheets successfully processed and those with errors.
 
2. processSingleSheet(sheet, columnFormats)
Processes an individual sheet by:
 
- Checking if the sheet has any columns and rows.
- Sorting the columns alphabetically based on headers.
- Writing the sorted data back to the sheet.
- Applying formatting to each column based on predefined rules.
- Returning the status of processing (success or error with notes).
 
3. applyFormat(range, formatInfo)
- Applies formatting to a given range in a sheet. It uses the format information specified in the formatInfo parameter. The function switches between different types of data (date, text, number, currency) and applies the corresponding format.
 
4. logStatus(logSheet, sheet, status)
- Logs the processing status of each sheet in the 'Formatting Logs' sheet. It records the sheet name, a hyperlink to the sheet, the timestamp of processing, and notes on the status (success or error).
 
5. setupLogSheet(logSheet)
- Sets up the initial headers in the 'Formatting Logs' sheet for tracking. Headers include "Sheet Name", "Hyperlink", "Timestamp", "Status", and "Status Notes".
 
6. isSheetProcessed(logSheet, sheetName)
- Checks if a sheet has already been processed successfully. It prevents reprocessing of sheets that are already formatted.
 
7. getColumnFormats()
- Crucial for the entire script, this function returns a hardcoded object mapping column names to their respective data types and formats. It is essential that the user manually matches these names and formats with the actual column names present in the sheets. Failure to correctly match these will result in improper or failed formatting. The formats for each type of data (like date, text, number, currency) need to be predefined here.
