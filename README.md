# Parsing MS Excel file with Google Apps Script

Parsing MS Excel files and returns values in JSON format.

## Description

### Method

MS Excel workbooks are zipped collections of XML files.
The method `Utilities.unzip(blob)` can be used to unzip MS Excel files.
Getting the unzipped XML files as string data can be extracted based on specific patters. This approach results in much faster processing time than using `XmlService.parse(xml)`.

### Usage

The function `parseMSExcelBlob(blob, requiredSheets)` is for parsing MS Excel file and returns data in 2D arrays.
* First parameter is the MS Excel `blob`.
* Second parameter is an array of required sheet names so you can restrict the parsing process for specific worksheets saving some time and resource.
If parameter `requiredSheets` is omitted all worksheets will be parsed in the workbook.

## Examples

### Sample script #1

Getting data from a MS Excel file saved in Google Drive from a worksheet called "Orders".

```javascript
var ss = SpreadsheetApp.getActiveSpreadsheet();

// getting a MS Excel file in Google Drive
var file = DriveApp.getFileById("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx");
var blob = file.getBlob();

// if second parameter is not provided all sheets will be parsed
var data = parseMSExcelBlob(blob, ["Orders"]);

// test if everything is good
if( data["Error"] ) return ss.toast(data["Error"], "Something went wrong 🙄", 10);

// here we have the data in 2D array
var tbl = data["Orders"];

// do your stuff
// ...
```

### Sample script #2

Getting data from a MS Excel file attachment of an email with subject "MyDailyReport".

```javascript
var ss = SpreadsheetApp.getActiveSpreadsheet();

// getting a MS Excel attachment from email
var subject = "MyDailyReport";
var threads = GmailApp.search('subject:"' + subject + '"');
var messages = GmailApp.getMessagesForThreads(threads);
var msg_id = messages[0][0].getId();
var file = GmailApp.getMessageById(msg_id).getAttachments()[0];
var blob = file.copyBlob();

// if second parameter is not provided all sheets will be parsed
var data = parseMSExcelBlob(blob);

// test if everything is good
if( data["Error"] ) return ss.toast(data["Error"], "Something went wrong 🙄", 10);

// here we have the data in 2D array
var firstSheet = Object.keys(data)[0];
var tbl = data[firstSheet];

// do your stuff
// ...
```


