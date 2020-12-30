# Parsing MS Excel file with Google Apps Script

Parsing MS Excel files and returns values in JSON format.

## 👀 Overview

So far there is no native Google Apps Script method to get data straight from MS Excel files stored on Google Drive or in Gmail attachment.  
The widespread workaround what I found is to first convert the MS Excel workbook to Google Spreadsheet and than use Apps Script `SpreadsheetApp` functions to get data.  
The function `parseMSExcelBlob(blob, requiredSheets)` provided in this repo can open and parse directly MS Excel files without any upload or conversion to Google Spreadsheet.

## 💻 Built with

* [Google Apps Script](https://developers.google.com/apps-script)

## ✍🏼 Description

### Method

MS Excel workbooks are zipped collections of XML files (except binary files). Using GAS function `Utilities.unzip(blob)` it can be unzipped and knowing the XML structure of the unzipped files you can extract the data needed.  
Then the seemingly straightforward solution to parse and access data in XML files would be to use Apps Script native function `XmlService.parse(xml)` but it turned out to be quite slow. However getting the unzipped XML files text content with the function `getDataAsString()` data can be retrieved based on specific XML patters. The function `parseMSExcelBlob(blob, requiredSheets)` uses this approach and has much faster processing time than using `XmlService.parse(xml)`.

### Usage

The function `parseMSExcelBlob(blob, requiredSheets)` is parsing MS Excel files and returns values in JSON format.
* First parameter is the MS Excel `blob`.
* Second parameter is an array of required sheet names so you can restrict the parsing process for specific worksheets saving some time and resource. If parameter `requiredSheets` is omitted all worksheets will be parsed in the workbook.

### Limitation

* For the moment `parseMSExcelBlob(blob, requiredSheets)` can process only XML based excel files (not .xlsb).
* Maximum blob size allowed in `Utilities.unzip(blob)` is 50MB.

## ⚙️ Examples

### Sample script #1

Getting data from a MS Excel file saved in Google Drive from a worksheet called "Orders" using spreadsheet bound script.

```javascript
function getDataFromDrive(){
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
}
```

### Sample script #2

Getting data from a MS Excel file attachment of an email with subject "MyDailyReport" using spreadsheet bound script.

```javascript
function getDataFromGmail(){
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
}
```

## 📝 License

This project is licensed under [MIT](https://opensource.org/licenses/MIT) license.

## ✉️ Contact

Csaba Csonka - cscsonka@gmail.com

Project Link: https://github.com/cscsonka/Parsing-MS-Excel-file-with-Google-Apps-Script

## 🤓 Contribute

Contributions are always welcome. Create a PR to show your idea!

## ⭐️ Support

Give a star if this project helped you!  

<a href="https://www.buymeacoffee.com/cscsonka" target="_blank">
  <img src="https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png" alt="Buy Me A Coffee">
</a>


