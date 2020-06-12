# MSController

A small C# library for interacting with Outlook and Excel using the Microsoft Interop libraries. 

**Dependencies**  
Microsoft.Office.Interop.Excel version 15 (Available via NuGet and [online](https://www.microsoft.com/en-us/download/details.aspx?id=3508))  
Microsoft.Office.Interop.Outlook version 15 (Available via NuGet and [online](https://www.microsoft.com/en-us/download/details.aspx?id=3508))  

**Getting Started**  
Download the DLL and XML from [releases](https://github.com/DStewart1997/MSController/releases), put the two files in your \bin\debug\ folder (Along with Microsoft.Office.Interop.Excel.dll and Microsoft.Office.Interop.Outlook.dll), add references to the DLLs and you're good to go.


## Quick examples - ExcelHandler

```C#
ExcelHandler excelHandler = new ExcelHandler();
excelHandler.open(FILEPATH);
string data = excelHandler.getCell("A", 1);  // Gets value from cell A1
string dataLast = excelHandler.getLastCellInColumn("A");  // Gets the value from the last occupied row in column A
excelHandler.writeCell("A", 2, data);  // Writes data to cell A2
excelHandler.close();
```


## Quick examples - OutlookHandler

```C#
OutlookHandler outlookHandler = new OutlookHandler();
// recipient and attachmentPath can either be strings or List<string>s - attachmentPath is optional
outlookHandler.sendMail(subject, body, recipient, attachmentPath);  
```
    
-------------------------------------------
    
#### Future changes
ExcelHandler
- Implement a writeLastCellInRow method.
- Implement deleteColumn and deleteRow methods.
- Implement isOpen method.
- Improve exception handling - with custom exceptions.

OutlookHandler
- Improve exception handling.
