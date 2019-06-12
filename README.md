# easy-pizzy-excel-management
An alternative to manage the creation an easy read excel files

This class require ClosedXML.Excel library to be able to function.

Constructors:
(1) ExcelFile():
This is the classic empty one.
Creates an empty instance of ExcelFile.

(2) ExcelFile(DataSet sheets, string path, string fileName):
Creates an instance of ExcelFile using a DataSet as the data source for the workbook.
You can use this one when you are working with multiple datatables and you want to create an excel file from them.

(3) ExcelFile(DataTable sheet, string path, string fileName):
Creates an instance of ExcelFile using a single DataTable as the data source for the workbook.
You can use this one when you are working with a single datatable and want to create an excel file from that.

(4) ExcelFile(List<dynamic> sheet, string path, string fileName):
Creates an instance of ExcelFile using a List<dynamic> as the data source for the workbook.
You can use this one when you recieve a JSON from another lenguage for example python or javascript and want to create an excel file
from it.
  
(5) ExcelFile(List<List<dynamic>> sheets, string path, string fileName):
Creates an instance of ExcelFile using a List<List<dynamic> as the data source for the workbook.
You can use this one when you recieve or when you are working with multiple JSON and want to create an excel from them.
  
(6) ExcelFile(string path) : base(path):
Creates an instance of ExcelFile using a path to read an excel from a file.
You can use this one when you have an excel file and want to convert it to a datatable or a dataset.

Functions:
(1) void CreateExcel():
Creates an Excel file in the indicated path with the indicated name.

(2) DataTable GetWorksheetAsDataTable(int worksheetIndex):
Returns the worksheet at the indicated index in the workbook as a DataTable.
When you create an instance of ExcelFile using the constructor number 6 it means you are reading an existing excel file.
With this instance you can call this method to get the index indicated worksheet converted to a datatable.

(3) DataSet GetWorkbookAsDataSet():
Returns the workbook as a DataSet.
When you create an instance of ExcelFile using the constructor number 6 it means you are reading an existing excel file.
With this instance you can call this method to get the entire workbook converted to a dataset with each worksheet as a dataset.

.--. ..- - --- . .-.. --.- ..- . .-.. --- .-.. . .-
