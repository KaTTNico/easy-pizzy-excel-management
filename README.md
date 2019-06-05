# easy-pizzy-excel-management
An alternative to manage the creation an easy read excel files

This class require ClosedXML.Excel library to be able to function.

Constructors:

(1) ExcelFile()
This is the classic empty one.
Creates an empty instance of ExcelFile.

(2) ExcelFile(DataSet sheets, string path, string fileName)
Creates an instance of ExcelFile using a DataSet as the data source for the workbook.
You can use this one when you are working with multiple datatables and you want to create an excel file from them.

(3) ExcelFile(DataTable sheet, string path, string fileName)
Creates an instance of ExcelFile using a single DataTable as the data source for the workbook.
You can use this one when you are working with a single datatable and want to create an excel file from that.

(4) ExcelFile(List<dynamic> sheet, string path, string fileName)
Creates an instance of ExcelFile using a List<dynamic> as the data source for the workbook.
You can use this one when you recieve a JSON from another lenguage for example python or javascript and want to create an excel file
from it.
  
(5) ExcelFile(List<List<dynamic>> sheets, string path, string fileName)
Creates an instance of ExcelFile using a List<List<dynamic> as the data source for the workbook.
You can use this one when you recieve or when you are working with multiple JSON and want to create an excel from them.
  
(6) ExcelFile(string path) : base(path)
Creates an instance of ExcelFile using a path to read an excel from a file.
You can use this one when you have an excel file and want to convert it to a datatable or a dataset.
