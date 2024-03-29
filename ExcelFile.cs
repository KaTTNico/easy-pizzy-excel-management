using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using ClosedXML.Excel;
using System.ComponentModel;
using System.Text.RegularExpressions;

/// <summary>Easy pizzy excel files management.</summary>
/// 
/// <author>Nicolas Aguirre</author>
namespace HerramientasNicolas.App_Code
{
    public class ExcelFile : XLWorkbook
    {
        //properties
        public string Path { get; set; }
        public string FileName { get; set; }

        //constructors
        /// <summary>
        /// Creates an empty instance of ExcelFile.
        /// </summary>
        public ExcelFile() : base() { }

        /// <summary>
        /// Creates an instance of ExcelFile using a DataSet as the data source for the workbook.
        /// </summary>
        /// <param name="sheets"></param>
        /// <param name="path"></param>
        /// <param name="fileName"></param>
        public ExcelFile(DataSet sheets, string path, string fileName)
        {
            this.Worksheets.Add(sheets);
            Path = path;
            FileName = fileName;
        }

        /// <summary>
        /// Creates an instance of ExcelFile using a single DataTable as the data source for the workbook.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="path"></param>
        /// <param name="fileName"></param>
        public ExcelFile(DataTable sheet, string path, string fileName)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(sheet);

            this.Worksheets.Add(ds);
            Path = path;
            FileName = fileName;
        }

        /// <summary>
        /// Creates an instance of ExcelFile using a List<dynamic> as the data source for the workbook.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="path"></param>
        /// <param name="fileName"></param>
        public ExcelFile(List<dynamic> sheet, string path, string fileName)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(ConvertListToDataTable(sheet, ""));

            this.Worksheets.Add(ds);
            Path = path;
            FileName = fileName;
        }

        /// <summary>
        /// Creates an instance of ExcelFile using a List<List<dynamic> as the data source for the workbook.
        /// </summary>
        /// <param name="sheets"></param>
        /// <param name="path"></param>
        /// <param name="fileName"></param>
        public ExcelFile(List<List<dynamic>> sheets, string path, string fileName)
        {
            this.Worksheets.Add(ConvertListToDataSet(sheets));
            Path = path;
            FileName = fileName;
        }

        //constructors to read excel from a file
        /// <summary>
        /// Creates an instance of ExcelFile using a path to read an excel from a file.
        /// </summary>
        /// <param name="path"></param>
        public ExcelFile(string path) : base(path)
        {
            FileName = Regex.Match(path, @"\w+\.\w+$").Value;
            Path = path.Replace(Regex.Match(path, @"(\\|\/|\\{2}|\/{2})\w+\.\w+$").Value, "");
        }

        //functions
        /// <summary>
        /// Creates an Excel file in the indicated path with the indicated name.
        /// </summary>
        public void CreateExcel()
        {
            try
            {
                //check this excel
                CheckExcelObject();

                //delete if exists
                if (File.Exists(this.Path + this.FileName))
                    File.Delete(this.Path + this.FileName);

                this.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                this.Style.Font.Bold = true;

                this.SaveAs((this.Path + this.FileName));

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Returns the worksheet at the indicated index in the workbook as a DataTable.
        /// </summary>
        /// <param name="worksheetIndex"></param>
        /// <returns></returns>
        public DataTable GetWorksheetAsDataTable(int worksheetIndex)
        {
            if (this.Worksheets.Worksheet(worksheetIndex) == null)
                return null;

            return GetDataTableFromWorksheet(this.Worksheets.Worksheet(worksheetIndex));
        }

        /// <summary>
        /// Returns the workbook as a DataSet.
        /// </summary>
        /// <returns></returns>
        public DataSet GetWorkbookAsDataSet()
        {
            if (this.Worksheets.Count == 0)
                return null;

            DataSet dsWorksheets = new DataSet();
            //loop through the worksheets
            foreach (IXLWorksheet worksheet in this.Worksheets)
                //add datatable
                dsWorksheets.Tables.Add(GetDataTableFromWorksheet(worksheet));
            
            return dsWorksheets;
        }

        //internal functions 
        protected DataTable GetDataTableFromWorksheet(IXLWorksheet worksheet)
        {
            //Create a new DataTable.
            DataTable dtWorksheet = new DataTable();

            //Read Sheet from Excel file.
            //Loop through the Worksheet rows.
            bool firstRow = true;
            foreach (IXLRow row in worksheet.Rows())
            {
                //Use the first row to add columns to DataTable.
                if (firstRow)
                {
                    foreach (IXLCell cell in row.Cells())
                        dtWorksheet.Columns.Add(cell.Value.ToString());
                    
                    firstRow = false;
                }
                else
                {
                    //Add rows to DataTable.
                    dtWorksheet.Rows.Add();
                    int i = 0;

                    foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                    {
                        dtWorksheet.Rows[dtWorksheet.Rows.Count - 1][i] = cell.Value.ToString();
                        i++;
                    }
                }
            }
            return dtWorksheet;
        }

        protected static DataTable ConvertListToDataTable(List<dynamic> listObject, string tableName)
        {
            try
            {
                //get properties of the first item on the list
                PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(listObject[0]);

                //create datatable
                DataTable table = new DataTable();

                //add columns
                foreach (PropertyDescriptor prop in properties)
                    table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);

                //add rows
                foreach (dynamic item in listObject)
                {
                    DataRow row = table.NewRow();

                    foreach (PropertyDescriptor prop in properties)
                        row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;

                    table.Rows.Add(row);
                }
                table.TableName = String.IsNullOrEmpty(tableName) ? "sheet1" : tableName;

                return table;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        protected static DataSet ConvertListToDataSet(List<List<dynamic>> listObjects)
        {
            try
            {
                //create dataset
                DataSet tables = new DataSet();

                //convert each list object into datatable
                int counter = 0;
                foreach (List<dynamic> listObject in listObjects)
                    tables.Tables.Add(ConvertListToDataTable(listObject, "Table " + (counter++)));

                return tables;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        protected void CheckExcelObject()
        {
            try
            {
                if (String.IsNullOrEmpty(this.Path.Trim())) throw new Exception("The path can't be empty.");

                if (!Directory.Exists(this.Path)) throw new Exception("The path directory is incorrect. '" + this.Path + "'");

                if (String.IsNullOrEmpty(this.FileName.Trim())) throw new Exception("The name of the file can't be empty.");

                if (this.Worksheets.Count == 0) throw new Exception("You must add at least one sheet on the sheets collection.");

                if (!String.IsNullOrEmpty(Regex.Match(this.FileName, @".\w+$").Value)) this.FileName = (this.FileName.Replace(Regex.Match(this.FileName, @".\w+$").Value, ".xlsx"));
                else this.FileName = this.FileName += ".xlsx";
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
