using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using FastMember;
using ClosedXML.Excel;
using System.Dynamic;
using System.ComponentModel;

namespace HerramientasNicolas.App_Code
{
    public class ExcelFile
    {
        //attributes
        private DataSet sheets;
        private string path;
        private string name;

        //properties
        public DataSet Sheets
        {
            get { return sheets; }
            set { sheets = value; }
        }

        public string Path
        {
            get { return path; }
            set { path = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        //constructors
        public ExcelFile() { }

        public ExcelFile(DataSet _sheets, string _path, string _name)
        {
            Sheets = _sheets;
            Path = _path;
            Name = _name;
        }

        //functions
        public static DataTable ConvertListToDataTable(List<dynamic> listObject, string tableName)
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
                table.TableName = tableName;

                return table;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static DataSet ConvertListToDataSet(List<List<dynamic>> listObjects)
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

        public void CreateExcel()
        {
            try
            {
                //check this excel
                CheckExcelObject();

                //delete if exists
                if (File.Exists(this.Path + this.Name))
                    File.Delete(this.Path + this.Name);

                using (XLWorkbook wb = new XLWorkbook())
                {
                    //create excel
                    wb.Worksheets.Add(this.Sheets);
                    wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    wb.Style.Font.Bold = true;
                    wb.SaveAs((this.Path + this.Name));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void CheckExcelObject()
        {
            try
            {
                if (String.IsNullOrEmpty(this.Path.Trim()))
                    throw new Exception("The path can't be empty.");

                if (!Directory.Exists(this.Path))
                    throw new Exception("The path directory is incorrect. '" + this.Path + "'");

                if (String.IsNullOrEmpty(this.Name.Trim()))
                    throw new Exception("The name of the file can't be empty.");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}