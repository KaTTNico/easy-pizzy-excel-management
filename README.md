# easy-pizzy-excel-management
An alternative to manage the creation of excel files

This class require ClosedXML.Excel library to be able to function.

In this class you can easily manage the creation of excel files because of it's functions.

Functions:
*ConvertListToDataTable(List<dynamic> listObject, string tableName)
So this function converts a list of dynamic objects into a datatable with it's column names as the attributes names of the dynamic objects of the list.

For example:
Let's say i have a string of data 'colectedData' wich each line of this string represents a product for example: '|1||||PRD103XXLBLUE|||' wich '1' represents the cuantity of the product and 'PRD103XXLBLUE' represents the code of the product.

  List<dynamic> products = new List<dynamic>();

            using (StringReader reader = new StringReader(colectedData))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    //crear producto
                    dynamic product = new ExpandoObject();
                    product.cuantity = int.Parse(Regex.Match(line, @"(?<=^\d{1}\|{2}\.\|\.\|)(.*)(?=\|{4})").Value);
                    product.code = Regex.Match(line, @"(?<=\|{4})(.*)(?=\|{1})").Value.Trim().ToUpper();
                    products.Add(product);
                }
            }
