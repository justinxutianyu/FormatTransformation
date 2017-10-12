using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.VisualBasic.FileIO;
using Format;
using System.Reflection;

namespace ExcelTransfomation
{


    class ExcelRead
    {
        /// <summary>
        /// this function read Excel xlsx file from path.
        /// </summary>
        /// <param name="excel_file_path"></param>
        public static void ReadFile(string excel_file_path) {

            var book = new LinqToExcel.ExcelQueryFactory(excel_file_path);

            var query =
                from row in book.Worksheet("Category")
                let item = new
                {
                    Id = row["Id"].Cast<string>(),
                    Name = row["Name"].Cast<string>(),
                    SeName = row["SeName"].Cast<string>(),
                }
                //where item.id == "28"
                select item;

            foreach (var item in query) {

                string categoryinfo = "id:{0}, name:{1}, sename:{2}";
                Console.WriteLine(string.Format(categoryinfo,item.Id,item.Name,item.SeName));


            }
        }

        public static DataTable GetDataTabletFromCSVFile(string csv_file_path)

        {
            Console.WriteLine(csv_file_path);

            DataTable csvData = new DataTable();

            try

            {

                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path, Encoding.GetEncoding("iso-8859-1")))

                { 
                    csvReader.ReadLine();

                    csvReader.SetDelimiters(new string[] { "," });

                    csvReader.HasFieldsEnclosedInQuotes = true;

                    string[] colFields = csvReader.ReadFields();

                    foreach (string column in colFields)

                    {
                        Console.WriteLine(column);

                        DataColumn datecolumn = new DataColumn(column);

                        datecolumn.AllowDBNull = true;

                        csvData.Columns.Add(datecolumn);

                    }

                    while (!csvReader.EndOfData)

                    {

                        string[] fieldData = csvReader.ReadFields();

                        //Making empty value as null

                        for (int i = 0; i < fieldData.Length; i++)

                        {

                            if (fieldData[i] == "")

                            {

                                fieldData[i] = null;

                            }

                        }

                        csvData.Rows.Add(fieldData);

                    }

                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return csvData;

        }

    }

    class ExcelWrite {

        /// <summary>
        /// this fucntion help create and edit existing Excel xlsx file.
        /// </summary>
        /// <param name="excel_file_path"></param>

        public static void Transform(string excel_file_path, string destination_file_path, 
                string[] attr, string sheet_name)
        {
            DataTable dt = new DataTable();

            var excel_mf = new LinqToExcel.ExcelQueryFactory(@"F:\Project\manufacturers.xlsx");

            var pricebook_manufacturer =
                from row in excel_mf.Worksheet("Manufacturer")
                let manufacturer = new
                {
                    id = row["Id"].Cast<string>(),
                    name = row["Name"].Cast<string>(),
                    //Name = row["Name"].Cast<string>(),
                    sename = row["SeName"].Cast<string>(),
                }
                //where item.id == "28"
                select manufacturer;
            
            Dictionary<string, Format.Bib_Manufacture> d = new Dictionary<string, Bib_Manufacture>();

            foreach ( var manufacturer in pricebook_manufacturer) {

                d.Add(manufacturer.name, new Bib_Manufacture(manufacturer.id,manufacturer.name));
            }

            DataTable csvData = ExcelRead.GetDataTabletFromCSVFile(excel_file_path);


            Console.WriteLine("Rows count:" + csvData.Rows.Count);
            /*
            var excel = new LinqToExcel.ExcelQueryFactory(@"‪F:\Project\PRODUCTS_Test.xlsx");
           
            var pricebook_product =
                from row in excel.Worksheet("PRODUCTS_Test")
                let product = new
                {
                    name = row["!MANUFACTURER"].Cast<string>(),
                    //Name = row["Name"].Cast<string>(),
                    //SeName = row["SeName"].Cast<string>(),
                }
                //where item.id == "28"
                select product;
           */
            DataColumn dc = csvData.Columns[sheet_name];

            foreach (DataRow r in csvData.Rows)
            {
                string item = r[dc].ToString();
                if (!d.ContainsKey(item)) {

                    d.Add(item, new Bib_Manufacture(Convert.ToString(d.Count+2), item));
                    string categoryinfo = "Add new name:{0}";
                    Console.WriteLine(string.Format(categoryinfo, item));

                }

            }

            foreach (string s in attr) {

                dt.Columns.Add(new DataColumn(s, typeof(string)));

            }
            // write data into Datatable

            foreach (var item in d) {

                DataRow dr = dt.NewRow();
                dr["Id"] = item.Value.Id;
                dr["Name"] = item.Key;
                dt.Rows.Add(dr);
                
            }

            using (FileStream stream = new FileStream(destination_file_path, FileMode.Create, FileAccess.Write))
            {
                IWorkbook wb = new XSSFWorkbook();
                ISheet sheet = wb.CreateSheet(sheet_name);
                ICreationHelper cH = wb.GetCreationHelper();

                IRow header = sheet.CreateRow(0);
                for (int j = 0; j < dt.Columns.Count; j++) {
                    ICell cell = header.CreateCell(j);
                    cell.SetCellValue(cH.CreateRichTextString(dt.Columns[j].ColumnName.ToString()));
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    IRow row = sheet.CreateRow(i+1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row.CreateCell(j);
                        Console.WriteLine(cH.CreateRichTextString(dt.Rows[i].ItemArray[j].ToString()));
                        cell.SetCellValue(cH.CreateRichTextString(dt.Rows[i].ItemArray[j].ToString()));
                    }
                }
                wb.Write(stream);
            }
        }

        /// <summary>
        /// from Bibendum csv file to Pricebook Category
        /// this fucntion help create and edit existing Excel xlsx file.
        /// </summary>
        /// <param name="excel_file_path"></param>
        public static void TransformCategory(DataTable csvData, string excel_file_path,
                                        string destination_file_path, string sheet_name)
        {
            // data table store final table written to destination file
            DataTable dt = new DataTable();

            var excel_mf = new LinqToExcel.ExcelQueryFactory(excel_file_path);

            //category
            excel_mf.AddMapping<Priceboo_Category>(x => x.Id, "Id");
            excel_mf.AddMapping<Priceboo_Category>(x => x.Name, "Name");
            excel_mf.AddMapping<Priceboo_Category>(x => x.Description, "Description");
            excel_mf.AddMapping<Priceboo_Category>(x => x.CategoryTemplateId, "CategoryTemplateId");
            excel_mf.AddMapping<Priceboo_Category>(x => x.MetaKeywords, "MetaKeywords");
            excel_mf.AddMapping<Priceboo_Category>(x => x.MetaDescription, "MetaDescription");
            excel_mf.AddMapping<Priceboo_Category>(x => x.MetaTitle, "MetaTitle");
            excel_mf.AddMapping<Priceboo_Category>(x => x.SeName, "SeName");
            excel_mf.AddMapping<Priceboo_Category>(x => x.ParentCategoryId, "ParentCategoryId");
            excel_mf.AddMapping<Priceboo_Category>(x => x.Picture, "Picture");
            excel_mf.AddMapping<Priceboo_Category>(x => x.PageSizeOptions, "PageSizeOptions");
            excel_mf.AddMapping<Priceboo_Category>(x => x.PriceRanges, "PriceRanges");
            excel_mf.AddMapping<Priceboo_Category>(x => x.ShowOnHomePage, "ShowOnHomePage");
            excel_mf.AddMapping<Priceboo_Category>(x => x.DisplayOrder, "DisplayOrder");
            excel_mf.AddMapping<Priceboo_Category>(x => x.IncludeInTopMenu, "IncludeInTopMenu");
            excel_mf.AddMapping<Priceboo_Category>(x => x.Published, "Published");
            excel_mf.AddMapping<Priceboo_Category>(x => x.DisplayOrder, "DisplayOrder");
            // read origin data in pricebook manufacture table
            var pricebook_category =
                from category in excel_mf.Worksheet<Priceboo_Category>(sheet_name)
                    //where item.id == "28"
                select category;

            // create dictionary in case overlapping add records
            Dictionary<string, Priceboo_Category> d = new Dictionary<string, Priceboo_Category>();

            foreach (var category in pricebook_category)
            {
                if (category.Name.Contains('/'))
                {
                    string[] templates = category.Name.Split('/');
                    category.Name = templates[0];
                    if (d.ContainsKey(templates[1]))
                    {
                        category.ParentCategoryId = d[templates[1]].Id;
                    }
                    else {
                        Priceboo_Category temp = new Priceboo_Category();
                        temp.Initializer();
                        temp.Id = Convert.ToString(d.Count + 2);
                        temp.Name = templates[1];
                        temp.ParentCategoryId = "2"; //default grandparent category is wine
                        d.Add(templates[1], temp);
                    }

                }
                if (!d.ContainsKey(category.Name))
                {
                    d.Add(category.Name, category);
                }

            }

            // get records
            DataColumn dc = csvData.Columns["!CATEGORY"];

            DataColumn region = csvData.Columns["!REGION"];

            DataColumn key = csvData.Columns["!PRODUCTCODE"];

            foreach (DataRow r in csvData.Rows)
            {
                // judge whether product exists 
                if (r[key] != null && r[dc] != null )
                {
                    if (!r[key].Equals("") && !r[dc].Equals("")) {

                        string parent_category = r[dc].ToString();
                        string category = r[region].ToString();
                        if (parent_category.Contains('/'))
                        {
                            string[] category_array = parent_category.Split('/');
                            // judge whether the manufacture exists 
                            if (!d.ContainsKey(category_array[0]))
                            {
                                Priceboo_Category temp = new Priceboo_Category();
                                temp.Initializer();
                                temp.Id = Convert.ToString(d.Count + 2);
                                temp.Name = category_array[0];
                                temp.ParentCategoryId = "2"; //default grandparent category is wine
                                d.Add(category_array[0], temp);
                                string info = "Add new main category:{0}";
                                Console.WriteLine(string.Format(info, category_array[0]));

                            }
                            if (!d.ContainsKey(category))
                            {
                                Priceboo_Category temp = new Priceboo_Category();
                                temp.Initializer();
                                temp.Id = Convert.ToString(d.Count + 2);
                                temp.Name = category;
                                temp.ParentCategoryId = d[category_array[0]].Id; //default grandparent category is wine
                                d.Add(category, temp);
                                string info = "Add new sub category:{0}";
                                Console.WriteLine(string.Format(info, category));

                            }
                        }
                        else
                        {
                            if (!d.ContainsKey(parent_category))
                            {
                                Priceboo_Category temp = new Priceboo_Category();
                                temp.Initializer();
                                temp.Id = Convert.ToString(d.Count + 2);
                                temp.Name = parent_category;
                                temp.ParentCategoryId = "2"; //default grandparent category is wine
                                d.Add(parent_category, temp);
                                string info = "Add new main category:{0}";
                                Console.WriteLine(string.Format(info, parent_category));

                            }
                            if (!d.ContainsKey(category))
                            {
                                Priceboo_Category temp = new Priceboo_Category();
                                temp.Initializer();
                                temp.Id = Convert.ToString(d.Count + 2);
                                temp.Name = category;
                                temp.ParentCategoryId = d[parent_category].Id; //default grandparent category is wine
                                d.Add(category, temp);
                                string info = "Add new sub category:{0}";
                                Console.WriteLine(string.Format(info, category));

                            }

                        }

                    }

                }

            }

            // add records to final table
            foreach (string column in excel_mf.GetColumnNames(sheet_name))
            {

                dt.Columns.Add(new DataColumn(column, typeof(string)));

            }

            // write data into Datatable
            foreach (var item in d)
            {
                // create a new row
                DataRow dr = dt.NewRow();

                // use reflection mechanism to call corresponding function 
                foreach (string column in excel_mf.GetColumnNames(sheet_name))
                {
                    MethodInfo m = typeof(Priceboo_Category).GetMethod(column);

                    Type type = item.Value.GetType();
                    PropertyInfo property = type.GetProperty(column);
                    Console.WriteLine(property.GetValue(item.Value));
                    dr[column] = property.GetValue(item.Value);
                }

                dt.Rows.Add(dr);

            }

            using (FileStream stream = new FileStream(destination_file_path, FileMode.Create, FileAccess.Write))
            {
                IWorkbook wb = new XSSFWorkbook();
                ISheet sheet = wb.CreateSheet(sheet_name);
                ICreationHelper cH = wb.GetCreationHelper();

                IRow header = sheet.CreateRow(0);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = header.CreateCell(j);
                    cell.SetCellValue(cH.CreateRichTextString(dt.Columns[j].ColumnName.ToString()));
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    IRow row = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row.CreateCell(j);
                        Console.WriteLine(cH.CreateRichTextString(dt.Rows[i].ItemArray[j].ToString()));
                        cell.SetCellValue(cH.CreateRichTextString(dt.Rows[i].ItemArray[j].ToString()));
                    }
                }
                
                wb.Write(stream);
            }

        }

        /// <summary>
        /// from Bibendum csv file to Pricebook Manufacturers
        /// this fucntion help create and edit existing Excel xlsx file.
        /// </summary>
        /// <param name="excel_file_path"></param>
        public static void TransformManufacture(DataTable csvData, string excel_file_path,
                                        string destination_file_path, string sheet_name) {
            // data table store final table written to destination file
            DataTable dt = new DataTable();

            var excel_mf = new LinqToExcel.ExcelQueryFactory(excel_file_path);

            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.Id, "Id");
            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.Name, "Name");
            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.ManufacturerTemplateId, "ManufacturerTemplateId");
            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.MetaKeywords, "MetaKeywords");
            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.MetaDescription, "MetaDescription");
            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.MetaTitle, "MetaTitle");
            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.SeName, "SeName");
            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.Picture, "Picture");
            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.PageSize, "PageSize");
            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.AllowCustomersToSelectPageSize, "AllowCustomersToSelectPageSize");
            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.PageSizeOptions, "PageSizeOptions");
            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.PriceRanges, "PriceRanges");
            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.Published, "Published");
            excel_mf.AddMapping<Format.Priceboo_Manufacture>(x => x.DisplayOrder, "DisplayOrder");

            // read origin data in pricebook manufacture table
            var pricebook_manufacturer =
                from item in excel_mf.Worksheet<Priceboo_Manufacture>(sheet_name)
                //where item.id == "28"
                select item;

            // create dictionary in case overlapping add records
            Dictionary<string, Priceboo_Manufacture> d = new Dictionary<string, Priceboo_Manufacture>();

            foreach (var item in pricebook_manufacturer)
            {

                d.Add(item.Name, item);
            }

            // get 
            DataColumn dc = csvData.Columns["!MANUFACTURER"];

            foreach (DataRow r in csvData.Rows)
            {
                // get manufacture name from Bibendum product table "!MANUFACTURE" table
                string name = r[dc].ToString();

                // judge whether the manufacture exists 
                if (!d.ContainsKey(name))
                {
                    Priceboo_Manufacture temp = new Priceboo_Manufacture();
                    temp.Initializer();
                    temp.Id = Convert.ToString(d.Count + 2);
                    temp.Name = name;
                    d.Add(name, temp);
                    string info = "Add new manufacture:{0}";
                    Console.WriteLine(string.Format(info, name));

                }

            }

            // add records to final table
            foreach (string column in excel_mf.GetColumnNames(sheet_name))
            {

                dt.Columns.Add(new DataColumn(column, typeof(string)));

            }

            // write data into Datatable
            foreach (var item in d)
            {
                // create a new row
                DataRow dr = dt.NewRow();

                // use reflection mechanism to call corresponding function 
                foreach (string column in excel_mf.GetColumnNames(sheet_name))
                {
                    MethodInfo m = typeof(Priceboo_Manufacture).GetMethod(column);

                    Type type = item.Value.GetType();
                    PropertyInfo property = type.GetProperty(column);
                    Console.WriteLine(property.GetValue(item.Value));
                    dr[column] = property.GetValue(item.Value);
                    //dt.Columns.Add(new DataColumn(column, typeof(string)));
                }

                dt.Rows.Add(dr);

            }

            using (FileStream stream = new FileStream(destination_file_path, FileMode.Create, FileAccess.Write))
            {
                IWorkbook wb = new XSSFWorkbook();
                ISheet sheet = wb.CreateSheet(sheet_name);
                ICreationHelper cH = wb.GetCreationHelper();

                IRow header = sheet.CreateRow(0);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = header.CreateCell(j);
                    cell.SetCellValue(cH.CreateRichTextString(dt.Columns[j].ColumnName.ToString()));
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    IRow row = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row.CreateCell(j);
                        Console.WriteLine(cH.CreateRichTextString(dt.Rows[i].ItemArray[j].ToString()));
                        cell.SetCellValue(cH.CreateRichTextString(dt.Rows[i].ItemArray[j].ToString()));
                    }
                }
                wb.Write(stream);
            }

        }

        /// <summary>
        /// from Bibendum csv file to Pricebook Products
        /// this fucntion help create and edit existing Excel xlsx file.
        /// </summary>
        /// <param name="excel_file_path"></param>
        public static void TransformProduct(DataTable csvData, string excel_file_path,
                 string destination_file_path, string sheet_name, Dictionary<string, string> map) {

            // data table store final table written to destination file
            DataTable dt = new DataTable();

            var excel_mf = new LinqToExcel.ExcelQueryFactory(excel_file_path);

            //category
            //excel_mf.AddMapping<Priceboo_Product>();
            //excel_mf.AddMapping<Priceboo_Product>(x => x.ProductType, "ProductType");
            //excel_mf.AddMapping<Priceboo_Product>(x => x.ParentGroupedProductId, "ParentGroupedProductId");
            //excel_mf.AddMapping<Priceboo_Product>(x => x.VisibleIndividually, "Description");
            //excel_mf.AddMapping<Priceboo_Product>(x => x.Name, "Name");

            // read origin data in pricebook manufacture table
            var pricebook_category =
                from product in excel_mf.Worksheet<Priceboo_Product>(sheet_name)
                    //where item.id == "28"
                select product;

            // create dictionary in case overlapping add records
            Dictionary<string, Priceboo_Product> d = new Dictionary<string, Priceboo_Product>();

            foreach (var product in pricebook_category)
            {

                if (!d.ContainsKey(product.SKU))
                {
                    d.Add(product.SKU, product);
                }

            }

            // get records
            DataColumn key = csvData.Columns["!PRODUCTCODE"];

            Dictionary<string, DataColumn> attr = new Dictionary<string, DataColumn>();

            foreach (var element in map) {

                if (!attr.ContainsKey(element.Key))
                {
                    attr.Add(element.Key, csvData.Columns[element.Value]);

                }

            }

            foreach (DataRow r in csvData.Rows)
            {
                // judge whether product exists 
                if (r[key] != null && !r[key].Equals(""))
                {
                    string SKU = "BIB" + r[key];
                    // attr format like <"ProductType", Column[!CATEGORY]> import relevant attribute to d <SKU, Proceboo_Product>
                    if (d.ContainsKey(SKU)) 
                    {
                        // this product exists in the dictionary
                        foreach (var element in attr)
                        {
                            Priceboo_Product product = d[SKU];
                            Type type = product.GetType();
                            PropertyInfo prop = type.GetProperty(element.Key);
                            //prop.SetValue(product, r[element.Value], null);
                            if (r[element.Value] == DBNull.Value)
                            {
                                prop.SetValue(product, string.Empty, null);
                            }
                            else
                            {
                                prop.SetValue(product, r[element.Value], null);
                            }
                        }


                    }
                    else {
                        // this product does not exist in dictionary
                        Priceboo_Product p = new Priceboo_Product();
                        p.Initializer();
                        d.Add(SKU, p);
                        foreach (var element in attr)
                        {
                            Priceboo_Product product = d[SKU];
                            Type type = product.GetType();
                            PropertyInfo prop = type.GetProperty(element.Key);
                            Console.WriteLine(SKU+",,,"+r[element.Value]);
                            if (r[element.Value] == DBNull.Value)
                            {
                                prop.SetValue(product, string.Empty, null);
                            }
                            else
                            {
                                prop.SetValue(product, r[element.Value], null);
                            }
                            

                        }

                    }

                }

            }

            // add records to final table
            foreach (string column in excel_mf.GetColumnNames(sheet_name))
            {

                dt.Columns.Add(new DataColumn(column, typeof(string)));

            }

            // write data into Datatable
            foreach (var item in d)
            {
                // create a new row
                DataRow dr = dt.NewRow();

                // use reflection mechanism to call corresponding function 
                foreach (string column in excel_mf.GetColumnNames(sheet_name))
                {
                    MethodInfo m = typeof(Priceboo_Category).GetMethod(column);

                    Type type = item.Value.GetType();
                    PropertyInfo property = type.GetProperty(column);
                    Console.WriteLine(property.GetValue(item.Value));
                    dr[column] = property.GetValue(item.Value);
                }

                dt.Rows.Add(dr);
                

            }

            using (FileStream stream = new FileStream(destination_file_path, FileMode.Create, FileAccess.Write))
            {
                IWorkbook wb = new XSSFWorkbook();
                ISheet sheet = wb.CreateSheet(sheet_name);
                ICreationHelper cH = wb.GetCreationHelper();

                IRow header = sheet.CreateRow(0);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = header.CreateCell(j);
                    cell.SetCellValue(cH.CreateRichTextString(dt.Columns[j].ColumnName.ToString()));
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    IRow row = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row.CreateCell(j);
                        Console.WriteLine(cH.CreateRichTextString(dt.Rows[i].ItemArray[j].ToString()));
                        cell.SetCellValue(cH.CreateRichTextString(dt.Rows[i].ItemArray[j].ToString()));
                    }
                }

                wb.Write(stream);
            }

        }


    }


}
