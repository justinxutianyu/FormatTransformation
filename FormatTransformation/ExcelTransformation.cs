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

namespace ExcelTransfomation
{

    class Manufacturer {

        public Manufacturer(string id, string name, string sename) {

            this.id = id;
            this.name = name;
            this.sename = sename;

        }

        private string id;

        public string Id { get => id; set => id = value; }

        private string name;

        public string Name { get => name; set => name = value; }

        private string sename;

        public string Sename { get => sename; set => sename = value; }



    }

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

                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))

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
            
            Dictionary<string, Manufacturer> d = new Dictionary<string, Manufacturer>();

            foreach ( var manufacturer in pricebook_manufacturer) {

                d.Add(manufacturer.name, new Manufacturer(manufacturer.id,manufacturer.name,manufacturer.sename));
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
            DataColumn dc = csvData.Columns["!MANUFACTURER"];

            foreach (DataRow r in csvData.Rows)
            {
                string item = r[dc].ToString();
                if (!d.ContainsKey(item)) {

                    d.Add(item, new Manufacturer(Convert.ToString(d.Count+2), item, null));
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
                dr["SeName"] = item.Value.Sename;
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


    }


}
