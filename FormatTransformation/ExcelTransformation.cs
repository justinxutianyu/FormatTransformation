using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

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

            var excel = new LinqToExcel.ExcelQueryFactory(excel_file_path);

            var query =
                from row in excel.Worksheet("Products(2)")
                let item = new
                {
                    name = row["!MANUFACTURER"].Cast<string>(),
                    //Name = row["Name"].Cast<string>(),
                    //SeName = row["SeName"].Cast<string>(),
                }
                //where item.id == "28"
                select item;

            var excel_mf = new LinqToExcel.ExcelQueryFactory(@"C:\Users\justi\Downloads\manufacturers");

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

            foreach (var item in query)
            {
                if (!d.ContainsKey(item.name)) {

                    d.Add(item.name, new Manufacturer(Convert.ToString(d.Count+2), item.name, null));
                    string categoryinfo = "Add new name:{0}";
                    Console.WriteLine(string.Format(categoryinfo, item.name));

                }

            }

            foreach (string s in attr) {

                dt.Columns.Add(new DataColumn(s, typeof(string)));

            }
            /*
            using (FileStream stream = new FileStream(excel_file_path, FileMode.Open, FileAccess.Read))
            {
                IWorkbook wb = new XSSFWorkbook(stream);
                ISheet sheet = wb.GetSheet(sheet_name);
                string holder;
                int i = 0;
                do
                {
                    DataRow dr = dt.NewRow();
                    IRow row = sheet.GetRow(i);
                    try
                    {
                        holder = row.GetCell(0, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString();
                    }
                    catch (Exception)
                    {
                        break;
                    }

                    string city = holder.Substring(0, holder.IndexOf(','));
                    string state = holder.Substring(holder.IndexOf(',') + 2, 2);
                    string zip = holder.Substring(holder.IndexOf(',') + 5, 5);
                    dr[0] = city;
                    dr[1] = state;
                    dr[2] = zip;
                    dt.Rows.Add(dr);
                    i++;
                } while (!String.IsNullOrEmpty(holder));
            }
            */
            // write data into Datatable

            foreach (var item in d) {

                DataRow dr = dt.NewRow();
                dr["Id"] = item.Value.Id;
                dr["Name"] = item.Key;
                dr["SeName"] = item.Value.Sename;
                


            }

            using (FileStream stream = new FileStream(destination_file_path, FileMode.Create, FileAccess.Write))
            {
                IWorkbook wb = new XSSFWorkbook();
                ISheet sheet = wb.CreateSheet(sheet_name);
                ICreationHelper cH = wb.GetCreationHelper();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    IRow row = sheet.CreateRow(i);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row.CreateCell(j);
                        cell.SetCellValue(cH.CreateRichTextString(dt.Rows[i].ItemArray[j].ToString()));
                    }
                }
                wb.Write(stream);
            }
        }


    }
}
