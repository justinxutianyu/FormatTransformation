using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTransfomation
{
    class ExcelRead
    {
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
}
