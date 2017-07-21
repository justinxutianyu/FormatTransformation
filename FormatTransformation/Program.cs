using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Microsoft.VisualBasic.FileIO;
using ET = ExcelTransfomation;

/// <summary>
/// This file is to transform csv file from Bibendum to corresponding format
/// of Pricebook.
/// </summary>

namespace ReadDataFromCSVFile
{
    static class Program
    {
        static void Main()
        {

            string csv_file_path = @"F:\Project\PRODUCTS.csv";

            string excel_file_path = @"F:\Project\PRODUCTS.csv";

            string dest_file_path = @"C:\Users\justi\Downloads\test.xlsx";

            string sheet_name = @"!MANUFACTURER";

            string[] attr = { "Id","Name", "Description",   "CategoryTemplateId",
                "MetaKeywords", "MetaDescription",  "MetaTitle", "SeName", "ParentCategoryId",
                "Picture", "PageSize",  "AllowCustomersToSelectPageSize",   "PageSizeOptions",
                "PriceRanges", "ShowOnHomePage", "IncludeInTopMenu", "Published", "DisplayOrder"};

            DataTable csvData = ET.ExcelRead.GetDataTabletFromCSVFile(csv_file_path);

            //test code set
            Console.WriteLine("我是谁，我在哪儿");

            Console.WriteLine("Rows count:" + csvData.Rows.Count);

            //ET.ExcelRead.ReadFile(excel_file_path);

            ET.ExcelWrite.Transform(excel_file_path,dest_file_path, attr, sheet_name);

            Console.WriteLine("Program End");

            Console.ReadLine();

        }

    }

}
