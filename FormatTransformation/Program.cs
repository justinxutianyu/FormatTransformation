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

            string excel_file_path_cat = @"F:\Project\categories.xlsx";

            string excel_file_path_man = @"F:\Project\manufacturers.xlsx";

            string dest_file_path_cat = @"C:\Users\justi\Downloads\category.xlsx";

            string dest_file_path_man = @"C:\Users\justi\Downloads\manufacture.xlsx";

            string sheet_name_cat = @"Category";

            string sheet_name_man = @"Manufacturer";

            Dictionary<string, string> BibToPB_Product = new Dictionary<string, string>();

            BibToPB_Product.Add("ProductType", "!CATEGORY"); //category = producttype + parentgroupid
            BibToPB_Product.Add("VisibleIndividually", "!FORSALE"); // Y = TRUE, H = FALSE
            BibToPB_Product.Add("Name", "!PRODUCT"); 
            BibToPB_Product.Add("ShortDescription", "!DESCR");
            BibToPB_Product.Add("FullDescription", "!FULLDESCR");
            BibToPB_Product.Add("SeName", "!CLEAN_URL"); 
            BibToPB_Product.Add("SKU", "!PRODUCTCODE"); // BIB + product code 
            BibToPB_Product.Add("ManufacturerPartNumber", "!MANUFACTURERID"); // BIB + MANUFACTURERID
            BibToPB_Product.Add("OrderMinimumQuantity", "!MIN_AMOUNT"); 
            BibToPB_Product.Add("StockQuantity", "!AVAIL"); 
            BibToPB_Product.Add("DisableBuyButton", "!FORSALE"); // Y =TRUE H = FALSE
            BibToPB_Product.Add("Price", "!LIST_PRICE");
            BibToPB_Product.Add("Categories", "!REGION");
            BibToPB_Product.Add("Manufacturers", "!MANUFACTURER");
            BibToPB_Product.Add("NotifyAdminForQuantityBelow", "!LOW_AVAIL_LIMIT"); 


            //string[] attr = { "Id","Name", "Description",   "CategoryTemplateId",
            //    "MetaKeywords", "MetaDescription",  "MetaTitle", "SeName", "ParentCategoryId",
            //    "Picture", "PageSize",  "AllowCustomersToSelectPageSize",   "PageSizeOptions",
            //    "PriceRanges", "ShowOnHomePage", "IncludeInTopMenu", "Published", "DisplayOrder"};

            DataTable csvData = ET.ExcelRead.GetDataTabletFromCSVFile(csv_file_path);

            //test code set

            Console.WriteLine("Rows count:" + csvData.Rows.Count);

            //ET.ExcelRead.ReadFile(excel_file_path);

            string testString = "Lusco Do Miño";

            byte[] bytes = Encoding.Default.GetBytes(testString);

            testString = Encoding.GetEncoding("iso-8859-1").GetString(bytes);

            Console.WriteLine(testString);

            //ET.ExcelWrite.TransformManufacture(csvData, excel_file_path_man,dest_file_path_man,sheet_name_man);

            ET.ExcelWrite.TransformCategory(csvData, excel_file_path_cat, dest_file_path_cat, sheet_name_cat);

            Console.WriteLine("Program End");

            Console.ReadLine();

        }

    }

}
