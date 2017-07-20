using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Diagnostics;

namespace Test
{
    class Test
    {
        public static void RenderToExcel()
        {
            /*   FileStream fs = null;


               if (!File.Exists("test.xlsx"))
               {
                   FileStream fs = new FileStream(@"D:\test.xlsx", FileMode.Create, FileAccess.Write);
               }
               else
               { */
            FileStream fs = File.Open(@"D:\test.xlsx", FileMode.Open, FileAccess.ReadWrite);

            // using (FileStream fs = new FileStream(@"D:\test.xlsx", FileMode.Create, FileAccess.Write))
            //  {


            IWorkbook workbook = new XSSFWorkbook(fs);
            {
                XSSFSheet sheet = workbook.GetSheetAt(0) as XSSFSheet;
                {
                    XSSFRow headerRow = sheet.GetRow(0) as XSSFRow;


                    XSSFCell cell = headerRow.GetCell(1) as XSSFCell;
                    Console.WriteLine("cell={0}", cell);
                    cell.SetCellValue("lanfang");
                    fs.Close();
                }
                using (var wook = new FileStream(@"D:\test.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(wook);
                }
            }
        }
        private static void WriteExcelFile01(string filename, string[] ExlKey, string[] ExlValue)
        {
            Console.WriteLine("************WriteExcelFile{0} Content***************", filename);
            var ext = Path.GetExtension(filename).ToLower();
            //   using (FileStream fs = File.OpenWrite(filename))  //open myxls.xls file     
            using (FileStream fs = File.Open(filename, FileMode.Open, FileAccess.Read))
            {

                IWorkbook wk01;
                if (ext.Contains(".xlsx"))
                    wk01 = new XSSFWorkbook(fs);
                else
                    wk01 = new HSSFWorkbook(fs);
                fs.Close();
                ISheet sheet01 = wk01.GetSheetAt(0);
                Console.WriteLine("Table Rows", sheet01.LastRowNum);
                for (int j = 0; j <= sheet01.LastRowNum; j++) //当前表的总行数，j=1 避开第一行
                {
                    IRow row = sheet01.GetRow(j); //读取当前行数据
                    if (row != null)
                    {
                        ICell cell01 = row.GetCell(0);//当前行第1个单元格内容==key
                        ICell cell02 = row.GetCell(1);//当前行第2个单元格内容==value                        
                        if (cell02 == null)
                        {
                            cell02 = row.CreateCell(1);
                            cell02.SetCellValue(string.Empty);
                        }
                        if (cell01 != null && cell01.ToString().StartsWith("var"))
                        {
                            if (ExlKey.Contains(cell01.ToString())) //如果文件中的key在key数组里
                            {
                                Console.WriteLine("cell01[{0}]={1},cell02={2}", j, cell01.ToString(), cell02.ToString());
                                Console.WriteLine("ExlValue[{0}]={1}", j, ExlValue[j]);
                                int x = 0;//在ExlValue数组里，元素从0开始
                                x = j - 1;
                                //   ReplaceExcel(filename, ExlValue[x], x);
                                cell01.SetCellValue(ExlValue[x]);
                            }
                        }
                    }
                }
                try
                {
                    FileStream fileStream = File.Open(filename, FileMode.Open, FileAccess.Write);
                    Console.WriteLine("保存的文件是：{0}", filename);
                    wk01.Write(fileStream);

                    fileStream.Close();
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
            }

        }

        private static void ReplaceExcel(string tempPath, string CellValue, int CellNum)
        {
            Console.WriteLine("************替换{0}里的内容***************", tempPath);
            IWorkbook wk = null;
            var ext1 = Path.GetExtension(tempPath).ToLower();
            using (FileStream fs = File.Open(tempPath, FileMode.Open,
            FileAccess.Read, FileShare.ReadWrite))
            {
                //把xls文件读入workbook变量里，之后就可以关闭了  
                if (ext1.Contains(".xlsx"))
                    wk = new XSSFWorkbook(fs);
                else
                    wk = new HSSFWorkbook(fs);
                fs.Close();
            }
            //在第二行创建行    
            ISheet sheet = wk.GetSheetAt(0);
            //   CellNum += 1;
            IRow row01 = sheet.GetRow(CellNum);
            if (row01 != null)
            {
                ICell cell01 = row01.GetCell(0);
                if (cell01 != null)
                {
                    cell01.SetCellValue(CellValue);
                }
            }
            // using (FileStream fileStream = File.Open(tempPath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            try
            {
                FileStream fileStream = File.OpenWrite(tempPath);
                wk.Write(fileStream);
                fileStream.Close();
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }

    }
}
