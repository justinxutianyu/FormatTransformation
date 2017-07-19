using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using LumenWorks.Framework.IO.Csv;
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

            string excel_file_path = @"C:\Users\justi\Downloads\categories.xlsx";

            DataTable csvData = GetDataTabletFromCSVFile(csv_file_path);

            //Console.WriteLine("Rows count:" + csvData.Rows.Count);

            ET.ExcelRead.ReadFile(excel_file_path);

            Console.WriteLine("Program End");

            Console.ReadLine();

        }

        public static void ReadCsv()
        {
            // open the file "data.csv" which is a CSV file with headers
            using (CsvReader csv = new CsvReader(
                   new StreamReader("data.csv"), true))
            {
                // missing fields will not throw an exception,
                // but will instead be treated as if there was a null value
                csv.DefaultParseErrorAction = ParseErrorAction.RaiseEvent;

                csv.ParseError += new ParseErrorEventHandler(csv_ParseError);
                int fieldCount = csv.FieldCount;

                string[] headers = csv.GetFieldHeaders();
                while (csv.ReadNextRecord())
                {
                    for (int i = 0; i < fieldCount; i++)
                        Console.Write(string.Format("{0} = {1};",
                                      headers[i], csv[i]));
                    Console.WriteLine();
                }
            }
        }
        public static void csv_ParseError(object sender, ParseErrorEventArgs e)
        {
            // if the error is that a field is missing, then skip to next line
            if (e.Error is MissingFieldCsvException)
            {
                Console.Write("--MISSING FIELD ERROR OCCURRED");
                e.Action = ParseErrorAction.AdvanceToNextLine;
            }
        }

        private static DataTable GetDataTabletFromCSVFile(string csv_file_path)

        {

            DataTable csvData = new DataTable();

            try

            {

                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))

                {
                    csvReader.ReadLine();

                    csvReader.SetDelimiters(new string[] { ";" });

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

            }

            return csvData;

        }

    }

}
