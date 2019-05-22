using Microsoft.Office.Interop.Excel;
using System;
using System.IO;

namespace ExcelReader
{//Class
    class ConvertFile
    {
        //Method name
        public string ReturnFile(string file)
        {//Converted file path destination and new converted file name
            string text = @"C:\Users\carlo\Documents\ConvertFile.txt";
            //User input
            Console.WriteLine("Please enter how many spreadsheets to convert ");
            double xlsx = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("ConvertingFile Please wait");


            Application Xlsx = new Application();

            //Open Xlsx File

            Workbook excelBook = Xlsx.Workbooks.Open(file);
            //Loop through spreadsheet quantity
            for (int sheet = 1; sheet < xlsx; sheet++)
            {

                Worksheet excelSheet = excelBook.Sheets[sheet];

                Range excelRange = excelSheet.UsedRange;

                int rows = excelRange.Rows.Count;
                int columns = excelRange.Columns.Count;

                //Loop through rows and columns
                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 1; j <= columns; j++)
                    {

                        //write to path @"C:\Users\carlo\Documents\ConvertedFile.txt
                        if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                        {

                            //Format strings with a comma as delimeter
                            string format = "  {0,12}, \n";
                            //Save all text
                            File.AppendAllText(text, string.Format(format, excelRange.Cells[i, j].Value2.ToString() + "\t"));

                        }
                    }

                }
                //Close Xlsx File
                Xlsx.Quit();
                //User message of file location
                Console.WriteLine("Saving file as ConvertFile.txt To Documents");
                Console.WriteLine("Finish");
                Console.ReadKey();           
                //Return method

            }
            return text;
        }
    }
}