using System;


namespace ExcelReader
{
    class Program
    {

        static void Main(string[] args)
        {   //Files
            string exceFile = @"C:\Users\Desktop\test-2019-09-24-09-56-17-9834.xlsx";
            string outputCsvFile = @"C:\Users\Desktop\New folder\Output\converted.csv";
            string log = @"C:\Users\carlo\Desktop\New folder\log.txt";
            //Method
      
            ExcelTOCsv.excelToCsv(exceFile,outputCsvFile,log);



        }
    }
}
