using System;


namespace ExcelReader
{
    class Program
    {

        static void Main(string[] args)
        {//Path from file location
            string fileLcation = @"C:\Users\carlo\Documents\";
            //Class instance
            ConvertFile newConvFile = new ConvertFile();
            //User Information message
            Console.WriteLine("Please enter xlsx workbook to convert or spreadsheet to convert  from your Documents ");
            //Variable and user input
            string inputfile = Console.ReadLine();
            //Class instance method calling 
            newConvFile.ReturnFile(fileLcation + inputfile);



        }
    }
}