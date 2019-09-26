using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
namespace TestCodeLibrary
    {
    class ExcelTOCsv {
        private static string time = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff");
        private static string outputFileNamePath;
        private static string inputFileNamePath;
        public static void excelToCsv(string excelFileToConvert,string saveAsCsvfileOutputPath,string logFilePath)
            {
            inputFileNamePath = excelFileToConvert;
            outputFileNamePath = saveAsCsvfileOutputPath;
            try
                {
                Excel.Application application = new Excel.Application();
                application.DisplayAlerts = false;
                // Open 
                Excel.Workbook excelWorkbook = application.Workbooks.Open(inputFileNamePath);
                // Save file.
                excelWorkbook.SaveAs(outputFileNamePath, Excel.XlFileFormat.xlCSV);
                // Close .
                excelWorkbook.Close();
                // Quit 
                application.Quit();
                }
            catch (Exception ex)
                {//Write to log file
                File.AppendAllText(logFilePath, "\n*******************************" + time + "\n" + ex.TargetSite + "\n " + ex.Message + "\n" + ex.Data + "\n" + ex.StackTrace + "\n*******************************");

                }
            }
    }
}
