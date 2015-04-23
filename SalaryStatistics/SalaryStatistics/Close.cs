using OfficeOpenXml;
using System;
using System.IO;
namespace SalaryStatistics
{
    public partial class Data
    {
        public void Close() {
            string newFilePath = Path.GetDirectoryName(filePath) + "Processed " + Path.GetFileName(filePath);
            excelFile.SaveAs(new FileStream(newFilePath, FileMode.Create));

            Console.WriteLine("File Saved and Closed.");
            Environment.Exit(1);
        }
    }
}