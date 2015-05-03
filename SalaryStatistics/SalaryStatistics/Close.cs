using OfficeOpenXml;
using System;
using System.IO;
using System.Windows.Forms;
namespace SalaryStatistics
{
    public partial class Data
    {
        public void Close() {
            string newFilePath = Path.GetDirectoryName(filePath) + "\\" + "Processed " + Path.GetFileName(filePath);
            excelFile.SaveAs(new FileStream(newFilePath, FileMode.Create));

            MessageBox.Show("File Saved and Closed to " + newFilePath);
            file.WriteLine("Operation ran successfully on: " + DateTime.Now);

            file.Close();

            Environment.Exit(1);
        }
    }
}