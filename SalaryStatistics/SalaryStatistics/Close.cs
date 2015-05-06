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

           // fixTheFormatting();

           // excelFile.Stream.Close();
            MessageBox.Show("File Saved and Closed to " + newFilePath);
            //file.writeline("File saved to " + newFilePath);
            //file.writeline("Operation ran successfully on: " + DateTime.Now);

           // file.Close();

        }

        private void fixTheFormatting()
        {
            foreach (string worksheetName in processedWorksheetNames)
            {
                ExcelWorksheet currentWorksheet = excelFile.Workbook.Worksheets[worksheetName];

                currentWorksheet.Cells["D2:G2"].Style.Numberformat.Format = "$###,###,##0";
                currentWorksheet.Cells["A:Z"].AutoFitColumns();
            }
        }
    }
}