using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
namespace SalaryStatistics
{
    public partial class Data
    {
        public void addStatistics()
        {
            char[] valueString;
            ExcelWorksheet currentWorksheet;
            string[] statisticHeaders = { "1st Quartile", "Mean", "3rd Quartile", "Median", "Associate Prof. Compression" };
            string currentJobTitle;
            List<string> jobTitles = new List<string>();
            int bottomOfSortedRows = 2;
            int numberOfThatJob;

            foreach (string worksheetName in processedWorksheetNames)
            {
                currentWorksheet = excelFile.Workbook.Worksheets[worksheetName];
                jobTitles = new List<string>();
                //Insert two blank Rows
                currentWorksheet.InsertRow(2, 2);
                //Insert the statistic column headers
                for (int r = 0; r < 5; r++)
                {
                    currentWorksheet.Cells[1, r + 4].Value = statisticHeaders[r];
                }
                bottomOfSortedRows = 4;
                //Insert the worksheet wide statistics
                currentWorksheet.Cells[2, 1].Value = "All";
                currentWorksheet.Cells[2, 4].Formula = "=QUARTILE(C:C,1)";
                currentWorksheet.Cells[2, 5].Formula = "=AVERAGE(C:C)";
                currentWorksheet.Cells[2, 6].Formula = "=QUARTILE(C:C,3)";
                currentWorksheet.Cells[2, 7].Formula = "=MEDIAN(C:C,1)";
                currentWorksheet.Cells[2, 8].Formula = "=G2/'Average New Asst Prof Salary'!D2";

                currentWorksheet.Cells["D2:G2"].Style.Numberformat.Format = "$###,###,##0";
                currentWorksheet.Cells["H2"].Style.Numberformat.Format = "#0.00";
                }
            }

        }

        struct cell
        {
            public int x;
            public int y;
            public string value;
            public cell(int _y, int _x, string _value)
            {
                x = _x;
                y = _y;
                value = _value;
            }
        }
 }
