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
            int endRow;

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

                //Insert the worksheet wide statistics
                currentWorksheet.Cells[2, 1].Value = "All";
                currentWorksheet.Cells[2, 4].Formula = "=QUARTILE(C:C,1)";
                currentWorksheet.Cells[2, 5].Formula = "=AVERAGE(C:C)";
                currentWorksheet.Cells[2, 6].Formula = "=QUARTILE(C:C,3)";
                currentWorksheet.Cells[2, 7].Formula = "=MEDIAN(C:C,1)";
                currentWorksheet.Cells[2, 8].Formula = "=G2/'Average New Asst Prof Salary'!D2";
                currentWorksheet.Row(2).Style.Font.Bold = true;
                currentWorksheet.Row(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                endRow = currentWorksheet.Dimension.End.Row;
                for (int r = 3; r < endRow; r++)
                {
                    if (currentWorksheet.Name[0] == 'H' || currentWorksheet.Name[0] == 'h')
                    {
                        if (currentWorksheet.Cells[r, 1].Value != currentWorksheet.Cells[r + 1, 1].Value
                            && currentWorksheet.Cells[r, 1].Value != null && currentWorksheet.Cells[r+1, 1].Value != null)
                        {
                            currentWorksheet.InsertRow(r, 2, 2); //Insert two rows, one to be blank the other as a header.
                            currentWorksheet.Cells[r + 2, 1].Value = currentWorksheet.Cells[r + 3, 1].Value; //Set the Header Name

                            //Insert the Formulae
                            currentWorksheet.Cells[r + 2, 4].Formula = "=QUARTILE(C:C,1)";
                            currentWorksheet.Cells[r + 2, 5].Formula = "=AVERAGE(C:C)";
                            currentWorksheet.Cells[r + 2, 6].Formula = "=QUARTILE(C:C,3)";
                            currentWorksheet.Cells[r + 2, 7].Formula = "=MEDIAN(C:C,1)";
                            currentWorksheet.Cells[r + 2, 8].Formula = "=" + currentWorksheet.Cells[r, 7].Start.Address + "/'Average New Asst Prof Salary'!D2";
                        }
                    }
                    else {
                        if (currentWorksheet.Cells[r, 2].Value != currentWorksheet.Cells[r + 1, 2].Value
                            && currentWorksheet.Cells[r, 2].Value != null && currentWorksheet.Cells[r + 1, 2].Value != null)
                        {
                            currentWorksheet.InsertRow(r, 2, 2); //Insert two rows, one to be blank the other as a header.
                            currentWorksheet.Cells[r + 2, 2].Value = currentWorksheet.Cells[r + 3, 2].Value; //Set the header name for the section

                            //Insert the Formulae
                            currentWorksheet.Cells[r + 2, 4].Formula = "=QUARTILE(C:C,1)";
                            currentWorksheet.Cells[r + 2, 5].Formula = "=AVERAGE(C:C)";
                            currentWorksheet.Cells[r + 2, 6].Formula = "=QUARTILE(C:C,3)";
                            currentWorksheet.Cells[r + 2, 7].Formula = "=MEDIAN(C:C,1)";
                            currentWorksheet.Cells[r + 2, 8].Formula = "=" + currentWorksheet.Cells[r, 7].Start.Address + "/'Average New Asst Prof Salary'!D2";
                        }
                        
                    }
                  }
                    //Apply Excel formatting
                    currentWorksheet.Cells["D:G"].Style.Numberformat.Format = "$###,###,##0";
                    currentWorksheet.Cells["H2"].Style.Numberformat.Format = "#0.00";
                    currentWorksheet.Cells["A:Z"].AutoFitColumns();
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
