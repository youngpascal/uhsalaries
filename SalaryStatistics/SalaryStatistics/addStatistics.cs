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
            int numberOfThatJob;
            int endRow;
            int statInsertionRow;
            int numberOfRows;

            //Loop through each worksheet we've created
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

                endRow = currentWorksheet.Dimension.End.Row;
                statInsertionRow = 3; //Header row is 1, All row is 2

                //Loop through all rows in the worksheet
                for (int r = 4; r < endRow; )
                {
                    //Keep track of how many rows have a given value
                    numberOfRows = 1;

                    //If it's a department sheet compare based on job title, if it's a job title compare basedon department
                    if (currentWorksheet.Name[0] == 'H' || currentWorksheet.Name[0] == 'h')

                    {   //Loop downwards to count how many rows have the given value
                        while (currentWorksheet.Cells[r, 1].Value == currentWorksheet.Cells[r + numberOfRows, 1].Value)
                        {
                            numberOfRows++;
                        }
                        currentWorksheet.InsertRow(statInsertionRow, 1, statInsertionRow - 1);
                        r++; //Incremnt r for the offset from inserting the row

                        //Set the Title for This Row
                        currentWorksheet.Cells[statInsertionRow, 1].Value = currentWorksheet.Cells[r, 1].Value;

                        //Insert the Formulae
                        currentWorksheet.Cells[statInsertionRow, 4].Formula = "QUARTILE(" + currentWorksheet.Cells[r, 3].Address + ":" + currentWorksheet.Cells[r + numberOfRows - 1, 3].Address + ",1)";
                        currentWorksheet.Cells[statInsertionRow, 5].Formula = "AVERAGE(" + currentWorksheet.Cells[r, 3].Address + ":" + currentWorksheet.Cells[r + numberOfRows - 1, 3].Address + ")";
                        currentWorksheet.Cells[statInsertionRow, 6].Formula = "QUARTILE(" + currentWorksheet.Cells[r, 3].Address + ":" + currentWorksheet.Cells[r + numberOfRows - 1, 3].Address + ",3)";
                        currentWorksheet.Cells[statInsertionRow, 7].Formula = "MEDIAN(" + currentWorksheet.Cells[r, 3].Address + ":" + currentWorksheet.Cells[r + numberOfRows - 1, 3].Address + ",1)";

                        var query = (from cell in excelFile.Workbook.Worksheets["Average New Asst Prof Salary"].Cells["B:B"] where cell.Value is string && (string)cell.Value == currentWorksheet.Name select cell);

                                                if (query == null)
                        {
                            currentWorksheet.Cells[statInsertionRow, 8].Formula = currentWorksheet.Cells[statInsertionRow, 7].Start.Address + "/" + query.First().FullAddress;
                        }
                        else
                        {
                            currentWorksheet.Cells[statInsertionRow, 8].Value = "Missing Assist. Average";
                        }
                        statInsertionRow++;
                    }
                    else
                    {   //Loop downwards to count how many rows have the given value
                        while (currentWorksheet.Cells[r, 2].Value == currentWorksheet.Cells[r + numberOfRows, 2].Value)
                        {
                            numberOfRows++;
                        }

                        currentWorksheet.InsertRow(statInsertionRow, 1, statInsertionRow - 1);
                        r++; //Incremnt r for the offset from inserting the row

                        //Set the Title for This Row
                        currentWorksheet.Cells[statInsertionRow, 2].Value = currentWorksheet.Cells[r, 2].Value;

                        //Insert the Formulae
                        currentWorksheet.Cells[statInsertionRow, 4].Formula = "QUARTILE(" + currentWorksheet.Cells[r, 3].Address + ":" + currentWorksheet.Cells[r + numberOfRows - 1, 3].Address + ",1)";
                        currentWorksheet.Cells[statInsertionRow, 5].Formula = "AVERAGE(" + currentWorksheet.Cells[r, 3].Address + ":" + currentWorksheet.Cells[r + numberOfRows - 1, 3].Address + ")";
                        currentWorksheet.Cells[statInsertionRow, 6].Formula = "QUARTILE(" + currentWorksheet.Cells[r, 3].Address + ":" + currentWorksheet.Cells[r + numberOfRows - 1, 3].Address + ",3)";
                        currentWorksheet.Cells[statInsertionRow, 7].Formula = "MEDIAN(" + currentWorksheet.Cells[r, 3].Address + ":" + currentWorksheet.Cells[r + numberOfRows - 1, 3].Address + ",1)";

                        var query = (from cell in excelFile.Workbook.Worksheets["Average New Asst Prof Salary"].Cells["B:B"]
                                     where cell.Value is string && String.Equals((string)cell.Value, (string)currentWorksheet.Cells[r, 2].Value, StringComparison.OrdinalIgnoreCase)
                                     select cell);

                        if (query.GetEnumerator().MoveNext())
                        {
                            currentWorksheet.Cells[statInsertionRow, 8].Formula = currentWorksheet.Cells[statInsertionRow, 7].Start.Address + "/" + query.First().FullAddress;
                        }
                        else
                        {
                            currentWorksheet.Cells[statInsertionRow, 8].Value = "Missing Assist. Average";
                        }
                            statInsertionRow++;
                    }

                    r = r + numberOfRows + 1; //Increment the row to the next job title 
                }

                //Insert the worksheet wide statistics
                currentWorksheet.Cells[2, 1].Value = "All";
                currentWorksheet.Cells[2, 4].Formula = "QUARTILE(C2:C2000,1)";
                currentWorksheet.Cells[2, 5].Formula = "AVERAGE(C2:C2000)";
                currentWorksheet.Cells[2, 6].Formula = "QUARTILE(C2:C2000,3)";
                currentWorksheet.Cells[2, 7].Formula = "MEDIAN(C2:C2000,1)";
                currentWorksheet.Cells[2, 8].Formula = "=G2/'Average New Asst Prof Salary'!D2";

                //Apply Excel formatting
                currentWorksheet.Row(2).Style.Font.Bold = true;
                currentWorksheet.Row(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                currentWorksheet.Column(8).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                currentWorksheet.Cells["D:G"].Style.Numberformat.Format = "$###,###,##0";
                currentWorksheet.Cells["H2:H" + statInsertionRow].Style.Numberformat.Format = "#0.00";
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
