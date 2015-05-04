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
            string[] statisticHeaders = { "1st Quartile", "Mean", "3rd Quartile", "Median", "Compression" };
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
                currentWorksheet.Cells[2, 8].Formula = "[Compression Formula]";

                //If it's not a department worksheet, skip the rest
                if (worksheetName[0] == 'H')
                {

                    for (int r = 3; r < currentWorksheet.Dimension.End.Row; r++)
                    {
                        if (!jobTitles.Contains(currentWorksheet.Cells[r, 1].Value) && currentWorksheet.Cells[r, 1].Value != null)
                        {
                            currentJobTitle = (string)currentWorksheet.Cells[r, 1].Value;
                            jobTitles.Add(currentJobTitle);
                        }
                    }
                    jobTitles.Sort();

                    foreach (string jobTitle in jobTitles)
                    {
                        //Select all cells with a the given jobTitle
                        var cellSortingQuery = (from cell in currentWorksheet.Cells["A:A"] where cell.Value is string && cell.Value.Equals(jobTitle) select cell);
                        var selectedCells = cellSortingQuery.ToArray();
                        numberOfThatJob = selectedCells.Count();

                        //Insert the header and statistics for that jobTitle section
                        currentWorksheet.InsertRow(bottomOfSortedRows, 1);
                        currentWorksheet.Cells[bottomOfSortedRows, 1].Value = jobTitle;
                        currentWorksheet.Cells[bottomOfSortedRows, 5].Formula = "=QUARTILE(C" + (bottomOfSortedRows + 1) + ":C" + (bottomOfSortedRows + numberOfThatJob) + ",1)";
                        currentWorksheet.Cells[bottomOfSortedRows, 6].Formula = "=AVERAGE((C" + (bottomOfSortedRows + 1) + ":C" + (bottomOfSortedRows + numberOfThatJob) + ")";
                        currentWorksheet.Cells[bottomOfSortedRows, 7].Formula = "=QUARTILE(C" + (bottomOfSortedRows + 1) + ":C" + (bottomOfSortedRows + numberOfThatJob) + "3)";
                        currentWorksheet.Cells[bottomOfSortedRows, 8].Formula = "=MEDIAN(C" + (bottomOfSortedRows + 1) + ":C" + (bottomOfSortedRows + numberOfThatJob) + ")";
                        currentWorksheet.Cells[bottomOfSortedRows, 9].Formula = "[Compression Formula]";

                        //Copy the row of each cell into the proper location.
                        foreach (var selectedCell in selectedCells)
                        {
                            currentWorksheet.InsertRow(bottomOfSortedRows, 1);
                            bottomOfSortedRows++;
                            currentWorksheet.Cells[selectedCell.Start.Row, 1, selectedCell.Start.Row, 3].Copy(currentWorksheet.Cells[bottomOfSortedRows, 1, bottomOfSortedRows, 3]);
                        }

                        currentWorksheet.InsertRow(bottomOfSortedRows, 1);
                        bottomOfSortedRows++;
                    }

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
}