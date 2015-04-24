using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
namespace SalaryStatistics
{
    public partial class Data
    {
        public void Process()
        {
            ExcelWorksheet preparedWorksheet = excelFile.Workbook.Worksheets["Prepared Data"];
            string[] keyHeaders = {"Job Title", "Pos Deptid"};
            Dictionary<string, int> keyColumns = searchForHeaderColumns(preparedWorksheet, keyHeaders, 1, 0);
            createAndPopulateWorksheets(keyColumns, preparedWorksheet);
        }

        private void createAndPopulateWorksheets(Dictionary<string, int> columns, ExcelWorksheet sourceWorksheet)
        {
            ExcelWorksheet destinationWorksheet = null;
            Dictionary<string, int> worksheetInsertionPoints = new Dictionary<string,int>();
            string cellValueSheetName = "";
            int endRow = sourceWorksheet.Dimension.End.Row;
            int worksheetsAdded = 0;
            int insertionPoint;

            //Loop through all rows in a column
            for (int row = 2; row <= endRow; row++)
            {
                foreach (KeyValuePair<string, int> column in columns)
                {	//Get the value of the cells in the row and replace any '/' with '-'
                    cellValueSheetName = replaceSlash(sourceWorksheet.Cells[row, column.Value].Value.ToString());

                    //Create a new sheet if the sheet dosen't already exist for this value
                    if (excelFile.Workbook.Worksheets[cellValueSheetName] == null)
                    {   //Create the worksheet
                        destinationWorksheet = excelFile.Workbook.Worksheets.Add(cellValueSheetName);
                        worksheetInsertionPoints.Add(cellValueSheetName, 2);
                        worksheetsAdded++;
                        Console.WriteLine("\tAdding worksheet: " + cellValueSheetName);
                        //Add headers to the worksheet
                        foreach (KeyValuePair<string, int> header in headerColumns)
                        {
                            destinationWorksheet.Cells[1, header.Value].Value = header.Key;
                        }
                    }

                    //Add the row to the correct worksheet at the corret line
                    insertionPoint = worksheetInsertionPoints[cellValueSheetName];
                    sourceWorksheet.Cells[row, 1, row, headerColumns.Count].Copy(destinationWorksheet.Cells[insertionPoint, 1, insertionPoint, headerColumns.Count]);
                    worksheetInsertionPoints[cellValueSheetName]++;
                }
            }
            Console.WriteLine("Added {0} worksheets", worksheetsAdded);
        }

        private string replaceSlash(string s)
        {
            string pattern = "\\/";
            Regex regex = new Regex(pattern);
            return regex.Replace(s, "-");
        }
    }
}
