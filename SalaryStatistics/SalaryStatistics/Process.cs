using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
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
                }
            }

            
            populateWorksheets(worksheetInsertionPoints, sourceWorksheet);
            Console.WriteLine("Added {0} worksheets", worksheetsAdded);
        }

        private string replaceSlash(string s)
        {
            string pattern = "\\/";
            Regex regex = new Regex(pattern);
            return regex.Replace(s, "-");
        }

        public void populateWorksheets(Dictionary<string, int> worksheet, ExcelWorksheet sourceWorksheet)
        {
            foreach (KeyValuePair<string,int> worksheetName in worksheet)
            {
                ExcelWorksheet currentWorksheet = excelFile.Workbook.Worksheets[worksheetName.Key];

                //Search sourceworksheet for all cells containing the key
                var query = (from cell in sourceWorksheet.Cells["A:C"] where cell.Value is string && (string)cell.Value == worksheetName.Key select cell);

                int startKey = worksheet[worksheetName.Key];
                //Returns the row of the first cell found.
                foreach (var cell in query)
                {
                    int rowToCopy = cell.Start.Row;

                    //Copy each row in each column
                    for (int i = 1; i <= headerColumns.Count; i++)
                    {
                        sourceWorksheet.Cells[rowToCopy, i].Copy(currentWorksheet.Cells[startKey, i]);
                    }

                    startKey++;
                }
            }

            //Go back through and add the headers to every worksheet
            int tracker = 1; //track the column
            foreach (KeyValuePair<string, int> header in headerColumns)
            {
                foreach (KeyValuePair<string, int> worksheetName in worksheet)
                {
                    ExcelWorksheet currentWorksheet = excelFile.Workbook.Worksheets[worksheetName.Key];
                    currentWorksheet.Cells[1, tracker].Value = header.Key;
                }
                tracker++;
            }
        }       
    }
}
