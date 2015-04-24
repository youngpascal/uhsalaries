using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SalaryStatistics
{
    public partial class Data
    {
        public void Prepare(string sourceSheetName, string preparedSheetName, string headerName)
        {       //Get the source worksheet in the workbook by name
                ExcelWorksheet sourceWorksheet = excelFile.Workbook.Worksheets[1];

                int headerRow = searchForHeaderRow(headerName, sourceWorksheet);
                Console.WriteLine("Headers on Row: {0}", headerRow);

                //Find the integer indexes of of the desired cooluns (Job Title, Proposed Salary, and Department ID)
                  //Gets the column that the proposed salaries begin at
                    string[] proposedColumName = {"Proposed"};
                    Dictionary<string, int> proposedColumn = searchForHeaderColumns(sourceWorksheet, proposedColumName, headerRow-1, 1);
                    
                  //Gets the colun number of the proposed salary by using the offeset of the propsed column
                    string[] proposedSalaryColumnName = {"Total Salary"};
                    Dictionary<string, int> totalSalary = searchForHeaderColumns(sourceWorksheet, proposedSalaryColumnName, headerRow, proposedColumn["Proposed"]);
    
                  //Gets the column numbers of the Job Title and department id columns
                    string[] headerColumnNames = { "Job Title", "Pos Deptid" };
                    Dictionary<string, int> headerColumns = searchForHeaderColumns(sourceWorksheet, headerColumnNames, headerRow, 1);
    
                    foreach (KeyValuePair<string, int> entry in totalSalary)
                    {
                        headerColumns.Add(entry.Key, entry.Value);
                        break; //Skips all entrys in totalSalary after the first, I think, just in case there's more than one returned. -kj
                    }
                
                //Add headers to the pareparedData worksheet then copy the three key columns of all the rows that don't have null job titles to it  
                  //Check to see if worksheet exists already
                    var preparedWorksheet = excelFile.Workbook.Worksheets[preparedSheetName];
                    if (preparedWorksheet == null)
                    {
                        preparedWorksheet = excelFile.Workbook.Worksheets.Add(preparedSheetName);
                    }
                    
                  //Copy the key columns (headers and values and empty rows)
                    int endRow = sourceWorksheet.Dimension.End.Row;
                    int tracker = 1;
                        foreach (KeyValuePair<string, int> column in headerColumns)
                        {
                            int col = column.Value;
                            sourceWorksheet.Cells[headerRow, col, endRow, col].Copy(preparedWorksheet.Cells[1, tracker, endRow, tracker]);
                            tracker++;
                        }
                    

                  //Find all the rows for deletion in the preparedData worksheet
                    int[] deletedRows = new int[1000];
                    int i = 0;
                    foreach (var cell in preparedWorksheet.Cells[2, 1, preparedWorksheet.Dimension.End.Row, headerColumns.Count])
                    {
                        if (cell.Value == null || cell.Value == "")
                        {
                            deletedRows[i] = cell.Start.Row;
                            i++;
                        }
                    }
                  //Delete the selected rows from the pareparedData worksheet
                    int offset = 0;
                    for (int x = 0; x < deletedRows.Length; x++)
                    {
                        if (deletedRows[x] != 0 && deletedRows[x] != null)
                        {
                            preparedWorksheet.DeleteRow(deletedRows[x] - offset, 1, true);
                            offset++;
                        }
                    }
        }

        private int searchForHeaderRow(string headerName, ExcelWorksheet currentWoksheet)
        {
            //Find all cells that match the query in the columsn A thorugh Z
            var query = (from cell in currentWoksheet.Cells["A:Z"] where cell.Value is string && (string)cell.Value==headerName select cell);

            //Returnt the row of the first cell found.
            foreach (var cell in query)
            {
                return cell.Start.Row;
            }

            return 0;
        }

        private Dictionary<string, int> searchForHeaderColumns(ExcelWorksheet currentWorksheet, string[] headers, int headerRow, int columnOffset) {
            Dictionary<string, int> foundColumns = new Dictionary<string, int> { };

            //Examine each cell in the given row after the given offset
            foreach (var cell in currentWorksheet.Cells.Offset(0,columnOffset))
            {   //Compare each cell's value to the desired column values
                foreach (string header in headers) {
                    if (cell.Value != null && cell.Value.Equals(header))
                    {
                        foundColumns.Add(header, cell.Start.Column);
                    }
                }
            }

            return foundColumns;
        }
    }
}