﻿using OfficeOpenXml;
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
                Dictionary<string, int> totalSalary;
                Dictionary<string, int> proposedColumn;

                int headerRow = searchForHeaderRow(headerName, sourceWorksheet);
                //file.writeline("Headers on Row: {0}", headerRow);

                  //Find the integer indexes of of the desired cooluns (Job Title, Proposed Salary, and Department ID)
                  //Gets the column that the proposed salaries begin at
                    string[] proposedColumName = {"Proposed"};
                    proposedColumn = searchForHeaderColumns(sourceWorksheet, proposedColumName, headerRow-1, 1);
                    
                  //Gets the colun number of the proposed salary by using the offeset of the propsed column
                    string[] proposedSalaryColumnName = {"Total Salary"};
                    totalSalary = searchForHeaderColumns(sourceWorksheet, proposedSalaryColumnName, headerRow, proposedColumn["Proposed"]);
    
                  //Gets the column numbers of the Job Title and department id columns
                    string[] headerColumnNames = { "Job Title", "Pos Deptid" };
                    headerColumns = searchForHeaderColumns(sourceWorksheet, headerColumnNames, headerRow, 1);
    
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
            preparedWorksheet.Cells["A:Z"].AutoFitColumns();
            preparedWorksheet.Cells["C:C"].Style.Numberformat.Format = "$###,###,##0";

            //Add the constants sheet
            addConstantSheet();

            //excelFile.Workbook.Worksheets.Add("Summary");
        }

        public int searchForHeaderRow(string headerName, ExcelWorksheet currentWoksheet)
        {
            //Find all cells that match the query in the columsn A thorugh Z
            var query = (from cell in currentWoksheet.Cells["A:Z"] where cell.Value is string && (string)cell.Value==headerName select cell);

            //Return the row of the first cell found.
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
                        if (!foundColumns.ContainsKey(header))
                              foundColumns.Add(header, cell.Start.Column);
                    }
                }
            }

            return foundColumns;
        }

        public void copyInputOne()
        {
            ExcelWorksheet inputOneWorksheet = inputOnePackage.Workbook.Worksheets[1];
            ExcelWorksheet destinationWorksheet = excelFile.Workbook.Worksheets.Add("Average New Asst Prof Salary");

            int endRow = inputOneWorksheet.Dimension.End.Row;
            int endCol = inputOneWorksheet.Dimension.End.Column;

            for (int i = 1; i <= endRow; i++)
            {
                inputOneWorksheet.Cells[i, 1, i, endCol].Copy(destinationWorksheet.Cells[i, 1, i, endCol]);
            }

            destinationWorksheet.InsertColumn(endCol + 1,1);
            destinationWorksheet.Cells[1, endCol + 1].Value = "Adjusted New Associate Salaries";
            destinationWorksheet.Cells[1, endCol + 2].Value = "Adjusted Professor Salaries";

            for (int i = 2; i <= endRow; i++)
            {
                destinationWorksheet.Cells[i, endCol + 1].Formula = "((A" + i + "*(Constants!B3)^(Constants!B4-C" + i + ")+7000))*(Constants!B2^Constants!B1)";
                destinationWorksheet.Cells[i, endCol + 2].Formula = "(D" + i + "+10000)*(Constants!B2^Constants!B5)";
               // destinationWorksheet.Cells[i, endCol + 1].Formula = "A" + i + "*((Constants!B3";
                destinationWorksheet.Cells[i, endCol + 1, i, endCol + 2].Style.Numberformat.Format = "$###,###,##0";
            }

            destinationWorksheet.Cells["A:Z"].AutoFitColumns();
        }

        public void copyInputTwo()
        {
            ExcelWorksheet inputTwoWorksheet = inputTwoPackage.Workbook.Worksheets[1];
            ExcelWorksheet destinationWorksheet = excelFile.Workbook.Worksheets.Add("Tier 1 Data");

            int endRow = inputTwoWorksheet.Dimension.End.Row;
            int endCol = inputTwoWorksheet.Dimension.End.Column;

            for (int i = 1; i <= endRow; i++)
            {
                inputTwoWorksheet.Cells[i, 1, i, endCol].Copy(destinationWorksheet.Cells[i, 1, i, endCol]);
            }

            destinationWorksheet.Cells["A:Z"].AutoFitColumns();
        }

        public void copyInputThree()
        {
            ExcelWorksheet inputThreeWorksheet = inputThreePackage.Workbook.Worksheets[1];
            ExcelWorksheet destinationWorksheet = excelFile.Workbook.Worksheets.Add("UH Specialty Averages");

            int endRow = inputThreeWorksheet.Dimension.End.Row;
            int endCol = inputThreeWorksheet.Dimension.End.Column;

            for (int i = 1; i <= endRow; i++)
            {
                inputThreeWorksheet.Cells[i, 1, i, endCol].Copy(destinationWorksheet.Cells[i, 1, i, endCol]);
            }

            destinationWorksheet.Cells["A:Z"].AutoFitColumns();
            
        }

        public void fetchFilters(Dictionary<string, int> columns, ExcelWorksheet sourceWorksheet)
        {
            Dictionary<string, int> worksheetInsertionPoints = new Dictionary<string, int>();
            string cellValueSheetName = "";
            int endRow = sourceWorksheet.Dimension.End.Row;

            //Loop through all rows in a column
            for (int row = 2; row <= endRow; row++)
            {
                foreach (KeyValuePair<string, int> column in columns)
                {	//Get the value of the cells in the row and replace any '/' with '-'
                    cellValueSheetName = replaceSlash(sourceWorksheet.Cells[row, column.Value].Value.ToString());
                    char[] checkForFilterType = cellValueSheetName.ToArray();

                    if (checkForFilterType[0].Equals('H'))
                    {
                        if (!listOfDepartmentFilters.Contains(cellValueSheetName))
                        {
                            listOfDepartmentFilters.Add(cellValueSheetName);
                        }
                    }
                    else if (!listOfJobFilters.Contains(cellValueSheetName))
                    {
                        listOfJobFilters.Add(cellValueSheetName);
                    }
                }
            }
        }

        public List<string> getJobFiltersList()
            {
                return listOfJobFilters;
            }

        public List<string> getDepartmentFiltersList()
        {
            return listOfDepartmentFilters;
        }
    
        public List<string> getLists(string option)
        {
            //Sort the lists by alphabetical order
            jobTitles.Sort();
            departments.Sort();

            if (option.Equals("Job Title"))
            {
                return jobTitles;
            }
            else
            {
                return departments;
            }
        }

    }
}