﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.Drawing;

namespace SalaryStatistics
{
    public partial class Data
    {

        Dictionary<string, int> keyColumns;

        public void Process(List<string> filters, bool isFiltered)
        {
            ExcelWorksheet preparedWorksheet = excelFile.Workbook.Worksheets["Prepared Data"];
           // string[] keyHeaders = {"Job Title", "Pos Deptid"};
           // keyColumns = searchForHeaderColumns(preparedWorksheet, keyHeaders, 1, 0);
            
            List<string> searchFilters = filters;
            createWorksheets(getKeyColumns(), filters, preparedWorksheet, isFiltered);


            //Seperate all worksheet names by department and job title
            foreach (string name in processedWorksheetNames)
            {
                if (name[0].Equals('H') || name[0].Equals('h'))
                {
                    departments.Add(name);
                }
                else
                {
                    jobTitles.Add(name);
                }
            }

            addStatistics();
        }

    //Supporting Functions

        private void createWorksheets(Dictionary<string, int> columns, List<string> filteredColumns, ExcelWorksheet sourceWorksheet, bool isFiltered)
        {
            ExcelWorksheet destinationWorksheet = null;
            Dictionary<string, int> worksheetInsertionPoints = new Dictionary<string,int>();
            string cellValueSheetName = "";
            int endRow = sourceWorksheet.Dimension.End.Row;
            int worksheetsAdded = 0;
            string[] filters = filteredColumns.ToArray();
            int insertionPoint;
            
            //Loop through all rows in a column
         if (isFiltered)
          {
            for (int row = 2; row <= endRow; row++)
            {
                foreach (KeyValuePair<string, int> column in columns)
                {	//Get the value of the cells in the row and replace any '/' with '-'
                    cellValueSheetName = replaceSlash(sourceWorksheet.Cells[row, column.Value].Value.ToString());
                    //Create a new sheet if the sheet dosen't already exist for this value
                    foreach (string filter in filters)
                    {
                        if (excelFile.Workbook.Worksheets[cellValueSheetName] == null && cellValueSheetName.Equals(filter))
                        {   //Create the worksheet
                            destinationWorksheet = excelFile.Workbook.Worksheets.Add(cellValueSheetName);
                            worksheetInsertionPoints.Add(cellValueSheetName, 2);
                            processedWorksheetNames.Add(cellValueSheetName);
                            worksheetsAdded++;
                            //file.writeline("\tAdding worksheet: " + cellValueSheetName);
                        }
                    }
                }//end column foreach             
            }
            destinationWorksheet = excelFile.Workbook.Worksheets.Add("Filters");
            processedWorksheetNames.Add("Filters");
            worksheetInsertionPoints.Add("Filters", 2);
            worksheetsAdded++;
        }
         else
         {
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
                             processedWorksheetNames.Add(cellValueSheetName);
                             worksheetsAdded++;
                             //file.writeline("\tAdding worksheet: " + cellValueSheetName);
                         }
                 }//end column foreach
             }
         }

            
            populateWorksheets(worksheetInsertionPoints, sourceWorksheet, isFiltered, filters);
            //file.writeline("Added {0} worksheets", worksheetsAdded);

        }

        private string replaceSlash(string s)
        {
            string pattern = "\\/";
            Regex regex = new Regex(pattern);
            return regex.Replace(s, "-");
        }

        //**********************Populate the worksheets with their respected data***************************************************//
        public void populateWorksheets(Dictionary<string, int> worksheet, ExcelWorksheet sourceWorksheet, bool isFiltered, string[] filters)
        {
            ExcelWorksheet currentWorksheet;

            //***********Copy all rows based on the name of the current worksheet, find the data in 'Prepared Data'**********************//
            foreach (KeyValuePair<string,int> worksheetName in worksheet)
            {
                currentWorksheet = excelFile.Workbook.Worksheets[worksheetName.Key];

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

                currentWorksheet.Cells["I1"].Hyperlink = new ExcelHyperLink("'Table Of Contents'!A1", "Back to Table Of Contents");
                currentWorksheet.Cells["I1"].Style.Font.UnderLine = true;
                currentWorksheet.Cells["I1"].Style.Font.Color.SetColor(Color.Blue);
                currentWorksheet.Cells["A:Z"].AutoFitColumns();

                currentWorksheet.Cells["A:Z"].AutoFitColumns();
            }
            //*************************************End row copy********************************************************//

            string filterSheetName = "";
            int filterSize = 0;
            ExcelWorksheet filteredSheet = excelFile.Workbook.Worksheets["Filters"];

            //Find out how big to make applicableFilters
            if (jobFilterCount > 1)
                filterSize = jobFilterCount;
            else
                filterSize = departmentFilterCount;

            string[] applicableFilters = new string[filterSize+1];

            //******************Handle filter cases****************************//
            if (isFiltered)
            {
                //If it is filtered by job title
                if (jobFilterCount > 1)
                {
                    int x = 0;
                    //Find the department that is being filtered
                    foreach (string filter in filters)
                    {
                        char[] checkForDepartmentFitler = filter.ToCharArray();


                        if (checkForDepartmentFitler[0] == 'H' || checkForDepartmentFitler[0] == 'h')
                        {
                            filterSheetName = filter;
                        }
                        else
                        {
                            applicableFilters[x] = filter;
                        }

                        x++;
                    }
                  

                    //Go to that worksheet
                    currentWorksheet = excelFile.Workbook.Worksheets[filterSheetName];
                    //Job title is first column
                    int col = 1;
                    int rowTracker = 1;
                    
                    foreach (string filter in applicableFilters)
                    {
                      for (int i = 1; i < currentWorksheet.Dimension.End.Row; i++)
                        {
                            if (currentWorksheet.Cells[i, col].Value == filter)
                            {
                                currentWorksheet.Cells[i, col, i, currentWorksheet.Dimension.End.Column].Copy(filteredSheet.Cells[rowTracker, col, rowTracker, currentWorksheet.Dimension.End.Column]);
                                rowTracker++;
                            }
                        }
                    }
                }
                else if (departmentFilterCount > 1)
                {
                    int x = 0;
                    //Find the department that is being filtered
                    foreach (string filter in filters)
                    {
                        char[] checkForDepartmentFitler = filter.ToCharArray();


                        if (checkForDepartmentFitler[0] == 'H' || checkForDepartmentFitler[0] == 'h')
                        {
                            applicableFilters[x] = filter;
                        }
                        else
                        {
                            filterSheetName = filter;
                        }

                        x++;
                    }


                    //Go to that worksheet
                    currentWorksheet = excelFile.Workbook.Worksheets[filterSheetName];
                    //file.writeline("Operating from worksheet {0}", filterSheetName);

                    //department is second column
                    int col = 2;
                    int rowTracker = 1;

                    foreach (string filter in applicableFilters)
                    {
                        for (int i = 1; i < currentWorksheet.Dimension.End.Row; i++)
                        {
                            if (currentWorksheet.Cells[i, col].Value == filter)
                            {
                                currentWorksheet.Cells[i, col-1, i, currentWorksheet.Dimension.End.Column].Copy(filteredSheet.Cells[rowTracker, col-1, rowTracker, currentWorksheet.Dimension.End.Column]);
                                rowTracker++;
                            }
                        }
                    }
                }
            }//***************************End filter case*******************************//

            //*******Go back through and add the headers to every worksheet***//
            int tracker = 1; //track the column
            foreach (KeyValuePair<string, int> header in headerColumns)
            {
                foreach (KeyValuePair<string, int> worksheetName in worksheet)
                {
                    currentWorksheet = excelFile.Workbook.Worksheets[worksheetName.Key];
                    currentWorksheet.Cells[1, tracker].Value = header.Key;
                    currentWorksheet.Cells["A:Z"].AutoFitColumns();
                    currentWorksheet.Cells["C:C"].Style.Numberformat.Format = "$###,###,##0";

                    if (worksheetName.Key[0] == 'H' || worksheetName.Key[0] == 'h')
                    {
                        sortWorksheet(currentWorksheet, "Job Title");
                    }
                    else
                    {
                        sortWorksheet(currentWorksheet, "Department");
                    }
                }
                tracker++;
            }
            //*******************End adding headers******************//
        }       
        //*********************End populate worksheets******************************************************************************//

        //******************Fetch key columns**************//
        public Dictionary<string, int> getKeyColumns()
        {
            ExcelWorksheet preparedWorksheet = excelFile.Workbook.Worksheets["Prepared Data"];
            string[] keyHeaders = { "Job Title", "Pos Deptid" };
            keyColumns = searchForHeaderColumns(preparedWorksheet, keyHeaders, 1, 0);
            return keyColumns;
        }
        //*****************End fetch key columns**********//

        //*************Get key columns based on selected filters****************//
        private Dictionary<string, int> getFilteredKeyColumns(List<string> filters)
        {
            ExcelWorksheet preparedWorksheet = excelFile.Workbook.Worksheets["Prepared Data"];
            string[] keyHeaders = filters.ToArray();
            keyColumns = searchForHeaderColumns(preparedWorksheet, keyHeaders, 1, 0);
            return keyColumns;
        }
        //*****************End fetch filtered key columns***********************//


        private void sortWorksheet(ExcelWorksheet currentWorksheet, string sortBy)
        {// This function assumes the first row is headers and dosen't sort it.
            List<row> rows = new List<row>();
            int endRow = currentWorksheet.Dimension.End.Row;

            //Create the list of rows
            for (int r = 2; r <= endRow; r++)
            {
                rows.Add(new row((string)currentWorksheet.Cells[r, 1].Value, (string)currentWorksheet.Cells[r, 2].Value, (double)currentWorksheet.Cells[r, 3].Value));
            }

            bool found;

            if (!currentWorksheet.Name[0].Equals('H') && !currentWorksheet.Name[0].Equals('h'))
            {
                foreach (string worksheetName in processedWorksheetNames)
                {
                    found = false;
                    if (worksheetName[0].Equals('H') || worksheetName[0].Equals('h'))
                    {
                        foreach (row aRow in rows)
                        {
                            if (aRow.departmentID.Equals(worksheetName))
                            {
                                found = true;
                                break;
                            }
                        }
                        if (!found)
                        {
                            rows.Add(new row("none", worksheetName, 0.0));
                        }
                    }
                }
            }

            //Sort the rows in the list
            if (sortBy == "Job Title")
            {
                rows.Sort((s1, s2) => s1.jobTitle.CompareTo(s2.jobTitle));
            }
            else if (sortBy == "Department")
            {
                rows.Sort((s1, s2) => s1.departmentID.CompareTo(s2.departmentID));
            }

            //Overwrite the sorted rows back into the worksheet
            for (int r=2; r<=endRow; r++)
            {
                currentWorksheet.Cells[r, 1].Value = rows[r - 2].jobTitle;
                currentWorksheet.Cells[r, 2].Value = rows[r - 2].departmentID;
                currentWorksheet.Cells[r, 3].Value = rows[r - 2].salary;
            }
        }

        struct row
        {
            public string jobTitle;
            public string departmentID;
            public double salary;

            public row(string j, string d, double s)
            {
                jobTitle = j;
                departmentID = d;
                salary = s;
            }
        }
    }
}
