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
                        break; //Skips all entrys in totalSalary after the first, I think -kj
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
                    

                    //Delete the empty rows in preparedWorksheet
                        int[] deletedRows = new int[1000];
                        int i = 0;
                    foreach (var cell in preparedWorksheet.Cells[2, 1, preparedWorksheet.Dimension.End.Row, headerColumns.Count])
                    {
                        if (cell.Value == null || cell.Value == "")
                        {
                            deletedRows[i] = cell.Start.Row;
                            //preparedWorksheet.DeleteRow(cell.Start.Row, 1, true);
                            preparedWorksheet.Cells[cell.Start.Row, cell.Start.Column + 1].Value = "<- Row Selected for Deletion";
                            i++;
                        }
                    }

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

        /*
        public void prepareData(string title1, string title2)
        {
            FileInfo existingPath = new FileInfo(filePath);
            using (ExcelPackage p = new ExcelPackage(existingPath))
            {
                ExcelWorksheet ws1 = p.Workbook.Worksheets[title1];
                ExcelWorksheet ws2 = p.Workbook.Worksheets[title2];
                ws1.Cells["B1:B1558"].Copy(ws2.Cells["A1:A1558"]);
                ws1.Cells["K1:K1558"].Copy(ws2.Cells["B1:B1558"]);
                ws1.Cells["AA1:AA1558"].Copy(ws2.Cells["C1:CC1558"]);
            }
        }

        public void sortPreparedData()
        {
            FileInfo existingPath = new FileInfo(filePath);
            using (ExcelPackage p = new ExcelPackage(existingPath))
            {
                ExcelWorksheet ws = p.Workbook.Worksheets["Prepared Data"];
                ExcelWorksheet ws2 = null;
                string temp = "", temp2 = "";
                int endCol = ws.Dimension.End.Column;
                int endRow = ws.Dimension.End.Row;
                for (int i = 2; i < endRow; i++)
                {
                    if (ws.Cells["A" + i].Value == null)
                        i++;
                    else
                    {
                        //Job title
                        temp = replaceSlash(ws.Cells["A" + i].Value.ToString());

                        if (p.Workbook.Worksheets[replaceSlash(temp)] == null)
                        {
                            ws2 = p.Workbook.Worksheets.Add(temp);
                            Console.WriteLine("\tAdding worksheet: " + temp);
                            int newCol = ws2.Dimension.End.Row + 1;
                            ws.Cells["A" + i].Copy(ws2.Cells["A" + newCol]);
                            ws.Cells["B" + i].Copy(ws2.Cells["B" + newCol]);
                            ws.Cells["C" + i].Copy(ws2.Cells["C" + newCol]);
                        }
                        else
                        {
                            ws2 = p.Workbook.Worksheets[temp];
                            int newCol = ws2.Dimension.End.Row + 1;
                            ws.Cells["A" + i].Copy(ws2.Cells["A" + newCol]);
                            ws.Cells["B" + i].Copy(ws2.Cells["B" + newCol]);
                            ws.Cells["C" + i].Copy(ws2.Cells["C" + newCol]);
                        }
                        //Department ID
                        temp2 = replaceSlash(ws.Cells["C" + i].Value.ToString());

                        if (p.Workbook.Worksheets[replaceSlash(temp)] == null)
                        {
                            p.Workbook.Worksheets.Add(temp);
                            ws2 = p.Workbook.Worksheets.Add(temp);
                            int newCol = ws2.Dimension.End.Row + 1;
                            ws.Cells["C" + i].Copy(ws2.Cells["A" + newCol]);
                            ws.Cells["B" + i].Copy(ws2.Cells["B" + newCol]);
                            ws.Cells["A" + i].Copy(ws2.Cells["C" + newCol]);
                            Console.WriteLine("\tAdding worksheet: " + temp);
                        }
                        else
                        {
                            ws2 = p.Workbook.Worksheets[temp];
                            int newCol = ws2.Dimension.End.Row + 1;
                            ws.Cells["C" + i].Copy(ws2.Cells["A" + newCol]);
                            ws.Cells["B" + i].Copy(ws2.Cells["B" + newCol]);
                            ws.Cells["A" + i].Copy(ws2.Cells["C" + newCol]);
                        }
                    }
                    //var sheet = p.Workbook.Worksheets[temp];



                }
            }
        }

        public void sortPreparedDeptData()
        {
            FileInfo existingPath = new FileInfo(filePath);

            using (ExcelPackage p = new ExcelPackage(existingPath))
            {
                ExcelWorksheet ws = p.Workbook.Worksheets["Prepared Data"];
                ExcelWorksheet ws2 = null;
                string temp = "";
                int endCol = ws.Dimension.End.Column;
                int endRow = ws.Dimension.End.Row;

                for (int i = 2; i < endRow; i++)
                {
                    if (ws.Cells["C" + i].Value == null)
                        i++;
                    else
                        temp = replaceSlash(ws.Cells["C" + i].Value.ToString());

                    //var sheet = p.Workbook.Worksheets[temp];

                    if (p.Workbook.Worksheets[replaceSlash(temp)] == null)
                    {
                        p.Workbook.Worksheets.Add(temp);
                        Console.WriteLine("\tAdding worksheet: " + temp);
                    }
                    else
                        ws2 = p.Workbook.Worksheets[temp];
                }
            }
        }

        public string replaceSlash(string s)
        {
            string pattern = "\\/";
            Regex regex = new Regex(pattern);
            return regex.Replace(s, "-");
        }
        */
    }
}