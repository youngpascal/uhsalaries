using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Reflection;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Text.RegularExpressions;
namespace SalaryStatistics
{
    public class Data
    {
        private string filePath = "";

        public Data(string path)
        {
            filePath = path;
        }

        public void load(int column)
        {
            FileInfo existingFile = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                // get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                int deletedRows = 0;

                while (column != worksheet.Dimension.End.Column)
                {
                    for (int row = 1; row < 51; row++)
                    {
                        if (worksheet.Cells[row, column].Value == null)
                        {
                            worksheet.DeleteRow(row, 1, true);
                            deletedRows++;
                            row--;
                        }
                        else if (deletedRows > 15)
                        {
                            column++;
                            deletedRows = 0;
                        }
                    }
                }//end while
                for (int row = 1; row < 10; row++)
                    Console.WriteLine("\tCell({0},{1}).Value={2}", row, column, worksheet.Cells[row, column].Value);


                //Add the found headers as new worksheets
                //Check to see if worksheet exists already
                var sheet = package.Workbook.Worksheets[worksheet.Cells[1, column].Value.ToString()];
                if (sheet == null)
                {
                    var ws = worksheet.Workbook.Worksheets.Add(worksheet.Cells[1, column].Value.ToString());
                    SaveFileDialog sfd = new SaveFileDialog();
                    using (FileStream fs = new FileStream(filePath, FileMode.Create))
                    {
                        package.SaveAs(fs);
                    }
                }
                Console.WriteLine("\tDeleted Rows: {0}", deletedRows);
            } // the using statement automatically calls Dispose() which closes the package.
        }

        public int searchForHeader(string s)
        {
            FileInfo existingFile = new FileInfo(filePath);
            bool f = false;

            using (ExcelPackage p = new ExcelPackage(existingFile))
            {
                ExcelWorksheet ws = p.Workbook.Worksheets[1]; //first worksheet
                int row = 1, col = 1;
                int numCols = ws.Dimension.End.Column;
                int numRows = ws.Dimension.End.Row;
                //search 1st row of columna for header
                while (f != true)
                {
                    try
                    {
                        for (col = 1; col < numCols; col++)
                        {
                            for (row = 1; row < numRows; row++)
                            {
                                //Console.WriteLine("Currently in column: " + col);
                                if (ws.Cells[row, col].Value == null)
                                {
                                    col++;
                                }
                                else if (ws.Cells[row, col].Value.Equals(s))
                                {
                                    f = true;
                                    Console.WriteLine("\tfound in column: " + col);
                                    return col;
                                }//end if
                            }
                        }//end col for
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine(e.ToString());
                        continue;
                    }
                }//end while
                return col;
            }//end using
        }//end searchForHeader
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

                using (FileStream fs = new FileStream(filePath, FileMode.Create))
                {
                    p.SaveAs(fs);
                }
            }
        }

        public void sortPreparedData()
        {
            FileInfo existingPath = new FileInfo(filePath);
            using (ExcelPackage p = new ExcelPackage(existingPath))
            {
                ExcelWorksheet ws = p.Workbook.Worksheets["Prepared Data"];
                ExcelWorksheet ws2 = null;
                string temp = "", temp2="";
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
                    using (FileStream fs = new FileStream(filePath, FileMode.Create))
                    {
                        p.SaveAs(fs);
                    }
            }
        }

        public void sortPreparedDeptData()
        {
            FileInfo existingPath = new FileInfo(filePath);

            using(ExcelPackage p = new ExcelPackage(existingPath))
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
                using (FileStream fs = new FileStream(filePath, FileMode.Create))
                {
                    p.SaveAs(fs);
                }
            }
        }

        public string replaceSlash(string s)
        {
            string pattern = "\\/";
            Regex regex = new Regex(pattern);
            return regex.Replace(s, "-");
        }

    }
}