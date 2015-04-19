using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Reflection;
using System.Windows.Forms;
using OfficeOpenXml;
namespace SalaryStatistics
{
    public class Data
    {
        private string filePath = "";
        private FileInfo inputFile;
        private ExcelPackage excelFile;

        //Uses the filePath set in the objects instantiation to load the file and create an ExcelPackage with it.
        public Data(string path)
        {
            filePath = path;
            inputFile = new FileInfo(filePath);
            excelFile = new ExcelPackage(inputFile);

            //TODO: Add error handling if the file can't be opened or it's not an excel file.
            Console.WriteLine("Data() ending.");
        }

        //Uses the searchForHeader's function to find the header row and delete everything above that
        //
        public void prepare()
       {
            //Get the first worksheet in the workbook
                ExcelWorksheet sourceData = excelFile.Workbook.Worksheets[1];
            //Create a new worksheet
                ExcelWorksheet preparedData = excelFile.Workbook.Worksheets["Prepared Data"];
            //Tracks the number of rows from sourceData that are ignored
                int ignoredRows = 0;
            //Gets the index of the header row in sourceData, everything above this is ignored
                int headerRow = 7; //getHeaderRow("Job Title");
            //The names of the columns we want in Prepared Data
                string[] keyColumnNames = {"Job Title"};
            //The indexes of where the desired columns are in sourceData
                Dictionary<string, int> keyColumns = getKeyColumns(sourceData, keyColumnNames, headerRow);
            //Keeps track of which row and column we're inserting to in the preparedData workshet
                int insertRow = 1;
                int insertCol = 1;

            //Loop through all the rows in sourceData
            for (int row = headerRow; row < sourceData.Dimension.End.Row; row++)
            {
                //Ignore rows without a Job Title value
                if (sourceData.Cells[row, keyColumns["Job Title"]] != null)
                {
                    insertCol = 1; //Reset insertions to the beginning of the row
                    //Copy each of the keyColumns over to the new new worksheet
                    foreach (KeyValuePair<string, int> column in keyColumns)
                    {
                        preparedData.Cells[insertRow, insertCol].Value = sourceData.Cells[row, column.Value];
                        insertCol++; //increment cells in the row
                    }
                    insertRow++; //Increment rows
                } else {
                    ignoredRows++; //Increment the number of ignored rows, that is those without a Job Title value
                }
            }
            Console.WriteLine("Ignored {0} rows.", ignoredRows);
            Console.WriteLine("prepare() ending.");
       } //End of prepare()

        public void close()
        {
            SaveFileDialog dialogue = new SaveFileDialog();
            using (FileStream newFile = new FileStream(filePath, FileMode.Create))
            {
                excelFile.SaveAs(newFile);
            }
            Console.WriteLine("close() ending.");
        }

        //Searches for the row with the specified string in the object's excelFile and returns its integer index.
        private int getHeaderRow(string header)
        {
            int row = 1;
            int col = 1;
            int numCols = 1;
            int numRows = 1;
            bool foundIt = false;
            ExcelWorksheet currentWorksheet = excelFile.Workbook.Worksheets[1];

            //count columns in document
            while (currentWorksheet.Cells[numRows, numCols].Value != null)
            {
                numCols++;
            }

            //search 1st row of columna for header
            while (foundIt != true)
            {
                for (col = 1; col < numCols; col++)
                {
                    if (currentWorksheet.Cells[row, col].Value.Equals(header))
                    {
                        foundIt = true;
                        Console.WriteLine("\tFound the header identification string '" + header + "' in column " + col);
                        return col;
                    }//end if
                }//end col for
            }//end while

            Console.WriteLine("getHeaderRow() returning.");
            return col;
        }


        //UNUSED
        private void populateWorksheet(string title)
        {
            ExcelWorksheet newWorksheet = excelFile.Workbook.Worksheets[title];
        }

        //Searches through the given headerRow for the strings in keyColumnNames in the worksheet
        // and returns a Dictionary with the names and column indexes.
        private Dictionary<string,int> getKeyColumns(ExcelWorksheet worksheet, string[] keyColumnNames, int headerRow)
        {
            Dictionary<string,int> keyColumns = new Dictionary<string,int>();

            foreach (string column in keyColumnNames)
            {
                for (int x = 1; x < worksheet.Dimension.End.Column; x++)
                {
                    if (worksheet.Cells[headerRow, x].Value == column)
                    {
                        keyColumns[column] = x;
                    }
                }
            }
            Console.WriteLine("getKeyColumns() returning.");
            return keyColumns;
        }

        //UNUSED: Sets the first row of the provided worksheet to the keys of the provided Dictionary and sets the row to bold
        private void setHeaderRows(ref ExcelWorksheet worksheet, Dictionary<string,int> keyColumns)
        {
            int currentColumn = 1;
            foreach (KeyValuePair<string, int> column in keyColumns)
            {
                worksheet.Cells[1, currentColumn].Value = column.Key;
                currentColumn++;
            }
            worksheet.Row(1).Style.Font.Bold = true;
            Console.WriteLine("setHeaderRows ending.");
        }
    }
}