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
            List<cell> cells = new List<cell>()
            {
                new cell(1, 0, "All"),
                new cell(2, 0, "Associate Professor"),
                new cell(3, 0, "Full Professor"),

                new cell(0, 1, "Mean"),
                new cell(1, 1, "=AVERAGE(C:C)"),
                new cell(2, 1, "=AVERAGE(C:C)"),
                new cell(3, 1, "=AVERAGE(C:C)"),

                new cell(0, 2, "Median"),
                new cell(1, 2, "=MEDIAN(C:C)"),
                new cell(2, 2, "=MEDIAN(C:C)"),
                new cell(3, 2, "=MEDIAN(C:C)"),

                new cell(0, 3, "Compression"),
                new cell(2, 3, "=G6/Constants!E4"),
                new cell(3, 3, "=G7/Constants!E5")
            };

            foreach (string worksheetName in processedWorksheetNames)
            {
                currentWorksheet = excelFile.Workbook.Worksheets[worksheetName];
                foreach (var currentCell in cells)
                    {
                        valueString = currentCell.value.ToCharArray();
                        if (valueString[0] == '=')
                        {
                            currentWorksheet.Cells[4 + currentCell.y, headerColumns.Count + 2 + currentCell.x].Formula = currentCell.value;
                        }
                        else
                        {
                            currentWorksheet.Cells[4 + currentCell.y, headerColumns.Count + 2 + currentCell.x].Value = currentCell.value;
                        }
                        currentWorksheet.Cells[4 + currentCell.y, headerColumns.Count + 2 + currentCell.x].AutoFitColumns();
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