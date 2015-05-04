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
        public void addConstantSheet()
        {   //Uses private excelFile variable as part of the Data class
            ExcelWorksheet constantSheet = excelFile.Workbook.Worksheets.Add("Constants");
            char[] valueString;
            Dictionary<string, string> cellsToInsert = new Dictionary<string, string>()
            {
                {"A1", "\"l\" Constant"},
                {"B1", constantL.ToString()},
                {"A2", "\"d\" Constant"},
                {"B2", constantD.ToString()},
                {"A3", "\"k\" Inflation Constant"},
                {"B3", constantK.ToString()}
            };

            //Inserts the constants on cells A1 thorugh B3
            foreach (KeyValuePair<string, string> cell in cellsToInsert)
            {
                valueString = cell.Value.ToCharArray();
                if (valueString[0] == '=')
                {
                    constantSheet.Cells[cell.Key].Formula = cell.Value;
                }
                else
                {
                    constantSheet.Cells[cell.Key].Value = cell.Value;
                }
                constantSheet.Cells[cell.Key].AutoFitColumns();
            }
        }
    }
}