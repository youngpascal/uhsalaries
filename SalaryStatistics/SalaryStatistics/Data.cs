using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace SalaryStatistics
{
    public partial class Data
    {
        private string filePath = "";
        private ExcelPackage excelFile;
        private float constantD;
        private float constantK;
        private float constantL;
        private Dictionary<string, int> headerColumns;

        public Data(string path, float _constantD, float _constantK, float _constantL)
        {
            filePath = path;
            constantD = _constantD;
            constantK = _constantK;
            constantL = _constantL;
            excelFile = new ExcelPackage(new FileInfo(filePath));
        }
    }
}