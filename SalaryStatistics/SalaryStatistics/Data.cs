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
        private float averageNewAssociateProfessorSalary = 50000;
        private float averageNewFullProfessorSalary = 100000;
        private Dictionary<string, int> headerColumns;
        private List<string> processedWorksheetNames = new List<string>();

        public Data(string path, float _constantD, float _constantK, float _constantL)//, float aNAPS, float aNFPS)
        {
            filePath = path;
            constantD = _constantD;
            constantK = _constantK;
            constantL = _constantL;
            excelFile = new ExcelPackage(new FileInfo(filePath));
            //averageNewAssociateProfessorSalary = aNAPS;
            //averageNewFullProfessorSalary = aNFPS;
        }
    }
}