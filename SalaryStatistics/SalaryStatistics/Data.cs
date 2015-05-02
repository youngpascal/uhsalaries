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
        private List<string> filters;
        private List<string> processedWorksheetNames = new List<string>();
        private List<string> listOfJobFilters = new List<string>();
        private List<string> listOfDepartmentFilters = new List<string>();
        private bool fitlered = false;
        private int jobFilterCount = 0;
        private int departmentFilterCount = 0;
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

        public ExcelPackage getExcelFile()
        {
            return excelFile;
        }

        public void setJobFilterCount(int count)
        {
            jobFilterCount = count;
        }

        public void setDepartmentFilterCount(int count)
        {
            departmentFilterCount = count;
        }
    }
}